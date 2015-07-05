Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = "M:RPGe Server <IP " & frmServer.Socket(0).LocalIP & " Port " & str(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If IsConnected(index) And player(index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Function IsLoggedIn(ByVal index As Long) As Boolean
    If IsConnected(index) And Trim(player(index).Login) <> "" Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim i As Long

    IsMultiAccounts = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And LCase(Trim(player(i).Login)) = LCase(Trim(Login)) Then
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
        If IsConnected(i) And Trim(GetPlayerIP(i)) = Trim(IP) Then
            n = n + 1
            
            If (n > 1) Then
                IsMultiIPOnline = True
                Exit Function
            End If
        End If
    Next i
End Function

Function IsBanned(ByVal IP As String) As Boolean
Dim filename As String, fIP As String, fName As String
Dim f As Long

    IsBanned = False
    
    filename = App.Path & "\banlist.txt"
    
    ' Check if file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If
    
    f = FreeFile
    Open filename For Input As #f
    
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



Sub SendDataTo(ByVal index As Long, ByVal Data As String)
Dim bufferLog As String
        bufferLog$ = Replace(Data$, END_CHAR, "||")
        bufferLog$ = Replace(Data$, SEP_CHAR, "¶")
'Dim aKet() As Byte
'Dim abCipher() As Byte
'aKet() = StrConv("6$db sYS5&(£'S HseT£ w4uaz5 \gw43 y\4wu", vbFromUnicode)
'Call blf_KeyInit(aKet)
    If IsConnected(index) Then
        frmDebugWindow.txtDebug.Text = frmDebugWindow.txtDebug.Text & bufferLog & vbCrLf
        'abCipher = blf_BytesEnc(StrConv(Data, vbFromUnicode))
        frmServer.Socket(index).SendData Data ' StrConv(abCipher, vbUnicode)
        DoEvents
    End If
    
'Dim i As Long, n As Long, startc As Long
'
'    If IsConnected(index) Then
'        frmServer.Socket(index).SendData Data
'        DoEvents
'    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToMap(ByVal mapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal mapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapNum And i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub GlobalMsg(ByVal msg As String, ByVal Color As Long)
Dim packet As String

    packet = "GLOBALMSG" & SEP_CHAR & msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(packet)
End Sub

Sub MassMsg(ByVal msg As String)
Dim i As Long
Dim packet As String

packet = "MASSMSG" & SEP_CHAR & msg & SEP_CHAR & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And LCase(Trim(map(player(i).Char(player(i).CharNum).map).street)) = "(mass)" Then
            Call SendDataTo(i, packet)
        End If
    Next i
End Sub

Sub ServerMsg(ByVal msg As String)
Dim packet As String

    packet = "SERVERMSG" & SEP_CHAR & msg & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(packet)
End Sub

Sub AdminMsg(ByVal msg As String, ByVal Color As Long)
Dim packet As String
Dim i As Long

    packet = "ADMINMSG" & SEP_CHAR & msg & SEP_CHAR & RGB_AdminColor & SEP_CHAR & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, packet)
        End If
    Next i
End Sub

Sub PlayerMsg(ByVal index As Long, ByVal msg As String, ByVal Color As Long)
Dim packet As String

    packet = "PLAYERMSG" & SEP_CHAR & msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub GuildMsg(ByVal index As Long, ByVal msg As String, ByVal Color As Long)
Dim packet As String

    packet = "GUILDMSG" & SEP_CHAR & msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub PlayerSign(ByVal index As Long, ByVal header As String, ByVal msg As String)
Dim packet As String
    
    packet = "QUESTMSG" & SEP_CHAR & header & vbCrLf & vbCrLf & msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub MapMsg(ByVal mapNum As Long, ByVal msg As String, ByVal Color As Long)
Dim packet As String
Dim Text As String

    packet = "MAPMSG" & SEP_CHAR & msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(mapNum, packet)
End Sub

Sub AlertMsg(ByVal index As Long, ByVal msg As String)
Dim packet As String

    packet = "ALERTMSG" & SEP_CHAR & msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, packet)
    Call CloseSocket(index)
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String, Optional killConn As Boolean = False)
On Error Resume Next
Dim filenum As Long
Dim strplayerName As String
strplayerName = GetPlayerName(index)
filenum = FreeFile()
    If index > 0 Then
        If IsPlaying(index) Then
            Call AddLog(strplayerName & " has been booted for (" & Reason & ")", "hack", index)
'            Open App.Path & "\hackLog.txt" For Append As #filenum
'                Print #filenum, strplayerName & " has been booted for (" & Reason & ")"
'            Close #filenum
            'Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If
    If killConn Then
        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & ".")
    End If
    End If
End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
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
End Sub

Sub SocketConnected(ByVal index As Long)
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
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
On Error Resume Next
'Dim MapDATAPacket As String
Dim Buffer As String
Dim packet As String
Dim top As String * 3
Dim Start As Integer
    If index > 0 Then
        frmServer.Socket(index).GetData Buffer, vbString, DataLength
        If Buffer = "top" Then
            top = str(TotalOnlinePlayers)
            Call SendDataTo(index, top)
            Call CloseSocket(index)
        End If
        Dim bufferLog As String
        bufferLog$ = Replace(Buffer$, END_CHAR, "||")
        bufferLog$ = Replace(Buffer$, SEP_CHAR, "¶")

        'If ShowDebug = True Then
            frmDebugWindow.txtDebug.Text = frmDebugWindow.txtDebug.Text & bufferLog & vbCrLf
        'End If
        'AddLog bufferLog, "raw", index

        player(index).Buffer = player(index).Buffer & Buffer
        
        Start = InStr(player(index).Buffer, END_CHAR)
        Do While Start > 0
            packet = Mid(player(index).Buffer, 1, Start - 1)
            player(index).Buffer = Mid(player(index).Buffer, Start + 1, Len(player(index).Buffer))
            player(index).DataPackets = player(index).DataPackets + 1
            Start = InStr(player(index).Buffer, END_CHAR)
            If Len(packet) > 0 Then
                Call HandleData(index, packet)
            End If
            DoEvents
        Loop
                
        ' Check if elapsed time has passed
        player(index).DataBytes = player(index).DataBytes + DataLength
        If GetTickCount >= player(index).DataTimer + 1000 Then
            player(index).DataTimer = GetTickCount
            player(index).DataBytes = 0
            player(index).DataPackets = 0
            Exit Sub
        End If
        
        ' Check for data flooding
        If player(index).DataBytes > 1000 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Data Flooding", True)
            Exit Sub
        End If
        
        ' Check for packet flooding
        If player(index).DataPackets > 25 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Packet Flooding", True)
            Exit Sub
        End If
    End If
End Sub

Sub HandleData(ByVal index As Long, ByVal Data As String)
On Error Resume Next

Dim Parse() As String
Dim Name As String
Dim Password As String
Dim Sex As Long
Dim Class As Long
Dim CharNum As Long
Dim msg As String
Dim IPMask As String
Dim BanSlot As Long
Dim MsgTo As Long
Dim Dir As Long
Dim InvNum As Long
Dim Ammount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Movement As Long
Dim i As Long, n As Long, x As Long, y As Long, f As Long
Dim mapNum As Long
Dim s As String
Dim s1 As String
Dim tMapStart As Long, tMapEnd As Long
Dim ShopNum As Long, itemnum As Long
Dim DurNeeded As Long, GoldNeeded As Long
Dim plrGuild As Long
Dim sprite As Long
Dim JailX As Long
Dim JailY As Long

Dim plrChr As Long
Dim plrCon As Long
Dim plrWiz As Long
Dim plrDex As Long
Dim plrStr As Long
Dim itmChr As Long
Dim itmCon As Long
Dim itmWiz As Long
Dim itmDex As Long
Dim itmStr As Long
        
    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
    ' :::::::::::::::::::::::::::::::
    ' :: change name colour packet ::
    ' :::::::::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 12) = "/namecolour " Then
            Name = Mid(Parse(1), 13)
            If GetPlayerAccess(index) > 0 Then
                Call SetPlayerColour(index, Name, False)
                'Call SavePlayer(index)
                Call SendPlayerData(index)
                Call SavePlayer(index, False)
            Exit Sub
        End If
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: change name colour packet ::
    ' :::::::::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 13) = "/ignoreblocks" Then
            If GetPlayerAccess(index) > 1 Then
                player(index).Char(player(index).CharNum).ingnoreBlocks = Not player(index).Char(player(index).CharNum).ingnoreBlocks
            Exit Sub
        End If
    End If
    
    
    
    ' :::::::::::::::::::::::::::::::
    ' :: change name colour packet ::
    ' :::::::::::::::::::::::::::::::
    If Left(LCase(Parse(0)), 12) = "requestquest" Then
        If player(index).Char(player(index).CharNum).CurrentQuest > 0 Then
        Select Case player(index).Char(player(index).CharNum).QuestStatus
            Case Is = 1
                Call SendDataTo(index, "questmsg" & SEP_CHAR & procQuestMsg(index, Quests(player(index).Char(player(index).CharNum).CurrentQuest).StartQuestMsg) & END_CHAR)
            Case Is = 2
            Debug.Print Quests(player(index).Char(player(index).CharNum).CurrentQuest).GetItemQuestMsg
                Call SendDataTo(index, "questmsg" & SEP_CHAR & procQuestMsg(index, Quests(player(index).Char(player(index).CharNum).CurrentQuest).GetItemQuestMsg) & END_CHAR)
            Case Is = 3
                Call SendDataTo(index, "questmsg" & SEP_CHAR & procQuestMsg(index, Quests(player(index).Char(player(index).CharNum).CurrentQuest).FinishQuestMessage) & END_CHAR)
            Case Else
                Call PlayerMsg(index, "You have no quest", RGB_AlertColor)
            
        End Select
        Else
            Call SendDataTo(index, "questmsg" & SEP_CHAR & "You currently have no quest" & END_CHAR)
        End If
    End If
    
    
    ' ::::::::::::::::::::::::
    ' :: Player Info packet ::
    ' ::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 6) = "/info " Then
            Name = Mid(Parse(1), 7)
            If GetPlayerAccess(index) > 0 Then
                Call PlayerMsg(index, "-:: Player Information ::-", RGB_HelpColor)
                Call PlayerMsg(index, "==========================", RGB_HelpColor)
                Call PlayerMsg(index, "Name:" & player(index).Char(player(index).CharNum).Name, RGB_HelpColor)
                Call PlayerMsg(index, "Guild:" & Guild(player(index).Char(player(index).CharNum).Guild).Name, RGB_HelpColor)
                Call PlayerMsg(index, "Level:" & player(index).Char(player(index).CharNum).level, RGB_HelpColor)
                Call PlayerMsg(index, "IP:" & frmServer.Socket(index).RemoteHostIP, RGB_HelpColor)
                
                Call SetPlayerColour(index, Name, False)
                'Call SavePlayer(index)
                Call SendPlayerData(index)
                Call SavePlayer(index, False)
            Exit Sub
        End If
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: PLAYER IS STUCK!!!!!!!!!! ::
    ' :::::::::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 6) = "/stuck" Then
        Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: change text colour packet ::
    ' :::::::::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 12) = "/textcolour " Then
            Name = Mid(Parse(1), 13)
            If GetPlayerAccess(index) > 0 Then
                Call SetPlayerColour(index, Name, True)
                'Call SavePlayer(index)
                Call SendPlayerData(index)
                Call SavePlayer(index, False)
            Exit Sub
        End If
    End If
    
    
    ' :::::::::::::::::::::::::
    ' :: Guild create packet ::
    ' :::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 13) = "/guildcreate " Then
            Name = Mid(Parse(1), 14)
            
            i = findGuild(Name)
            If i > 0 Then
                'error guild exists
                Call PlayerMsg(index, "Guild Name In Use, Pick Another.", RGB_AlertColor)
            Else
                'guild name not in use
                Call CreateGuild(index, Name)
            End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Guild message packet ::
    ' ::::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 1) = "@" Then
            msg = Mid(Parse(1), 2)
            i = getPlayerGuildID(index)
            If i <= 0 Then
                'error no guild exists
                Call PlayerMsg(index, "Guild does not exist.", RGB_AlertColor)
            Else
                'guild name not in use
                Call sendGuildMsg(msg, i)
            End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Guild invite packet ::
    ' :::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 13) = "/guildinvite " Then
    'Dim guildarr
            Name = Mid(Parse(1), 14)
            'guildarr = Split(name, " ")
            i = player(FindPlayer(Trim(Name))).Char(player(FindPlayer(Trim(Name))).CharNum).Guild
            'i = findGuild(guildarr(1))
            If i > 0 Then
                'guild exists
                
                Call PlayerMsg(index, "Player in guild.", RGB_AlertColor)
            Else
                'no such guild
                Call guildInvite(index, Name, i)
                
            End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Guild accept packet ::
    ' :::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 13) = "/guildaccept " Then
    'Dim guildarr
            Name = Mid(Parse(1), 14)
            'guildarr = Split(name, " ")
            i = findGuild(Trim(Name))
            If i > 0 Then
                'guild exists
                Call guildAccept(index, i)
                Call PlayerMsg(index, "You have joined the guild", RGB_AlertColor)
            Else
                'no such guild
                Call PlayerMsg(index, "Sorry No such Guild Exists", RGB_AlertColor)
            End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Guild remove player packet ::
    ' ::::::::::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 13) = "/guildremove " Then
    'Dim guildarr
            Name = Mid(Parse(1), 14)
            'guildarr = Split(name, " ")
            i = FindPlayer(Trim(Name))
            If i > 0 Then
                'guild exists
                plrGuild = player(index).Char(player(index).CharNum).Guild
                Call guildRemovePlayer(index, plrGuild, Name)
            Else
                'no such player
                Call PlayerMsg(index, "Sorry No such Player Exists", RGB_AlertColor)
            End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::
    ' :: Guild promote player packet ::
    ' :::::::::::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 14) = "/guildpremote " Then
    'Dim guildarr
            Name = Mid(Parse(1), 15)
            'guildarr = Split(name, " ")
            i = FindPlayer(Trim(Name))
            If i > 0 Then
                'guild exists
                plrGuild = player(index).Char(player(index).CharNum).Guild
                Call guildPremotePlayer(index, plrGuild, Name)
            Else
                'no such player
                Call PlayerMsg(index, "Sorry No such Player Exists", RGB_AlertColor)
            End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Guild demote player packet ::
    ' ::::::::::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 13) = "/guilddemote " Then
    'Dim guildarr
            Name = Mid(Parse(1), 14)
            'guildarr = Split(name, " ")
            i = FindPlayer(Trim(Name))
            If i > 0 Then
                'guild exists
                plrGuild = player(index).Char(player(index).CharNum).Guild
                Call guildDemotePlayer(index, plrGuild, Name)
            Else
                'no such player
                Call PlayerMsg(index, "Sorry No such Player Exists", RGB_AlertColor)
            End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Guild remove guild  packet ::
    ' ::::::::::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 12) = "/guildremove" Then
            Call guildDispand(index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Guild edit guild  packet ::
    ' ::::::::::::::::::::::::::::::
    If Left(LCase(Parse(1)), 11) = "/guildedit " Then
        'plrGuild = Player(index).Char(Player(index).CharNum).Guild
        Name = Mid(Parse(1), 12)
            Call guildeditDiscription(index, Name)
        Exit Sub
    End If
    ' :::::::::::::::::::::::
    ' :: Guild info packet ::
    ' :::::::::::::::::::::::
    If Left(LCase(Parse(1)), 6) = "/guild" Then
        Call GuildInfo(index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Requesting admin menu help ::
    ' ::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "adminmenu" Then
        If IsPlaying(index) Then
        Dim lngaccessLevel As Long
            lngaccessLevel = GetPlayerAccess(index)
                Select Case lngaccessLevel
                    Case Is = ADMIN_MONITER
                        Call PlayerMsg(index, "-=: Admin Monitor :=-", RGB_HelpColor)
                        Call PlayerMsg(index, "Social Commands:", RGB_HelpColor)
                        Call PlayerMsg(index, """msghere = Global Admin Message", RGB_HelpColor)
                        Call PlayerMsg(index, "=msghere = Private Admin Message", RGB_HelpColor)
                        Call PlayerMsg(index, "Available Commands: /admin, /kick, /loc, /jail", RGB_HelpColor)
                    Case Is = ADMIN_MAPPER
                        Call PlayerMsg(index, "-=: Admin Mapper :=-", RGB_HelpColor)
                        Call PlayerMsg(index, "Social Commands:", RGB_HelpColor)
                        Call PlayerMsg(index, """msghere = Global Admin Message", RGB_HelpColor)
                        Call PlayerMsg(index, "=msghere = Private Admin Message", RGB_HelpColor)
                        Call PlayerMsg(index, "Available Commands: /admin, /loc, /mapeditor, /warpmeto, /warptome, /warpto, /mapreport, /kick, /ban, /jail, /respawn, /motd", RGB_HelpColor)
                    Case Is = ADMIN_DEVELOPER
                        Call PlayerMsg(index, "-=: Admin Developer :=-", RGB_HelpColor)
                        Call PlayerMsg(index, "Social Commands:", RGB_HelpColor)
                        Call PlayerMsg(index, """msghere = Global Admin Message", RGB_HelpColor)
                        Call PlayerMsg(index, "=msghere = Private Admin Message", RGB_HelpColor)
                        Call PlayerMsg(index, "Available Commands: /admin, /loc, /mapeditor, /warpmeto, /warptome, /warpto, /setsprite, /mapreport, /kick, /ban, /jail, /edititem, /respawn, /editnpc, /motd, /editshop, /editspell", RGB_HelpColor)
                    Case Is >= ADMIN_CREATOR
                        Call PlayerMsg(index, "-=: Admin Creator :=-", RGB_HelpColor)
                        Call PlayerMsg(index, "Social Commands:", RGB_HelpColor)
                        Call PlayerMsg(index, """msghere = Global Admin Message", RGB_HelpColor)
                        Call PlayerMsg(index, "=msghere = Private Admin Message", RGB_HelpColor)
                        Call PlayerMsg(index, "Available Commands: /admin, /loc, /mapeditor, /warpmeto, /warptome, /warpto, /setsprite, /mapreport, /kick, /ban, /jail, /edititem, /respawn, /editnpc, /motd, /editshop, /editspell", RGB_HelpColor)
                End Select
        End If
        Exit Sub
    End If
     
     
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Requesting classes for making a character ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "getclasses" Then
        If Not IsPlaying(index) Then
            Call SendNewCharClasses(index)
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::::
    ' :: New account packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "newaccount" Then
        If Not IsPlaying(index) And Not IsLoggedIn(index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
        
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
                Call AlertMsg(index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = Asc(Mid(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next i
            
            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(index, Name, Password)
                Call TextAdd(frmServer.txtText, "Account " & Name & " has been created.", True)
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                Call AlertMsg(index, "Your account has been created!")
            Else
                Call AlertMsg(index, "Sorry, that account name is already taken!")
            End If
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Delete account packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "delaccount" Then
        If Not IsPlaying(index) And Not IsLoggedIn(index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
            
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
                Call AlertMsg(index, "The name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If
            
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If
                        
            ' Delete names from master name file
            Call LoadPlayer(index, Name)
            For i = 1 To MAX_CHARS
                If Trim(player(index).Char(i).Name) <> "" Then
                    Call DeleteName(player(index).Char(i).Name)
                End If
            Next i
            Call ClearPlayer(index)
            
            ' Everything went ok
            Call Kill(App.Path & "\accounts\" & Trim(Name) & ".ini")
            Call AddLog("Account " & Trim(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(index, "Your account has been deleted.")
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::
    ' :: Login packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "login" Then
        If Not IsPlaying(index) And Not IsLoggedIn(index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
        
            ' Check versions
            If Val(Parse(3)) <> CLIENT_MAJOR Or Val(Parse(4)) <> CLIENT_MINOR Or Val(Parse(5)) <> CLIENT_REVISION Then
                'tell the client to launch the updater
                If Val(Parse(3)) <= CLIENT_MAJOR_UPDATE And Val(Parse(4)) <= CLIENT_MINOR_UPDATE And Val(Parse(5)) <= CLIENT_REVISION_UPDATE Then
                    Call SendDataTo(index, "VERSION1" & SEP_CHAR & END_CHAR)
                    CloseSocket (index)
                Else
                    Call SendDataTo(index, "VERSION" & SEP_CHAR & END_CHAR)
                    CloseSocket (index)
                End If
                Exit Sub
            End If
            
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
                Call AlertMsg(index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If
        
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If
        
            If IsMultiAccounts(Name) Then
                Call AlertMsg(index, "Multiple account logins is not authorized.")
                Exit Sub
            End If
                
            ' Everything went ok
    
            ' Load the player
            Call LoadPlayer(index, Name)
            Call SendChars(index)
    
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
            Call TextAdd(frmServer.txtText, GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", True)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Add character packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "addchar" Then
        If Not IsPlaying(index) Then
            Name = Parse(1)
            Sex = Val(Parse(2))
            Class = Val(Parse(3))
            CharNum = Val(Parse(4))
            sprite = Val(Parse(5))
        
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Then
                Call AlertMsg(index, "Character name must be at least three characters in length.")
                Exit Sub
            End If
            
            ' Prevent being me
            If LCase(Trim(Name)) = "consty" Then
                Call AlertMsg(index, "Lets get one thing straight, you are not me, ok? :)")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = Asc(Mid(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next i
                                    
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Prevent hacking
            If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
                Call HackingAttempt(index, "Invalid Sex (dont laugh)")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Class < 0 Or Class > Max_Classes Then
                Call HackingAttempt(index, "Invalid Path")
                Exit Sub
            End If
        
            ' Check if char already exists in slot
            If CharExist(index, CharNum) Then
                Call AlertMsg(index, "Character already exists!")
                Exit Sub
            End If
            
            ' Check if name is already in use
            If FindChar(Name) Then
                Call AlertMsg(index, "Sorry, but that name is in use!")
                Exit Sub
            End If
            
            ' Check if name is already in use
            If sprite < 0 Then
                Call AlertMsg(index, "Sorry, invalid Sprite!")
                Exit Sub
            End If
        
            ' Everything went ok, add the character
            Call AddChar(index, Name, Sex, Class, CharNum, sprite)
            'Call SavePlayer(index)
            Call AddLog("Character " & Name & " added to " & GetPlayerLogin(index) & "'s account.", PLAYER_LOG)
            Call AlertMsg(index, "Character has been created!")
        End If
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::::::::
    ' :: Deleting character packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "delchar" Then
        If Not IsPlaying(index) Then
            CharNum = Val(Parse(1))
        
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(index, "Invalid CharNum")
                Exit Sub
            End If
            
            Call DelChar(index, CharNum)
            Call AddLog("Character deleted on " & GetPlayerLogin(index) & "'s account.", PLAYER_LOG)
            Call AlertMsg(index, "Character has been deleted!")
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Using character packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "usechar" Then
        If Not IsPlaying(index) Then
            CharNum = Val(Parse(1))
        
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Check to make sure the character exists and if so, set it as its current char
            If CharExist(index, CharNum) Then
                player(index).CharNum = CharNum
                Call JoinGame(index)
            
                CharNum = player(index).CharNum
                Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
                Call TextAdd(frmServer.txtText, GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".", True)
                Call UpdateCaption
                'Name = loadPet(index, GetFreePetID(), player(index).Char(player(index).CharNum).PetName, player(index).Char(player(index).CharNum).PetSprite)
                'Call SendUpdatePetToAll(Name)
                ' Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
                If Not FindChar(GetPlayerName(index)) Then
                    f = FreeFile
                    Open App.Path & "\accounts\charlist.txt" For Append As #f
                        Print #f, GetPlayerName(index)
                    Close #f
                End If
            Else
                Call AlertMsg(index, "Character does not exist!")
            End If
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "saymsg" Then
        msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(msg)
            If Asc(Mid(msg, i, 1)) < 32 Or Asc(Mid(msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Say Text Modification")
                Exit Sub
            End If
        Next i
        'CHANGE
        Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & ": " & msg, PLAYER_LOG)
        Dim myCol As Long
        myCol = GetPlayerColour(index, True)
        If myCol = 0 Then myCol = 15
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & ": " & msg, myCol)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "emotemsg" Then
        msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(msg)
            If Asc(Mid(msg, i, 1)) < 32 Or Asc(Mid(msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Emote Text Modification")
                Exit Sub
            End If
        Next i
        
        Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & msg, PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & msg, RGB_EmoteColor)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "broadcastmsg" Then
        If blnBroadcast = False Then
            Exit Sub
        End If
            msg = Parse(1)
            
            ' Prevent hacking
            For i = 1 To Len(msg)
                If Asc(Mid(msg, i, 1)) < 32 Or Asc(Mid(msg, i, 1)) > 126 Then
                    Call HackingAttempt(index, "Broadcast Text Modification")
                    Exit Sub
                End If
            Next i
            
            s = GetPlayerName(index) & ": " & msg
            Call AddLog(s, PLAYER_LOG)
            
            Call GlobalMsg(s, RGB_BroadcastColor)
            Call TextAdd(frmServer.txtText, s, True)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "globalmsg" Then
        msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(msg)
            If Asc(Mid(msg, i, 1)) < 32 Or Asc(Mid(msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Global Text Modification")
                Exit Sub
            End If
        Next i
        
        If GetPlayerAccess(index) > 0 Then
            s = "(global) " & GetPlayerName(index) & ": " & msg
            Call AddLog(s, ADMIN_LOG)
            Call GlobalMsg(s, RGB_GlobalColor)
            Call TextAdd(frmServer.txtText, s, True)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "massmsg" Then
        msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(msg)
            If Asc(Mid(msg, i, 1)) < 32 Or Asc(Mid(msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Mass Text Modification")
                Exit Sub
            End If
        Next i
        
        If GetPlayerAccess(index) > 0 Then
            s = "Priest" & ": " & msg
            Call AddLog(s, ADMIN_LOG)
            Call MassMsg(s)
            Call TextAdd(frmServer.txtText, s, True)
        End If
        Exit Sub
    End If
    
    
    
    If LCase(Parse(0)) = "adminmsg" Then
        msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(msg)
            If Asc(Mid(msg, i, 1)) < 32 Or Asc(Mid(msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Admin Text Modification")
                Exit Sub
            End If
        Next i
        
        If GetPlayerAccess(index) > 0 Then
            Call AddLog("(admin " & GetPlayerName(index) & ") " & msg, ADMIN_LOG)
            Call AdminMsg("(admin " & GetPlayerName(index) & ") " & msg, RGB_AdminColor)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "playermsg" Then
        MsgTo = FindPlayer(Parse(1))
        msg = Parse(2)
        
        ' Prevent hacking
        For i = 1 To Len(msg)
            If Asc(Mid(msg, i, 1)) < 32 Or Asc(Mid(msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Player Msg Text Modification")
                Exit Sub
            End If
        Next i
        
        ' Check if they are trying to talk to themselves
        If MsgTo <> index Then
            If MsgTo > 0 Then
                Call AddLog(GetPlayerName(index) & " tells " & GetPlayerName(MsgTo) & ", " & msg & "'", PLAYER_LOG)
                Call PlayerMsg(MsgTo, GetPlayerName(index) & " tells you, '" & msg & "'", RGB_TellColor)
                Call PlayerMsg(index, "You tell " & GetPlayerName(MsgTo) & ", '" & msg & "'", RGB_TellColor)
            Else
                Call PlayerMsg(index, "Player is not online.", RGB_AlertColor)
            End If
        Else
            Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " begins to mumble to himself, what a wierdo...", PLAYER_LOG)
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " begins to mumble to himself, what a wierdo...", RGB_WHITE)
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playermove" And player(index).GettingMap = NO Then
        Dir = Val(Parse(1))
        Movement = Val(Parse(2))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(index, "Invalid Direction")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Movement < 1 Or Movement > 2 Then
            Call HackingAttempt(index, "Invalid Movement")
            Exit Sub
        End If
        
        ' Prevent player from moving if they have casted a spell
        If player(index).CastedSpell = YES Then
            ' Check if they have already casted a spell, and if so we can't let them move
            If GetTickCount > player(index).AttackTimer + 1000 Then
                player(index).CastedSpell = NO
            Else
                Call SendPlayerXY(index)
                Exit Sub
            End If
        End If
        
        Call PlayerMove(index, Dir, Movement)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerdir" And player(index).GettingMap = NO Then
        Dir = Val(Parse(1))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(index, "Invalid Direction")
            Exit Sub
        End If
        
        Call SetPlayerDir(index, Dir)
        Call SendDataToMapBut(index, GetPlayerMap(index), "PLAYERDIR" & SEP_CHAR & index & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
        
    ' :::::::::::::::::::::
    ' :: Use item packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "useitem" Then
        InvNum = Val(Parse(1))
        CharNum = player(index).CharNum
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(index, "Invalid CharNum")
            Exit Sub
        End If
        'Debug.Print GetPlayerInvItemNum(index, InvNum)
        
        If (GetPlayerInvItemNum(index, InvNum) > 0) And (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
            Call SendUpdateItemTo(index, GetPlayerInvItemNum(index, InvNum))
            n = Item(GetPlayerInvItemNum(index, InvNum)).Data2
                 plrChr = GetPlayerCHA(index)
                 plrCon = GetPlayerCON(index)
                 plrWiz = GetPlayerWIZ(index)
                 plrDex = GetPlayerDEX(index)
                 plrStr = GetPlayerSTR(index)
                 itmChr = Item(GetPlayerInvItemNum(index, InvNum)).cha
                 itmCon = Item(GetPlayerInvItemNum(index, InvNum)).con
                 itmWiz = Item(GetPlayerInvItemNum(index, InvNum)).wiz
                 itmDex = Item(GetPlayerInvItemNum(index, InvNum)).dex
                 itmStr = Item(GetPlayerInvItemNum(index, InvNum)).str
' Find out what kind of item it is
            Select Case Item(GetPlayerInvItemNum(index, InvNum)).type
                Case ITEM_TYPE_ARMOR
                    If InvNum <> GetPlayerArmorSlot(index) Then
                        'If Int(GetPlayerDEX(Index)) < n Then
                        If plrChr < itmChr And plrCon < itmCon And plrWiz < itmWiz And plrStr < itmStr And plrDex < itmDex Then
                            Call PlayerMsg(index, "Your stats are too low to wear this armor!  Required stats: ", RGB_HelpColor)
                            Call PlayerMsg(index, "DEX: " & itmDex & " STR: " & itmStr & " WIZ: " & itmWiz & " CON: " & itmCon & " CHR: " & itmChr, RGB_HelpColor)
                            Exit Sub
                        End If
                        Call SetPlayerArmorSlot(index, InvNum)
                    Else
                        Call SetPlayerArmorSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum <> GetPlayerWeaponSlot(index) Then
                        If plrChr < itmChr Or plrCon < itmCon Or plrWiz < itmWiz Or plrStr < itmStr Or plrDex < itmDex Then
                            Call PlayerMsg(index, "Your stats are too low to wear this armor!  Required stats: ", RGB_HelpColor)
                            Call PlayerMsg(index, "DEX: " & itmDex & " STR: " & itmStr & " WIZ: " & itmWiz & " CON: " & itmCon & " CHR: " & itmChr, RGB_HelpColor)
                            Exit Sub
                        End If
                        Call SetPlayerWeaponSlot(index, InvNum)
                    Else
                        Call SetPlayerWeaponSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
                        
                Case ITEM_TYPE_HELMET
                    If InvNum <> GetPlayerHelmetSlot(index) Then
                        If plrChr < itmChr And plrCon < itmCon And plrWiz < itmWiz And plrStr < itmStr And plrDex < itmDex Then
                            Call PlayerMsg(index, "Your stats are too low to wear this armor!  Required stats: ", RGB_HelpColor)
                            Call PlayerMsg(index, "DEX: " & itmDex & " STR: " & itmStr & " WIZ: " & itmWiz & " CON: " & itmCon & " CHR: " & itmChr, RGB_HelpColor)
                            Exit Sub
                        End If
                        Call SetPlayerHelmetSlot(index, InvNum)
                    Else
                        Call SetPlayerHelmetSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
            
                Case ITEM_TYPE_SHIELD
                    If InvNum <> GetPlayerShieldSlot(index) Then
                        Call SetPlayerShieldSlot(index, InvNum)
                    Else
                        Call SetPlayerShieldSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
            
                Case ITEM_TYPE_POTIONADDHP
                    Call SetPlayerHP(index, GetPlayerHP(index) + Item(player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendHP(index)
                
                Case ITEM_TYPE_POTIONADDMP
                    Call SetPlayerMP(index, GetPlayerMP(index) + Item(player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendMP(index)
        
                Case ITEM_TYPE_POTIONADDSP
                    Call SetPlayerSP(index, GetPlayerSP(index) + Item(player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendSP(index)

                Case ITEM_TYPE_POTIONSUBHP
                    Call SetPlayerHP(index, GetPlayerHP(index) - Item(player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendHP(index)
                
                Case ITEM_TYPE_POTIONSUBMP
                    Call SetPlayerMP(index, GetPlayerMP(index) - Item(player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendMP(index)
        
                Case ITEM_TYPE_POTIONSUBSP
                    Call SetPlayerSP(index, GetPlayerSP(index) - Item(player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendSP(index)
                    
                Case ITEM_TYPE_POTIONADDPP
                    Call SetPlayerPP(index, GetPlayerPP(index) + Item(player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendPP(index)
                    
                Case ITEM_TYPE_KEY
                    Select Case GetPlayerDir(index)
                        Case DIR_UP
                            If GetPlayerY(index) > 0 Then
                                x = GetPlayerX(index)
                                y = GetPlayerY(index) - 1
                            Else
                                Exit Sub
                            End If
                            
                        Case DIR_DOWN
                            If GetPlayerY(index) < MAX_MAPY Then
                                x = GetPlayerX(index)
                                y = GetPlayerY(index) + 1
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_LEFT
                            If GetPlayerX(index) > 0 Then
                                x = GetPlayerX(index) - 1
                                y = GetPlayerY(index)
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_RIGHT
                            If GetPlayerX(index) < MAX_MAPY Then
                                x = GetPlayerX(index) + 1
                                y = GetPlayerY(index)
                            Else
                                Exit Sub
                            End If
                    End Select
                    
                    ' Check if a key exists
                    If map(GetPlayerMap(index)).Tile(x, y).type = TILE_TYPE_KEY Then
                        ' Check if the key they are using matches the map key
                        If GetPlayerInvItemNum(index, InvNum) = map(GetPlayerMap(index)).Tile(x, y).Data1 Then
                            TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                            TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                            
                            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", RGB_HelpColor)
                            
                            ' Check if we are supposed to take away the item
                            If map(GetPlayerMap(index)).Tile(x, y).Data2 = 1 Then
                                Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                Call PlayerMsg(index, "The key disolves.", RGB_AlertColor)
                            End If
                        End If
                    End If
                    
                Case ITEM_TYPE_SPELL
                    ' Get the spell num
                    n = Item(GetPlayerInvItemNum(index, InvNum)).Data1
                    
                    If n > 0 Then
                        ' Make sure they are the right class
                        If Spell(n).ClassReq - 1 = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                            ' Make sure they are the right level
                            i = GetSpellReqLevel(index, n)
                            If i <= GetPlayerLevel(index) Then
                                i = FindOpenSpellSlot(index)
                                
                                ' Make sure they have an open spell slot
                                If i > 0 Then
                                    ' Make sure they dont already have the spell
                                    If Not HasSpell(index, n) Then
                                        Call SetPlayerSpell(index, i, n)
                                        Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                        Call PlayerMsg(index, "You study the spell carefully...", RGB_HelpColor)
                                        Call PlayerMsg(index, "You have learned a new spell!", RGB_HelpColor)
                                    Else
                                        Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                        Call PlayerMsg(index, "You have already learned this spell!  The spells crumbles into dust.", RGB_AlertColor)
                                    End If
                                Else
                                    Call PlayerMsg(index, "You have learned all that you can learn!", RGB_AlertColor)
                                End If
                            Else
                                Call PlayerMsg(index, "You must be level " & i & " to learn this spell.", RGB_AlertColor)
                            End If
                        Else
                            Call PlayerMsg(index, "This spell can only be learned by a " & GetClassName(Spell(n).ClassReq - 1) & ".", RGB_AlertColor)
                        End If
                    Else
                        Call PlayerMsg(index, "This scroll is not connected to a spell, please inform an admin!", RGB_AlertColor)
                    End If
                    
                Case ITEM_TYPE_PRAYER
                    ' Get the spell num
                    n = Item(GetPlayerInvItemNum(index, InvNum)).Data1
                    
                    If n > 0 Then
                        ' Make sure they are the right class
                        If Prayer(n).ClassReq - 1 = GetPlayerClass(index) Or Prayer(n).ClassReq = 0 Then
                            ' Make sure they are the right level
                            i = GetPrayerReqLevel(index, n)
                            If i <= GetPlayerLevel(index) Then
                                i = FindOpenPrayerSlot(index)
                                
                                ' Make sure they have an open prayer slot
                                If i > 0 Then
                                    ' Make sure they dont already have the spell
                                    If Not HasPrayer(index, n) Then
                                        Call SetPlayerPrayer(index, i, n)
                                        Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                        Call PlayerMsg(index, "You study the prayer carefully...", RGB_AlertColor)
                                        Call PlayerMsg(index, "You have learned a new prayer!", RGB_HelpColor)
                                    Else
                                        Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                        Call PlayerMsg(index, "You have already learned this prayer!  The prayer crumbles into dust.", RGB_AlertColor)
                                    End If
                                Else
                                    Call PlayerMsg(index, "You have learned all that you can learn!", RGB_AlertColor)
                                End If
                            Else
                                Call PlayerMsg(index, "You must be level " & i & " to learn this prayer.", RGB_AlertColor)
                            End If
                        Else
                            Call PlayerMsg(index, "This prayer can only be learned by a " & GetClassName(Prayer(n).ClassReq - 1) & ".", RGB_AlertColor)
                        End If
                    Else
                        Call PlayerMsg(index, "This scroll is not connected to a prayer, please inform an admin!", RGB_AlertColor)
                    End If
                    
            End Select
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "attack" Then
        ' Try to attack a player
        For i = 1 To MAX_PLAYERS
            ' Make sure we dont try to attack ourselves
            If i <> index Then
                ' Can we attack the player?
                If CanAttackPlayer(index, i) Then
                    If Not CanPlayerBlockHit(i) Then
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(index) Then
                            Damage = GetPlayerDamage(index) - GetPlayerProtection(i)
                        Else
                            n = GetPlayerDamage(index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                            Call PlayerMsg(index, "You feel a surge of energy upon swinging!", RGB_AlertColor)
                            Call PlayerMsg(i, GetPlayerName(index) & " swings with enormous might!", RGB_AlertColor)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(index, i, Damage)
                        Else
                            Call PlayerMsg(index, "Your attack does nothing.", RGB_AlertColor)
                        End If
                    Else
                        Call PlayerMsg(index, GetPlayerName(i) & "'s " & Trim(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked your hit!", RGB_AlertColor)
                        Call PlayerMsg(i, "Your " & Trim(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", RGB_AlertColor)
                    End If
                    
                    Exit Sub
                End If
            End If
        Next i
        
        ' Try to attack a npc
        For i = 1 To MAX_MAP_NPCS
            ' Can we attack the npc?
            If CanAttackNpc(index, i) Then
                ' Get the damage we can do
                If Not CanPlayerCriticalHit(index) Then
                    Damage = GetPlayerDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).def / 2)
                Else
                    n = GetPlayerDamage(index)
                   
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).def / 2)
                    Call PlayerMsg(index, "You feel a surge of energy upon swinging!", RGB_AlertColor)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(index, i, Damage)
                Else
                    Call PlayerMsg(index, "Your attack does nothing.", RGB_AlertColor)
                End If
                Exit Sub
            End If
        Next i
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Use stats packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "usestatpoint" Then
        PointType = Val(Parse(1))
        
        ' Prevent hacking
        If (PointType < 0) Or (PointType > 5) Then
            Call HackingAttempt(index, "Invalid Point Type")
            Exit Sub
        End If
                
        ' Make sure they have points
        If GetPlayerPOINTS(index) > 0 Then
            ' Take away a stat point
            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)
            
            ' Everything is ok
            Select Case PointType
                Case 0
                    Call SetPlayerSTR(index, GetPlayerSTR(index) + 1)
                    Call PlayerMsg(index, "You have gained more strength!", RGB_HelpColor)
                Case 1
                    Call SetPlayerINT(index, GetPlayerINT(index) + 1)
                    'Call SetPlayerDEF(index, GetPlayerDEF(index) + 1)
                    Call PlayerMsg(index, "You have became more inteligent!", RGB_HelpColor)
                Case 2
                    Call SetPlayerDEX(index, GetPlayerDEX(index) + 1)
                    'Call SetPlayerMAGI(index, GetPlayerMAGI(index) + 1)
                    Call PlayerMsg(index, "You have gained more speed!", RGB_HelpColor)
                Case 3
                    Call SetPlayerCON(index, GetPlayerCON(index) + 1)
                    'Call SetPlayerSPEED(index, GetPlayerSPEED(index) + 1)
                    Call PlayerMsg(index, "You have gained more constitution!", RGB_HelpColor)
                Case 4
                    Call SetPlayerWIZ(index, GetPlayerWIZ(index) + 1)
                    Call PlayerMsg(index, "You became wiser!", RGB_HelpColor)
                Case 5
                    Call SetPlayerCHA(index, GetPlayerCHA(index) + 1)
                    Call PlayerMsg(index, "You have gained more charisma!", RGB_HelpColor)
            End Select
        Else
            Call PlayerMsg(index, "You have no skill points to train with!", RGB_AlertColor)
        End If
        
        ' Send the update
        Call SendStats(index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Player info request packet ::
    ' ::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinforequest" Then
        If GetPlayerAccess(index) < 0 Then
            Exit Sub
        End If
        
        Name = Parse(1)
        
        i = FindPlayer(Name)
        If i > 0 Then
            Call PlayerMsg(index, "Account: " & Trim(player(i).Login) & ", Name: " & GetPlayerName(i), RGB_HelpColor)
            If GetPlayerAccess(index) > ADMIN_MONITER Then
                Call PlayerMsg(index, "-=- Stats for " & GetPlayerName(i) & " -=-", RGB_HelpColor)
                Call PlayerMsg(index, "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), RGB_HelpColor)
                Call PlayerMsg(index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), RGB_HelpColor)
                Call PlayerMsg(index, "Str: " & GetPlayerSTR(i) & "  Int: " & GetPlayerINT(i) & "  Dex: " & GetPlayerDEX(i) & "  Con: " & GetPlayerCON(i) & "  Wiz: " & GetPlayerWIZ(i) & "  Cha: " & GetPlayerCHA(i), RGB_HelpColor)
                n = Int(GetPlayerDEX(i) / 2) + Int(GetPlayerLevel(i) / 2)
                i = Int(GetPlayerDEX(i) / 2) + Int(GetPlayerLevel(i) / 2)
                If n > 100 Then n = 100
                If i > 100 Then i = 100
                Call PlayerMsg(index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", RGB_HelpColor)
                Call PlayerMsg(index, "Guild:" & Guild(player(index).Char(player(index).CharNum).Guild).Name, RGB_HelpColor)
                Call PlayerMsg(index, "IP:" & frmServer.Socket(index).RemoteHostIP, RGB_AlertColor)
            End If
        Else
            Call PlayerMsg(index, "Player is not online.", RGB_AlertColor)
        End If
        Exit Sub
    End If
    

    
    ' :::::::::::::::::::::::
    ' :: Warp me to packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "warpmeto" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MONITER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        n = FindPlayer(Parse(1))
        
        If n <> index Then
            If n > 0 Then
                Call PlayerWarp(index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
                Call PlayerMsg(n, GetPlayerName(index) & " has warped to you.", RGB_HelpColor)
                Call PlayerMsg(index, "You have been warped to " & GetPlayerName(n) & ".", RGB_HelpColor)
                Call AddLog(GetPlayerName(index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(index, "Player is not online.", RGB_AlertColor)
            End If
        Else
            Call PlayerMsg(index, "You cannot warp to yourself!", RGB_AlertColor)
        End If
                
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::
    ' :: Warp to me packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "warptome" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MONITER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        n = FindPlayer(Parse(1))
        
        If n <> index Then
            If n > 0 Then
                Call PlayerWarp(n, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(index) & ".", RGB_AlertColor)
                Call PlayerMsg(index, GetPlayerName(n) & " has been summoned.", RGB_AlertColor)
                Call AddLog(GetPlayerName(index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(index) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(index, "Player is not online.", RGB_AlertColor)
            End If
        Else
            Call PlayerMsg(index, "You cannot warp yourself to yourself!", RGB_AlertColor)
        End If
        
        Exit Sub
    End If


    ' ::::::::::::::::::::::::
    ' :: Warp to map packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "warpto" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MONITER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        'map coords
        Dim YWarp As Long
        Dim XWarp As Long
        'Dim parse1() As String
        'parse1 = Split(Parse(1), "|")
        
        If UBound(Parse()) = 3 Then
            XWarp = Val(Parse(2))
            YWarp = Val(Parse(3))
            n = Val(Parse(1))
        Else
            ' The map
            XWarp = GetPlayerX(index)
            YWarp = GetPlayerY(index)
            n = Val(Parse(1))
        End If
        
        
        ' Prevent hacking
        If n < 0 Or n > MAX_MAPS Then
            Call HackingAttempt(index, "Invalid map")
            Exit Sub
        End If
        
        Call PlayerWarp(index, n, XWarp, YWarp)
        Call PlayerMsg(index, "You have been warped to map #" & n, RGB_AlertColor)
        Call AddLog(GetPlayerName(index) & " warped to map #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Warp to map packet used for /ignoreblocks ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "warpto_u" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < 2 Or player(index).Char(player(index).CharNum).ingnoreBlocks = False Then
            Exit Sub
        End If
        
        'map coords
        Dim YWarp_u As Long
        Dim XWarp_u As Long
        'Dim parse1() As String
        'parse1 = Split(Parse(1), "|")
        
        If UBound(Parse()) = 3 Then
            XWarp_u = Val(Parse(2))
            YWarp_u = Val(Parse(3))
            n = Val(Parse(1))
        Else
            ' The map
            XWarp_u = GetPlayerX(index)
            YWarp_u = GetPlayerY(index)
            n = Val(Parse(1))
        End If
        
        
        ' Prevent hacking
        If n < 0 Or n > MAX_MAPS Then
            Call HackingAttempt(index, "Invalid map")
            Exit Sub
        End If
        
        Call PlayerWarp(index, n, XWarp_u, YWarp_u)
        Call AddLog(GetPlayerName(index) & " warped to map #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    

    
    ' :::::::::::::::::::::::
    ' :: Set sprite packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "setsprite" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The sprite
        n = Val(Parse(1))
        
        Call SetPlayerSprite(index, n)
        Call SendPlayerData(index)
        Exit Sub
    End If
                
    ' ::::::::::::::::::::::::::
    ' :: Stats request packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "getstats" Then
        SendStatsInfo (index)
'        Call PlayerMsg(index, "-=- Stats for " & GetPlayerName(index) & " -=-", White)
'        Call PlayerMsg(index, "Level: " & GetPlayerLevel(index) & "  Exp: " & GetPlayerExp(index) & "/" & GetPlayerNextLevel(index), White)
'        Call PlayerMsg(index, "HP: " & GetPlayerHP(index) & "/" & GetPlayerMaxHP(index) & "  MP: " & GetPlayerMP(index) & "/" & GetPlayerMaxMP(index) & "  SP: " & GetPlayerSP(index) & "/" & GetPlayerMaxSP(index), White)
'        Call PlayerMsg(index, "Str: " & GetPlayerSTR(index) & "  Int: " & GetPlayerINT(index) & "  Dex: " & GetPlayerDEX(index) & "  Con: " & GetPlayerCON(index) & "  Wiz: " & GetPlayerWIZ(index) & "  Cha: " & GetPlayerCHA(index), White)
'        n = Int(GetPlayerDEX(index) / 2) + Int(GetPlayerLevel(index) / 2)
'        i = Int(GetPlayerDEX(index) / 2) + Int(GetPlayerLevel(index) / 2)
'        If n > 100 Then n = 100
'        If i > 100 Then i = 100
'        Call PlayerMsg(index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", White)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player request for a new map ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requestnewmap" Then
        Dir = Val(Parse(1))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(index, "Invalid Direction")
            Exit Sub
        End If
                
        Call PlayerMove(index, Dir, 1)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    Dim tempStorage As String ' for map part 1
    
    If LCase(Parse(0)) = "mapdata" Then
         ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Or GetPlayerAccess(index) = ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = 1
        
        mapNum = GetPlayerMap(index)
        map(mapNum).Name = Parse(n + 1)
        map(mapNum).Revision = map(mapNum).Revision + 1
        map(mapNum).Moral = Val(Parse(n + 3))
        map(mapNum).Up = Val(Parse(n + 4))
        map(mapNum).Down = Val(Parse(n + 5))
        map(mapNum).Left = Val(Parse(n + 6))
        map(mapNum).Right = Val(Parse(n + 7))
        map(mapNum).Music = Val(Parse(n + 8))
        map(mapNum).BootMap = Val(Parse(n + 9))
        map(mapNum).BootX = Val(Parse(n + 10))
        map(mapNum).BootY = Val(Parse(n + 11))
        map(mapNum).Shop = Val(Parse(n + 12))
        
        n = n + 13
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                map(mapNum).Tile(x, y).Ground = Val(Parse(n))
                map(mapNum).Tile(x, y).Mask = Val(Parse(n + 1))
                map(mapNum).Tile(x, y).Anim = Val(Parse(n + 2))
                map(mapNum).Tile(x, y).Fringe = Val(Parse(n + 3))
                map(mapNum).Tile(x, y).type = Val(Parse(n + 4))
                map(mapNum).Tile(x, y).Data1 = Val(Parse(n + 5))
                map(mapNum).Tile(x, y).Data2 = Val(Parse(n + 6))
                map(mapNum).Tile(x, y).Data3 = Val(Parse(n + 7))
                map(mapNum).Tile(x, y).Data4 = Val(Parse(n + 8))
                map(mapNum).Tile(x, y).Data5 = Val(Parse(n + 9))
                'map(MapNum).Tile(x, y).Data6 = Val(Parse(n + 10))
                'map(MapNum).Tile(x, y).Data7 = Val(Parse(n + 11))
                'map(MapNum).Tile(x, y).Data8 = Val(Parse(n + 12))
                'map(MapNum).Tile(x, y).Data9 = Val(Parse(n + 13))
                'map(MapNum).Tile(x, y).Data10 = Val(Parse(n + 14))
                map(mapNum).Tile(x, y).TileSheet_Ground = Val(Parse(n + 10))
                map(mapNum).Tile(x, y).TileSheet_Fringe = Val(Parse(n + 11))
                map(mapNum).Tile(x, y).TileSheet_Anim = Val(Parse(n + 12))
                map(mapNum).Tile(x, y).TileSheet_Mask = Val(Parse(n + 13))
            
                n = n + 14
            Next x
        Next y
        
        For x = 1 To MAX_MAP_NPCS
            map(mapNum).Npc(x) = Val(Parse(n))
            n = n + 1
            Call ClearMapNpc(x, mapNum)
        Next x
        map(mapNum).Respawn = (Parse(n))
        map(mapNum).Night = Val(Parse(n + 1))
        map(mapNum).Bank = CBool(Parse(n + 2))
        map(mapNum).street = Trim(Parse(n + 3))
        Call SendMapNpcsToMap(mapNum)
        Call SpawnMapNpcs(mapNum)
        
        ' Save the map
        Call SaveMap(mapNum)
        
        ' Refresh map for everyone online
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = mapNum Then
                Call PlayerWarp(i, mapNum, GetPlayerX(i), GetPlayerY(i))
            End If
        Next i
        
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Need map yes/no packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "needmap" Then
        ' Get yes/no value
        s = LCase(Parse(1))
                'MsgBox s
        If s = "yes" Then
            Call SendMap(index, GetPlayerMap(index))
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            player(index).GettingMap = NO
            Call SendDataTo(index, "MAPDONE" & SEP_CHAR & END_CHAR)
            DoEvents
            Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Else
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            player(index).GettingMap = NO
            Call SendDataTo(index, "MAPDONE" & SEP_CHAR & END_CHAR)
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to pick up something packet ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapgetitem" Then
        Call PlayerMapGetItem(index)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to pick up something packet ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapgetsign" Then
        Call PlayerMapGetSign(index)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to pick up something packet ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapgetlevel" Then
        Call PlayerMapGetLevel(index)
        Exit Sub
    End If

    
    ' ::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to drop something packet ::
    ' ::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapdropitem" Then
        InvNum = Val(Parse(1))
        Ammount = Val(Parse(2))
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_INV Then
            Call HackingAttempt(index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Ammount > GetPlayerInvItemValue(index, InvNum) Then
            Call HackingAttempt(index, "Item ammount modification")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Item(GetPlayerInvItemNum(index, InvNum)).type = ITEM_TYPE_CURRENCY Then
            ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
            If Ammount <= 0 Then
                Call HackingAttempt(index, "Trying to drop 0 ammount of currency")
                Exit Sub
            End If
        End If
            
        Call PlayerMapDropItem(index, InvNum, Ammount)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Respawn map packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "maprespawn" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Or GetPlayerAccess(index) = ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Clear out it all
        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
            Call ClearMapItem(i, GetPlayerMap(index))
        Next i
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(index))
        
        ' Respawn NPCS
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, GetPlayerMap(index))
        Next i
        
        Call PlayerMsg(index, "Map respawned.", RGB_AlertColor)
        Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Map report packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "mapreport" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Or GetPlayerAccess(index) = ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        s = "Free Maps: "
        tMapStart = 1
        tMapEnd = 1
        s1 = SEP_CHAR
        For i = 1 To MAX_MAPS
            s1 = s1 & Trim(map(i).Name) & SEP_CHAR
            If Trim(map(i).Name) = "" Then
                tMapEnd = tMapEnd + 1
            Else
                If tMapEnd - tMapStart > 0 Then
                    s = s & Trim(str(tMapStart)) & "-" & Trim(str(tMapEnd - 1)) & ", "
                End If
                tMapStart = i + 1
                tMapEnd = i + 1
            End If
        Next i
        
        s = s & Trim(str(tMapStart)) & "-" & Trim(str(tMapEnd - 1)) & ", "
        s = Mid(s, 1, Len(s) - 2)
        s = s & "."
        
        Call PlayerMsg(index, s, RGB_LIGHTGREY)
        Call SendDataTo(index, "MAPREPORT" & s1 & END_CHAR)
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Kick player packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "kickplayer" Then
        ' Prevent hacking
        If GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player index
        n = FindPlayer(Parse(1))
        
        If n <> index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(index) Then
                    Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(index) & "!", RGB_GlobalColor)
                    Call AddLog(GetPlayerName(index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                    Call AlertMsg(n, "You have been kicked by " & GetPlayerName(index) & "!")
                Else
                    Call PlayerMsg(index, "That is a higher access admin then you!", RGB_AlertColor)
                End If
            Else
                Call PlayerMsg(index, "Player is not online.", RGB_AlertColor)
            End If
        Else
            Call PlayerMsg(index, "You cannot kick yourself!", RGB_AlertColor)
        End If
                
        Exit Sub
    End If
        
    ' :::::::::::::::::::::
    ' :: Ban list packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "banlist" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = 1
        f = FreeFile
        Open App.Path & "\banlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            Input #f, Name
            
            Call PlayerMsg(index, n & ": Banned IP " & s & " by " & Name, RGB_AlertColor)
            n = n + 1
        Loop
        Close #f
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Ban destroy packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "bandestroy" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_CREATOR Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call Kill(App.Path & "\banlist.txt")
        Call PlayerMsg(index, "Ban list destroyed.", RGB_AlertColor)
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::
    ' :: Ban player packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "banplayer" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player index
        n = FindPlayer(Parse(1))
        
        If n <> index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(index) Then
                    Call BanIndex(n, index)
                Else
                    Call PlayerMsg(index, "That is a higher access admin then you!", RGB_AlertColor)
                End If
            Else
                Call PlayerMsg(index, "Player is not online.", RGB_AlertColor)
            End If
        Else
            Call PlayerMsg(index, "You cannot ban yourself!", RGB_AlertColor)
        End If
                
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Request edit map packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditmap" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Or GetPlayerAccess(index) = ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "EDITMAP" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Request edit BIO        ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditbio" Then
        ' Prevent hacking
'        If GetPlayerAccess(index) < ADMIN_MAPPER Or GetPlayerAccess(index) = ADMIN_DEVELOPER Then
'            Call HackingAttempt(index, "Admin Cloning")
'            Exit Sub
'        End If
        
        Call SendDataTo(index, "EDITBIO" & SEP_CHAR & player(index).Bio & SEP_CHAR & player(index).RealName & SEP_CHAR & player(index).Email & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit item packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requestedititem" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "ITEMEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit quest packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditquest" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "QUESTEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If


    ' ::::::::::::::::::::::
    ' :: Edit item packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "edititem" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The item #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid Item Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing item #" & n & ".", ADMIN_LOG)
        Call SendEditItemTo(index, n)
    End If
    
    ' ::::::::::::::::::::::
    ' :: Save item packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "saveitem" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid Item Index")
            Exit Sub
        End If
        
        ' Update the item
        Item(n).Name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Item(n).BaseDamage = Val(Parse(8))
        Item(n).str = Val(Parse(9))
        Item(n).intel = Val(Parse(10))
        Item(n).dex = Val(Parse(11))
        Item(n).con = Val(Parse(12))
        Item(n).wiz = Val(Parse(13))
        Item(n).cha = Val(Parse(14))
        
        Item(n).Description = Parse(15)
        Item(n).weaponType = Parse(16)
        ' Save it
        Call SendUpdateItemToAll(n)
        Call SaveItem(n)
        Call AddLog(GetPlayerName(index) & " saved item #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::
    ' :: Request edit npc packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditnpc" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "NPCEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Edit npc packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "editnpc" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The npc #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(index, "Invalid NPC Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing npc #" & n & ".", ADMIN_LOG)
        Call SendEditNpcTo(index, n)
    End If
    
    ' :::::::::::::::::::::
    ' :: Save npc packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "savenpc" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(index, "Invalid NPC Index")
            Exit Sub
        End If
        
        ' Update the npc
        Npc(n).Name = Parse(2)
        Npc(n).AttackSay = Parse(3)
        Npc(n).sprite = Val(Parse(4))
        Npc(n).SpawnSecs = Val(Parse(5))
        Npc(n).Behavior = Val(Parse(6))
        Npc(n).Range = Val(Parse(7))
        Npc(n).DropChance = Val(Parse(8))
        Npc(n).DropItem = Val(Parse(9))
        Npc(n).DropItemValue = Val(Parse(10))
        Npc(n).str = Val(Parse(11))
        Npc(n).def = Val(Parse(12))
        Npc(n).speed = Val(Parse(13))
        Npc(n).MAGI = Val(Parse(14))
        Npc(n).HP = Val(Parse(15))
        Npc(n).ExpGiven = Val(Parse(16))
        Npc(n).Respawn = Parse(17)
        Npc(n).Attack_with_Poison = Parse(18)
        Npc(n).Poison_length = Parse(19)
        Npc(n).Poison_vital = Parse(20)
        Npc(n).QuestID = Parse(21)
        Npc(n).opensBank = Parse(22)
        Npc(n).opensShop = Parse(23)
        Npc(n).type = Parse(24)
        
        ' Save it
        Call SendUpdateNpcToAll(n)
        Call SaveNpc(n)
        Call AddLog(GetPlayerName(index) & " saved npc #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
            
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit shop packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditshop" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "SHOPEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "editshop" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The shop #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SHOPS Then
            Call HackingAttempt(index, "Invalid Shop Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing shop #" & n & ".", ADMIN_LOG)
        Call SendEditShopTo(index, n)
    End If
    
    ' ::::::::::::::::::::::
    ' :: Save shop packet ::
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "saveshop") Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ShopNum = Val(Parse(1))
        
        ' Prevent hacking
        If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
            Call HackingAttempt(index, "Invalid Shop Index")
            Exit Sub
        End If
        
        ' Update the shop
        Shop(ShopNum).Name = Parse(2)
        Shop(ShopNum).JoinSay = Parse(3)
        Shop(ShopNum).LeaveSay = Parse(4)
        Shop(ShopNum).FixesItems = Val(Parse(5))
        
        n = 6
        For i = 1 To MAX_TRADES
            Shop(ShopNum).TradeItem(i).GiveItem = Val(Parse(n))
            Shop(ShopNum).TradeItem(i).GiveValue = Val(Parse(n + 1))
            Shop(ShopNum).TradeItem(i).GetItem = Val(Parse(n + 2))
            Shop(ShopNum).TradeItem(i).GetValue = Val(Parse(n + 3))
            n = n + 4
        Next i
        
        ' Save it
        Call SendUpdateShopToAll(ShopNum)
        Call SaveShop(ShopNum)
        Call AddLog(GetPlayerName(index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Request edit spell packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditspell" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "SPELLEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Request edit prayer packet ::
    ' ::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditprayer" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "PRAYEREDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Request edit sign packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditsign" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "SIGNEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Request edit sign packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditspell" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "SIGNEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "editspell" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The spell #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(index, "Invalid Spell Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing spell #" & n & ".", ADMIN_LOG)
        Call SendEditSpellTo(index, n)
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Edit prayer packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "editprayer" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The prayer #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(index, "Invalid Spell Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing prayer #" & n & ".", ADMIN_LOG)
        Call SendEditPrayerTo(index, n)
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Edit quest packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "editquest" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The quest #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_QUESTS Then
            Call HackingAttempt(index, "Invalid Quest Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing quest #" & n & ".", ADMIN_LOG)
        Call SendEditQuestTo(index, n)
    End If
    
    
    ' ::::::::::::::::::::::
    ' :: Edit sign packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "editsign" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The sign #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SIGNS Then
            Call HackingAttempt(index, "Invalid Sign Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing sign #" & n & ".", ADMIN_LOG)
        Call SendEditsignTo(index, n)
    End If
    
    ' :::::::::::::::::::::::
    ' :: Save spell packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "savespell") Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Spell #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(index, "Invalid Spell Index")
            Exit Sub
        End If
        
        ' Update the spell
        Spell(n).Name = Parse(2)
        Spell(n).ClassReq = Val(Parse(3))
        Spell(n).LevelReq = Val(Parse(4))
        Spell(n).type = Val(Parse(5))
        Spell(n).Data1 = Val(Parse(6))
        Spell(n).Data2 = Val(Parse(7))
        Spell(n).Data3 = Val(Parse(8))
        Spell(n).sound = Val(Parse(9))
        Spell(n).manaUse = Val(Parse(10))
                
        ' Save it
        Call SendUpdateSpellToAll(n)
        Call SaveSpell(n)
        Call AddLog(GetPlayerName(index) & " saving spell #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Save prayer packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "saveprayer") Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Spell #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(index, "Invalid Spell Index")
            Exit Sub
        End If
        
        ' Update the spell
        Prayer(n).Name = Parse(2)
        Prayer(n).ClassReq = Val(Parse(3))
        Prayer(n).LevelReq = Val(Parse(4))
        Prayer(n).type = Val(Parse(5))
        Prayer(n).Data1 = Val(Parse(6))
        Prayer(n).Data2 = Val(Parse(7))
        Prayer(n).Data3 = Val(Parse(8))
        Prayer(n).sound = Val(Parse(9))
        Prayer(n).manaUse = Val(Parse(10))
                
        ' Save it
        Call SendUpdatePrayerToAll(n)
        Call SavePrayer(n)
        Call AddLog(GetPlayerName(index) & " saving prayer #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Save quest packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "savequest") Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' quest #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_QUESTS Then
            Call HackingAttempt(index, "Invalid quest Index")
            Exit Sub
        End If
        
        ' Update the quest
        With Quests(n)
            .ExpGiven = Parse(2)
            .FinishQuestMessage = Parse(3)
            .GetItemQuestMsg = Parse(4)
            .ItemGiven = Parse(5)
            .ItemToObtain = Parse(6)
            .ItemValGiven = Parse(7)
            .requiredLevel = Parse(8)
            .StartQuestMsg = Parse(9)
            .GetItemQuestMsg = Parse(10)
            .goldGiven = Parse(11)
        End With
                
        ' Save it
        Call SendUpdateQuestToAll(n)
        Call SaveQuest(n)
        Call AddLog(GetPlayerName(index) & " saving quest #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Save BIO    packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "savebio") Then
        
        ' Update the quest
        With player(index)
            .Bio = Parse(1)
            .RealName = Parse(2)
            .Email = Parse(3)
        End With
                
        ' Save it
        Call SaveBio(index)
        Exit Sub
    End If
    ' :::::::::::::::::::::::
    ' :: Save sign packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "savesign") Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' sign #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(index, "Invalid sign Index")
            Exit Sub
        End If
        
        ' Update the spell
        Signs(n).header = Parse(2)
        Signs(n).msg = Parse(3)
                
        ' Save it
        Call SendUpdateSignToAll(n)
        Call SaveSign(n)
        Call AddLog(GetPlayerName(index) & " saving sign #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Set access packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "setaccess" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_CREATOR Then
            Call HackingAttempt(index, "Trying to use powers not available")
            Exit Sub
        End If
        
        ' The index
        n = FindPlayer(Parse(1))
        ' The access
        i = Val(Parse(2))
        
        
        ' Check for invalid access level
        If i >= 0 Or i <= 3 Then
            ' Check if player is on
            If n > 0 Then
                Select Case i
                'SET ACCESS TEXTs
                    Case Is = ADMIN_MONITER
                    'You have been granted the power of a monitor
                        Call GlobalMsg(GetPlayerName(n) & " has been granted the power of a monitor.", BrightBlue)
                    Case Is = ADMIN_MAPPER
                        Call GlobalMsg(GetPlayerName(n) & " has been granted the power of a mapper.", BrightBlue)
                    Case Is = ADMIN_DEVELOPER
                        Call GlobalMsg(GetPlayerName(n) & " has been granted the power of a developer.", BrightBlue)
                    Case Is = ADMIN_CREATOR
                        Call GlobalMsg(GetPlayerName(n) & " has been granted the power of a creator.", BrightBlue)
                End Select
                If i = 0 Then
                    Call SetPlayerColour(n, 15, False)
                    Call SetPlayerColour(n, 0, True)
                End If
                Call SetPlayerAccess(n, i)
                Call SendPlayerData(n)
                Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(n) & "'s access to " & i & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(index, "Player is not online.", RGB_AlertColor)
            End If
        Else
            Call PlayerMsg(index, "Invalid access level.", RGB_AlertColor)
        End If
                
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Who online packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "whosonline" Then
        Call SendWhosOnline(index)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Set MOTD packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "setmotd" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PutVar(App.Path & "\motd.ini", "MOTD", "Msg", Parse(1))
        Call GlobalMsg("MOTD changed to: " & Parse(1), RGB_GlobalColor)
        Call AddLog(GetPlayerName(index) & " changed MOTD to: " & Parse(1), ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: ITEM lib packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "itemlib" Then
        ' Prevent hacking
        n = getItemNo(Parse(1))
        If n > 0 Then
            Call sendItemLib(index, n)
        Else
            Call PlayerMsg(index, "There is no such item.", RGB_AlertColor)
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: check level packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "checklevel" Then
        CheckPlayerLevelUp (index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "trade" Then
    If blnNight <> True Then
        If map(GetPlayerMap(index)).Shop > 0 Then
            Call SendTrade(index, map(GetPlayerMap(index)).Shop)
        Else
            Call PlayerMsg(index, "There is no shop here.", RGB_AlertColor)
        End If
    Else
        Call PlayerMsg(index, "This shop has closed for the night.", RGB_HelpColor)
    End If
        Exit Sub
    End If
    ' ::::::::::::::::::
    ' :: BIO packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "bio" Then
        
        Name = Parse(1)
        Name = Trim(Name)
        i = FindPlayer(Name)
        msg = Name & "'s Bio" & vbCrLf
        msg = msg & "-------------------------------" & vbCrLf
        
        msg = msg & "Name: " & player(i).RealName & vbCrLf
        msg = msg & "e-mail: " & player(i).Email & vbCrLf
        msg = msg & "Bio: " & vbCrLf & player(i).Bio
        Dim packet As String

        packet = "QUESTMSG" & SEP_CHAR & msg & SEP_CHAR & END_CHAR
        Call SendDataTo(index, packet)
        Exit Sub
    End If
    

    ' ::::::::::::::::::::::::::
    ' :: Trade request packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "traderequest" Then
        ' Trade num
        n = Val(Parse(1))
        
        ' Prevent hacking
        If (n <= 0) Or (n > MAX_TRADES) Then
            Call HackingAttempt(index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Index for shop
        i = map(GetPlayerMap(index)).Shop
        
        ' Check if inv full
        x = FindOpenInvSlot(index, Shop(i).TradeItem(n).GetItem)
        If x = 0 Then
            Call PlayerMsg(index, "Trade unsuccessful, inventory full.", RGB_AlertColor)
            Exit Sub
        End If
        
        ' Check if they have the item
        If HasItem(index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
            Call TakeItem(index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
            Call GiveItem(index, Shop(i).TradeItem(n).GetItem, Shop(i).TradeItem(n).GetValue)
            Call PlayerMsg(index, "The trade was successful!", RGB_HelpColor)
        Else
            Call PlayerMsg(index, "Trade unsuccessful.", RGB_AlertColor)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Fix item packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "fixitem" Then
        ' Inv num
        n = Val(Parse(1))
        
        ' Make sure its a equipable item
        If Item(GetPlayerInvItemNum(index, n)).type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(index, n)).type > ITEM_TYPE_SHIELD Then
            Call PlayerMsg(index, "You can only fix weapons, armors, helmets, and shields.", RGB_AlertColor)
            Exit Sub
        End If
        
        ' Check if they have a full inventory
        If FindOpenInvSlot(index, GetPlayerInvItemNum(index, n)) <= 0 Then
            Call PlayerMsg(index, "You have no inventory space left!", RGB_AlertColor)
            Exit Sub
        End If
        
        ' Now check the rate of pay
        itemnum = GetPlayerInvItemNum(index, n)
        i = Int(Item(GetPlayerInvItemNum(index, n)).Data2 / 5)
        If i <= 0 Then i = 1
        
        DurNeeded = Item(itemnum).Data1 - GetPlayerInvItemDur(index, n)
        GoldNeeded = Int(DurNeeded * i / 2)
        If GoldNeeded <= 0 Then GoldNeeded = 1
        
        ' Check if they even need it repaired
        If DurNeeded <= 0 Then
            Call PlayerMsg(index, "This item is in perfect condition!", RGB_HelpColor)
            Exit Sub
        End If
        
        ' Check if they have enough for at least one point
        If HasItem(index, 2) >= i Then
            ' Check if they have enough for a total restoration
            If HasItem(index, 2) >= GoldNeeded Then
                Call TakeItem(index, 2, GoldNeeded)
                Call SetPlayerInvItemDur(index, n, Item(itemnum).Data1)
                Call PlayerMsg(index, "Item has been totally restored for " & GoldNeeded & " shards!", RGB_HelpColor)
            Else
                ' They dont so restore as much as we can
                DurNeeded = (HasItem(index, 1) / i)
                GoldNeeded = Int(DurNeeded * i / 2)
                If GoldNeeded <= 0 Then GoldNeeded = 1
                
                Call TakeItem(index, 2, GoldNeeded)
                Call SetPlayerInvItemDur(index, n, GetPlayerInvItemDur(index, n) + DurNeeded)
                Call PlayerMsg(index, "Item has been partially fixed for " & GoldNeeded & " shards!", RGB_HelpColor)
            End If
        Else
            Call PlayerMsg(index, "Insufficient shards to fix this item!", RGB_AlertColor)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: getcoords packet  from right clicking a square::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "getcords" Then
        x = Val(Parse(1))
        y = Val(Parse(2))
        If (GetPlayerAccess(index) > 0) Then
            Call PlayerMsg(index, "Map: " & GetPlayerMap(index) & "  " & "X: " & x & "  " & "Y: " & y, RGB_AlertColor)
            'Call PlayerMsg(index, "X: " & x, RGB_AlertColor)
            'Call PlayerMsg(index, "Y: " & y, RGB_AlertColor)
        End If
    End If
        

    ' :::::::::::::::::::
    ' :: Search packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "search" Then
        x = Val(Parse(1))
        y = Val(Parse(2))
        
        ' Prevent subscript out of range
        If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then
            Exit Sub
        End If
        
        ' Check for a player
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) And GetPlayerX(i) = x And GetPlayerY(i) = y Then
                
                ' Consider the player
                If GetPlayerLevel(i) >= GetPlayerLevel(index) + 5 Then
                    Call PlayerMsg(index, "You wouldn't stand a chance.", RGB_AlertColor)
                Else
                    If GetPlayerLevel(i) > GetPlayerLevel(index) Then
                        Call PlayerMsg(index, "This one seems to have an advantage over you.", RGB_HelpColor)
                    Else
                        If GetPlayerLevel(i) = GetPlayerLevel(index) Then
                            Call PlayerMsg(index, "This would be an even fight.", RGB_HelpColor)
                        Else
                            If GetPlayerLevel(index) >= GetPlayerLevel(i) + 5 Then
                                Call PlayerMsg(index, "You could slaughter that player.", RGB_AlertColor)
                            Else
                                If GetPlayerLevel(index) > GetPlayerLevel(i) Then
                                    Call PlayerMsg(index, "You would have an advantage over that player.", RGB_HelpColor)
                                End If
                            End If
                        End If
                    End If
                End If
            
                ' Change target
                player(index).target = i
                player(index).TargetType = TARGET_TYPE_PLAYER
                Call PlayerMsg(index, "Your target is now " & GetPlayerName(i) & ".", RGB_HelpColor)
                Exit Sub
            End If
        Next i
        
        ' Check for an item
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(GetPlayerMap(index), i).num > 0 Then
                If MapItem(GetPlayerMap(index), i).x = x And MapItem(GetPlayerMap(index), i).y = y Then
                    Call PlayerMsg(index, "You see a " & Trim(Item(MapItem(GetPlayerMap(index), i).num).Name) & ".", RGB_HelpColor)
                    Exit Sub
                End If
            End If
        Next i
        
        ' Check for an npc
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(index), i).num > 0 Then
                If MapNpc(GetPlayerMap(index), i).x = x And MapNpc(GetPlayerMap(index), i).y = y Then
                    If map(GetPlayerMap(index)).Shop > 0 And blnNight = False And Npc(MapNpc(GetPlayerMap(index), i).num).opensShop = True Then
                        'send a load trade panel
                        Call SendTrade(index, map(GetPlayerMap(index)).Shop)
                    End If
                    If map(GetPlayerMap(index)).Bank = True And blnNight = False And Npc(MapNpc(GetPlayerMap(index), i).num).opensBank = True Then
                        'send a load bank panel
                        Call sendBank(index)
                    End If
                    If Npc(MapNpc(GetPlayerMap(index), i).num).QuestID > 0 Then
                        Call StartQuest(index, Npc(MapNpc(GetPlayerMap(index), i).num).QuestID)
                    End If
                    ' Change target
                    player(index).target = i
                    player(index).TargetType = TARGET_TYPE_NPC
                    Call PlayerMsg(index, "Your target is now a " & Trim(Npc(MapNpc(GetPlayerMap(index), i).num).Name) & ".", RGB_HelpColor)
                    Exit Sub
                End If
            End If
        Next i
        
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Party packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "party" Then
        n = FindPlayer(Parse(1))
        
        ' Prevent partying with self
        If n = index Then
            Exit Sub
        End If
                
        ' Check for a previous party and if so drop it
        If player(index).InParty = YES Then
            Call PlayerMsg(index, "You are already in a party!", RGB_AlertColor)
            Exit Sub
        End If
        
        If n > 0 Then
            ' Check if its an admin
            If GetPlayerAccess(index) > ADMIN_MONITER Then
                Call PlayerMsg(index, "You can't join a party, you are an admin!", RGB_AlertColor)
                Exit Sub
            End If
        
            If GetPlayerAccess(n) > ADMIN_MONITER Then
                Call PlayerMsg(index, "Admins cannot join parties!", RGB_AlertColor)
                Exit Sub
            End If
            
            ' Make sure they are in right level range
            If GetPlayerLevel(index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(index) - 5 > GetPlayerLevel(n) Then
                Call PlayerMsg(index, "There is more then a 5 level gap between you two, party failed.", RGB_AlertColor)
                Exit Sub
            End If
            
            ' Check to see if player is already in a party
            If player(n).InParty = NO Then
                Call PlayerMsg(index, "Party request has been sent to " & GetPlayerName(n) & ".", RGB_HelpColor)
                Call PlayerMsg(n, GetPlayerName(index) & " wants you to join their party.  Type /join to join, or /leave to decline.", RGB_HelpColor)
            
                player(index).PartyStarter = YES
                player(index).PartyPlayer = n
                player(n).PartyPlayer = index
            Else
                Call PlayerMsg(index, "Player is already in a party!", RGB_AlertColor)
            End If
        Else
            Call PlayerMsg(index, "Player is not online.", RGB_AlertColor)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Join party packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "joinparty" Then
        n = player(index).PartyPlayer
        
        If n > 0 Then
            ' Check to make sure they aren't the starter
            If player(index).PartyStarter = NO Then
                ' Check to make sure that each of there party players match
                If player(n).PartyPlayer = index Then
                    Call PlayerMsg(index, "You have joined " & GetPlayerName(n) & "'s party!", RGB_HelpColor)
                    Call PlayerMsg(n, GetPlayerName(index) & " has joined your party!", RGB_HelpColor)
                    
                    player(index).InParty = YES
                    player(n).InParty = YES
                Else
                    Call PlayerMsg(index, "Party failed.", RGB_AlertColor)
                End If
            Else
                Call PlayerMsg(index, "You have not been invited to join a party!", RGB_AlertColor)
            End If
        Else
            Call PlayerMsg(index, "You have not been invited into a party!", RGB_AlertColor)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Join party packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "stamina" Then
        n = Parse(1)
        If Val(n) > GetPlayerMaxSP(index) Then n = GetPlayerMaxSP(index)
        Call SetPlayerSP(index, Val(n))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Leave party packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "leaveparty" Then
        n = player(index).PartyPlayer
        
        If n > 0 Then
            If player(index).InParty = YES Then
                Call PlayerMsg(index, "You have left the party.", RGB_AlertColor)
                Call PlayerMsg(n, GetPlayerName(index) & " has left the party.", RGB_AlertColor)
                
                player(index).PartyPlayer = 0
                player(index).PartyStarter = NO
                player(index).InParty = NO
                player(n).PartyPlayer = 0
                player(n).PartyStarter = NO
                player(n).InParty = NO
            Else
                Call PlayerMsg(index, "Declined party request.", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " declined your request.", RGB_AlertColor)
                
                player(index).PartyPlayer = 0
                player(index).PartyStarter = NO
                player(index).InParty = NO
                player(n).PartyPlayer = 0
                player(n).PartyStarter = NO
                player(n).InParty = NO
            End If
        Else
            Call PlayerMsg(index, "You are not in a party!", RGB_AlertColor)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "spells" Then
        Call SendPlayerSpells(index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Prayers packet ::
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "prayers" Then
        Call SendPlayerPrayers(index)
        Exit Sub
    End If
    
    ' :::::::::::::::::
    ' :: Cast packet ::
    ' :::::::::::::::::
    If LCase(Parse(0)) = "cast" Then
        ' Spell slot
        n = Val(Parse(1))
        
        Call CastSpell(index, n)
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Cast a prayer packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "castp" Then
        ' prayer slot
        n = Val(Parse(1))
        
        Call CastPrayer(index, n)
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: prayer packet ::  ' boosts players PP
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "prayer" Then
        
        'Call Pray(index)
        
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Location packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "requestlocation" Then
        If GetPlayerAccess(index) < 1 Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Map: " & GetPlayerMap(index) & ", X: " & GetPlayerX(index) & ", Y: " & GetPlayerY(index), RGB_LIGHTGREY)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: bank update packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerbank" Then
    If map(GetPlayerMap(index)).Bank = True Then
        Call sendBank(index)
    End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: bank update packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "bankitem" Then
    n = Val(Parse(1))
        If GetPlayerInvItemNum(index, n) > 2 And GetPlayerInvItemNum(index, n) <= MAX_ITEMS Then
            For i = 1 To 30 Step 1
            'Debug.Print "player bank " & GetPlayerBankItemNum(index, i)
                If GetPlayerBankItemNum(index, i) = 0 Then
                    'Debug.Print GetPlayerInvItemNum(index, n)
                    
                    Call SetPlayerBankItemNum(index, i, GetPlayerInvItemNum(index, n))
                    Call SetPlayerBankItemDur(index, i, GetPlayerInvItemDur(index, n))
                    Call SetPlayerBankItemValue(index, i, GetPlayerInvItemValue(index, n))
                    
                    Call SetPlayerInvItemNum(index, n, 0)
                    Call SetPlayerInvItemDur(index, n, 0)
                    Call SetPlayerInvItemValue(index, n, 0)
                    Exit For
                End If
            Next i
            Call SendInventoryUpdate(index, n)
            Call SendBankUpdate(index, n)
            'Call sendBank(index)
            DoEvents
            Call SavePlayerBank(index)
        End If
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: bank update packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "unbankitem" Then
    n = Val(Parse(1))
        If GetPlayerBankItemNum(index, n) > 2 And GetPlayerBankItemNum(index, n) <= MAX_ITEMS Then
            For i = 1 To 30 Step 1
            'Debug.Print "player bank " & GetPlayerInvItemNum(index, i)
                If GetPlayerInvItemNum(index, i) = 0 Then
                    'Debug.Print GetPlayerBankItemNum(index, n)
                    
                    Call SetPlayerInvItemNum(index, i, GetPlayerBankItemNum(index, n))
                    Call SetPlayerInvItemDur(index, i, GetPlayerBankItemDur(index, n))
                    Call SetPlayerInvItemValue(index, i, GetPlayerBankItemValue(index, n))
                    DoEvents
                    Call SetPlayerBankItemNum(index, n, 0)
                    Call SetPlayerBankItemDur(index, n, 0)
                    Call SetPlayerBankItemValue(index, n, 0)
                    Exit For
                End If
            Next i
            Call SendInventoryUpdate(index, i)
            Call SendBankUpdate(index, n)
            'Call sendBank(index)
            DoEvents
            Call SavePlayerBank(index)
        End If
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: bank update packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "bankgold" Then
    Dim hasPaid As Boolean
    Dim hasRecieved As Boolean
    
    n = Val(Parse(1))
        'If GetPlayerBankItemNum(index, n) > 2 And GetPlayerBankItemNum(index, n) <= MAX_ITEMS Then
        If HasItem(index, 2) >= n Then
            For i = 1 To 30 Step 1
            
                If GetPlayerBankItemNum(index, i) = 2 And hasRecieved = False Then
                    Call SetPlayerBankItemValue(index, i, GetPlayerBankItemValue(index, i) + n)
                    hasRecieved = True
                    Call SendBankUpdate(index, i)
                    DoEvents
                End If
                If GetPlayerInvItemNum(index, i) = 2 And hasPaid = False Then
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - n)
                    Call SendInventoryUpdate(index, i)
                    hasPaid = True
                    DoEvents
                End If
            Next i
                If hasRecieved = False Then
                    For i = 1 To 30 Step 1
                        If GetPlayerBankItemNum(index, i) = 0 Then
                            Call SetPlayerBankItemNum(index, i, 2)
                            Call SetPlayerBankItemValue(index, i, GetPlayerBankItemValue(index, i) + n)
                            Call SendBankUpdate(index, i)
                            hasRecieved = True
                            DoEvents
                            Exit For
                        End If
                    Next i
                End If
                
        End If
            
            
            'Call sendBank(index)
            DoEvents
            Call SavePlayerBank(index)
        'End If
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: bank update packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "unbankgold" Then
    'Dim hasPaid As Boolean
    'Dim hasRecieved As Boolean
    
    n = Val(Parse(1))
        'If GetPlayerBankItemNum(index, n) > 2 And GetPlayerBankItemNum(index, n) <= MAX_ITEMS Then
        If BankHasItem(index, 2) >= n Then
            For i = 1 To 30 Step 1
            
                If GetPlayerBankItemNum(index, i) = 2 And hasPaid = False Then
                    Call SetPlayerBankItemValue(index, i, GetPlayerBankItemValue(index, i) - n)
                    hasPaid = True
                    Call SendBankUpdate(index, i)
                    DoEvents
                End If
                If GetPlayerInvItemNum(index, i) = 2 And hasRecieved = False Then
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + n)
                    Call SendInventoryUpdate(index, i)
                    hasRecieved = True
                    DoEvents
                End If
            Next i
                If hasRecieved = False Then
                    For i = 1 To 30 Step 1
                        If GetPlayerInvItemNum(index, i) = 0 Then
                            Call SetPlayerInvItemNum(index, i, 2)
                            Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + n)
                            Call SendInventoryUpdate(index, i)
                            hasRecieved = True
                            DoEvents
                            Exit For
                        End If
                    Next i
                End If
                
        End If
            
            
            'Call sendBank(index)
            DoEvents
            Call SavePlayerBank(index)
        'End If
        
        Exit Sub
    End If
    

' ::::::::::::::::::::::::
' :: Jail player packet ::
' ::::::::::::::::::::::::
If LCase(Parse(0)) = "jailplayer" Then
' Prevent hacking
If GetPlayerAccess(index) < ADMIN_MONITER Then
Call HackingAttempt(index, "Admin Cloning")
Exit Sub
End If

' The map number and player to jail
n = FindPlayer(Parse(1))

If n <> index Then
If n > 0 Then
    Select Case Val(Parse(2))
        Case Is = 1
            JailX = 1
            JailY = 3
        Case Is = 2
            JailX = 5
            JailY = 3
        Case Is = 3
            JailX = 10
            JailY = 3
        Case Is = 4
            JailX = 1
            JailY = 10
        Case Is = 5
            JailX = 5
            JailY = 10
        Case Is = 6
            JailX = 10
            JailY = 10
        Case Else
            JailX = 6
            JailY = 8
    End Select
Call PlayerWarp(n, 1000, JailX, JailY)
Call PlayerMsg(n, "You have been jailed by " & GetPlayerName(index) & ".", RGB_AlertColor)
Call PlayerMsg(index, GetPlayerName(n) & " has been jailed.", RGB_AlertColor)
Call AddLog(GetPlayerName(index) & " has jailed " & GetPlayerName(n) & GetPlayerMap(index) & ".", ADMIN_LOG)
Else
Call PlayerMsg(index, "Player is not online.", RGB_AlertColor)
End If
Else
Call PlayerMsg(index, "You cannot jail yourself!", RGB_AlertColor)
End If

Exit Sub
End If
'Dim i As Long
For i = 0 To UBound(Parse)
    'Debug.Print "Parse(" & i & ") = " & Parse(i)
Next i
'Debug.Print "Parse(0) = " & Parse(0)
End Sub

Sub CloseSocket(ByVal index As Long)
    ' Make sure player was/is playing the game, and if so, save'm.
    If index > 0 Then
        Call LeftGame(index)
    
        Call TextAdd(frmServer.txtText, "Connection from " & GetPlayerIP(index) & " has been terminated.", True)
        
        frmServer.Socket(index).Close
            
        Call UpdateCaption
        Call ClearPlayer(index)
    End If
End Sub

Sub SendWhosOnline(ByVal index As Long)
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
        
    Call PlayerMsg(index, s, RGB_WhoColor)
End Sub

Sub sendItemLib(ByVal index As Long, ByVal number As Long)
Dim packet As String
    
    packet = "ITEMLIB" & SEP_CHAR
    packet = packet & Trim(Item(number).Name) & SEP_CHAR & Trim(Item(number).Description) & SEP_CHAR & Trim(Item(number).Pic) & SEP_CHAR
    packet = packet & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub


Sub SendChars(ByVal index As Long)
Dim packet As String
Dim i As Long
    
    packet = "ALLCHARS" & SEP_CHAR
    For i = 1 To MAX_CHARS
        packet = packet & Trim(player(index).Char(i).Name) & SEP_CHAR & Trim(Class(player(index).Char(i).Class).Name) & SEP_CHAR & player(index).Char(i).level & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub SendJoinMap(ByVal index As Long)
Dim packet As String
Dim i As Long

    packet = ""
    
    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index And GetPlayerMap(i) = GetPlayerMap(index) Then
            packet = packet & "PLAYERDATA" & SEP_CHAR & i & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & GetPlayerSprite(i) & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & GetPlayerX(i) & SEP_CHAR & GetPlayerY(i) & SEP_CHAR & GetPlayerDir(i) & SEP_CHAR & GetPlayerAccess(i) & SEP_CHAR & GetPlayerPK(i) & SEP_CHAR & GetPlayerColour(i, False) & SEP_CHAR & GetPlayerHP(i) & SEP_CHAR & GetPlayerMP(i) & SEP_CHAR & GetPlayerMaxHP(i) & SEP_CHAR & GetPlayerMaxMP(i) & END_CHAR
            Call SendDataTo(index, packet)
        End If
    Next i
    
    ' Send index's player data to everyone on the map including himself
    packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & GetPlayerColour(index, False) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerMaxMP(index) & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal mapNum As Long)
Dim packet As String

    packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & GetPlayerColour(index, False) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerMaxMP(index) & END_CHAR
    Call SendDataToMapBut(index, mapNum, packet)
End Sub

Sub SendPlayerData(ByVal index As Long)
Dim packet As String

    ' Send index's player data to everyone including himself on th emap
    packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & GetPlayerColour(index, False) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerMaxMP(index) & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)
End Sub

Sub SendMap(ByVal index As Long, ByVal mapNum As Long)
Dim packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long
    packet = "MAPDATA" & SEP_CHAR & mapNum & SEP_CHAR & Trim(map(mapNum).Name) & SEP_CHAR & map(mapNum).Revision & SEP_CHAR & map(mapNum).Moral & SEP_CHAR & map(mapNum).Up & SEP_CHAR & map(mapNum).Down & SEP_CHAR & map(mapNum).Left & SEP_CHAR & map(mapNum).Right & SEP_CHAR & map(mapNum).Music & SEP_CHAR & map(mapNum).BootMap & SEP_CHAR & map(mapNum).BootX & SEP_CHAR & map(mapNum).BootY & SEP_CHAR & map(mapNum).Shop & SEP_CHAR
    
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With map(mapNum).Tile(x, y)
                packet = packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Fringe & SEP_CHAR & .type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .Data4 & SEP_CHAR & .Data5 & SEP_CHAR & .TileSheet_Ground & SEP_CHAR & .TileSheet_Fringe & SEP_CHAR & .TileSheet_Anim & SEP_CHAR & .TileSheet_Mask & SEP_CHAR
                'packet = packet & .TileSheet_Ground & SEP_CHAR & .TileSheet_Fringe & SEP_CHAR & .TileSheet_Anim & SEP_CHAR & .TileSheet_Mask & SEP_CHAR
            End With
        Next x
    Next y
    'Call SendDataTo(index, packet & END_CHAR)
    'DoEvents
    'packet = "MAPDATA2" & SEP_CHAR
     
    
    For x = 1 To MAX_MAP_NPCS
        packet = packet & map(mapNum).Npc(x) & SEP_CHAR
    Next x
    
    packet = packet & map(mapNum).Respawn & SEP_CHAR & map(mapNum).Night & SEP_CHAR & map(mapNum).Bank & SEP_CHAR & map(mapNum).street & SEP_CHAR & mapNum & SEP_CHAR & END_CHAR
    
    'x = Int(Len(packet) / 2)
    'P1 = Mid(packet, 1, x)
    'P2 = Mid(packet, x + 1, Len(packet) - x)
    Call SendDataTo(index, packet)
    DoEvents
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal mapNum As Long)
Dim packet As String
Dim i As Long

    packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        packet = packet & MapItem(mapNum, i).num & SEP_CHAR & MapItem(mapNum, i).value & SEP_CHAR & MapItem(mapNum, i).Dur & SEP_CHAR & MapItem(mapNum, i).x & SEP_CHAR & MapItem(mapNum, i).y & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub SendMapItemsToAll(ByVal mapNum As Long)
Dim packet As String
Dim i As Long

    packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        packet = packet & MapItem(mapNum, i).num & SEP_CHAR & MapItem(mapNum, i).value & SEP_CHAR & MapItem(mapNum, i).Dur & SEP_CHAR & MapItem(mapNum, i).x & SEP_CHAR & MapItem(mapNum, i).y & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    
    Call SendDataToMap(mapNum, packet)
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal mapNum As Long)
Dim packet As String
Dim i As Long

    packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        packet = packet & MapNpc(mapNum, i).num & SEP_CHAR & MapNpc(mapNum, i).x & SEP_CHAR & MapNpc(mapNum, i).y & SEP_CHAR & MapNpc(mapNum, i).Dir & SEP_CHAR & MapNpc(mapNum, i).maxHP & SEP_CHAR & MapNpc(mapNum, i).HP & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub SendMapNpcsToMap(ByVal mapNum As Long)
Dim packet As String
Dim i As Long

    packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        packet = packet & MapNpc(mapNum, i).num & SEP_CHAR & MapNpc(mapNum, i).x & SEP_CHAR & MapNpc(mapNum, i).y & SEP_CHAR & MapNpc(mapNum, i).Dir & SEP_CHAR & MapNpc(mapNum, i).maxHP & SEP_CHAR & MapNpc(mapNum, i).HP & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    
    Call SendDataToMap(mapNum, packet)
End Sub

Sub SendItems(ByVal index As Long)
Dim packet As String
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Trim(Item(i).Name) <> "" Then
            Call SendUpdateItemTo(index, i)
        End If
    Next i
End Sub

Sub SendSigns(ByVal index As Long)
Dim packet As String
Dim i As Long

    For i = 1 To MAX_SIGNS
        If Trim(Signs(i).header) <> "" Then
            Call SendUpdateSignTo(index, i)
        End If
    Next i
End Sub

Sub SendNpcs(ByVal index As Long)
Dim packet As String
Dim i As Long

    For i = 1 To MAX_NPCS
        If Trim(Npc(i).Name) <> "" Then
            Call SendUpdateNpcTo(index, i)
        End If
    Next i
End Sub

'Sub SendPets(ByVal index As Long)
'Dim packet As String
'Dim i As Long
'
'    For i = 1 To MAX_PLAYERS
'        If Trim(Pets(i).Name) <> "" Then
'            Call SendUpdatePetTo(index, i)
'        End If
'    Next i
'End Sub

Sub SendInventory(ByVal index As Long)
Dim packet As String
Dim i As Long

    packet = "PLAYERINV" & SEP_CHAR
    For i = 1 To MAX_INV
        packet = packet & GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub sendBank(ByVal index As Long)
Dim packet As String
Dim i As Long

    packet = "PLAYERBANK" & SEP_CHAR
    For i = 1 To MAX_BANK
        packet = packet & GetPlayerBankItemNum(index, i) & SEP_CHAR & GetPlayerBankItemValue(index, i) & SEP_CHAR & GetPlayerBankItemDur(index, i) & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    'Debug.Print packet
    Call SendDataTo(index, packet)
    
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
Dim packet As String
    
    packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & END_CHAR
    'Debug.Print packet
    Call SendDataTo(index, packet)
End Sub

Sub SendBankUpdate(ByVal index As Long, ByVal InvSlot As Long)
Dim packet As String
    
    packet = "PLAYERBANKUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerBankItemNum(index, InvSlot) & SEP_CHAR & GetPlayerBankItemValue(index, InvSlot) & SEP_CHAR & GetPlayerBankItemDur(index, InvSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendWornEquipment(ByVal index As Long)
Dim packet As String
    
    packet = "PLAYERWORNEQ" & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendHP(ByVal index As Long)
Dim packet As String

    packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendMP(ByVal index As Long)
Dim packet As String

    packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendSP(ByVal index As Long)
Dim packet As String

    packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendPP(ByVal index As Long)
Dim packet As String

    packet = "PLAYERPP" & SEP_CHAR & GetPlayerMaxPP(index) & SEP_CHAR & GetPlayerPP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendNight(ByVal Night As Boolean)
Dim packet As String
If Night Then
    packet = "NIGHT" & SEP_CHAR & "1" & END_CHAR
Else
    packet = "NIGHT" & SEP_CHAR & "0" & END_CHAR
End If
    Call SendDataToAll(packet)
End Sub

Sub SendStatsInfo(ByVal index As Long)
Dim packet As String
Dim n As Long
Dim i As Long
n = Int(GetPlayerDEX(index) / 2) + Int(GetPlayerLevel(index) / 2)
i = Int(GetPlayerDEX(index) / 2) + Int(GetPlayerLevel(index) / 2)
If n > 100 Then n = 100
If i > 100 Then i = 100
    
    packet = "PLAYERSTATSINFO" & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerINT(index) & SEP_CHAR & GetPlayerDEX(index) & SEP_CHAR & GetPlayerCON(index) & SEP_CHAR & GetPlayerWIZ(index) & SEP_CHAR & GetPlayerCHA(index) & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerNextLevel(index) & SEP_CHAR & n & SEP_CHAR & i & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendStats(ByVal index As Long)
Dim packet As String
    
    packet = "PLAYERSTATS" & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerINT(index) & SEP_CHAR & GetPlayerDEX(index) & SEP_CHAR & GetPlayerCON(index) & SEP_CHAR & GetPlayerWIZ(index) & SEP_CHAR & GetPlayerCHA(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendWelcome(ByVal index As Long)
Dim MOTD As String
Dim f As Long

    ' Send them welcome
    Call PlayerMsg(index, "Welcome to " & GAME_NAME & "!  Version " & CLIENT_MAJOR & "." & CLIENT_MINOR & "." & CLIENT_REVISION, RGB_JoinLeftColor)
    Call PlayerMsg(index, "Type /help for help on commands.  Use arrow keys to move, hold down shift to run, and use ctrl to attack.", RGB_HelpColor)
    ' Send them MOTD
    MOTD = GetVar(App.Path & "\motd.ini", "MOTD", "Msg")
    If Trim(MOTD) <> "" Then
        Call PlayerMsg(index, "MOTD: " & MOTD, RGB_WHITE)
    End If
    
    ' Send whos online
    'Call SendWhosOnline(index)
End Sub

Sub SendClasses(ByVal index As Long)
Dim packet As String
Dim i As Long

    packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        packet = packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).str & SEP_CHAR & Class(i).intel & SEP_CHAR & Class(i).dex & SEP_CHAR & Class(i).con & SEP_CHAR & Class(i).wiz & SEP_CHAR & Class(i).cha & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub SendNewCharClasses(ByVal index As Long)
Dim packet As String
Dim i As Long

    packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        packet = packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).str & SEP_CHAR & Class(i).intel & SEP_CHAR & Class(i).dex & SEP_CHAR & Class(i).con & SEP_CHAR & Class(i).wiz & SEP_CHAR & Class(i).cha & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub SendLeftGame(ByVal index As Long)
Dim packet As String

    packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & "" & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & END_CHAR
    Call SendDataToAllBut(index, packet)
End Sub

Sub SendPlayerXY(ByVal index As Long)
Dim packet As String

    packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendUpdateItemToAll(ByVal itemnum As Long)
Dim packet As String

    packet = "UPDATEITEM" & SEP_CHAR & itemnum & SEP_CHAR & Trim(Item(itemnum).Name) & SEP_CHAR & Item(itemnum).Pic & SEP_CHAR & Item(itemnum).type & SEP_CHAR & Item(itemnum).Data1 & SEP_CHAR & Item(itemnum).Data2 & SEP_CHAR & Item(itemnum).Data3 & SEP_CHAR & Item(itemnum).BaseDamage & SEP_CHAR & Item(itemnum).str & SEP_CHAR & Item(itemnum).intel & SEP_CHAR & Item(itemnum).dex & SEP_CHAR & Item(itemnum).con & SEP_CHAR & Item(itemnum).wiz & SEP_CHAR & Item(itemnum).cha & SEP_CHAR & Item(itemnum).Description & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal itemnum As Long)
Dim packet As String

    packet = "UPDATEITEM" & SEP_CHAR & itemnum & SEP_CHAR & Trim(Item(itemnum).Name) & SEP_CHAR & Item(itemnum).Pic & SEP_CHAR & Item(itemnum).type & SEP_CHAR & SEP_CHAR & Item(itemnum).Pic & SEP_CHAR & Item(itemnum).Description & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendUpdateSignToAll(ByVal SignNum As Long)
Dim packet As String

    packet = "UPDATESIGN" & SEP_CHAR & SignNum & SEP_CHAR & Trim(Signs(SignNum).header) & SEP_CHAR & Trim(Signs(SignNum).msg) & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateSignTo(ByVal index As Long, ByVal SignNum As Long)
Dim packet As String

    packet = "UPDATESIGN" & SEP_CHAR & SignNum & SEP_CHAR & Trim(Signs(SignNum).header) & SEP_CHAR & Trim(Signs(SignNum).msg) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendEditItemTo(ByVal index As Long, ByVal itemnum As Long)
Dim packet As String

    packet = "EDITITEM" & SEP_CHAR & itemnum & SEP_CHAR & Trim(Item(itemnum).Name) & SEP_CHAR & Item(itemnum).Pic & SEP_CHAR & Item(itemnum).type & SEP_CHAR & Item(itemnum).Data1 & SEP_CHAR & Item(itemnum).Data2 & SEP_CHAR & Item(itemnum).Data3 & SEP_CHAR & Item(itemnum).BaseDamage & SEP_CHAR & Item(itemnum).str & SEP_CHAR & Item(itemnum).intel & SEP_CHAR & Item(itemnum).dex & SEP_CHAR & Item(itemnum).con & SEP_CHAR & Item(itemnum).wiz & SEP_CHAR & Item(itemnum).cha & SEP_CHAR & Item(itemnum).Description & SEP_CHAR & Item(itemnum).weaponType & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim packet As String

    packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).sprite & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NpcNum As Long)
Dim packet As String

    packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).sprite & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

'Sub SendUpdatePetTo(ByVal index As Long, ByVal PetId As Long)
'Dim packet As String
'
'    packet = "UPDATEPET" & SEP_CHAR & PetId & SEP_CHAR & Trim(Pets(PetId).Name) & SEP_CHAR & Pets(PetId).sprite & SEP_CHAR & Pets(PetId).x & SEP_CHAR & Pets(PetId).y & SEP_CHAR & Pets(PetId).map & SEP_CHAR & Pets(PetId).Dir & SEP_CHAR & END_CHAR
'    Call SendDataTo(index, packet)
'End Sub

'Sub SendUpdatePetToAll(ByVal PetId As Long)
'Dim packet As String
'
'    packet = "UPDATEPET" & SEP_CHAR & PetId & SEP_CHAR & Trim(Pets(PetId).Name) & SEP_CHAR & Pets(PetId).sprite & SEP_CHAR & Pets(PetId).x & SEP_CHAR & Pets(PetId).y & SEP_CHAR & Pets(PetId).map & SEP_CHAR & Pets(PetId).Dir & SEP_CHAR & END_CHAR
'    Call SendDataToAll(packet)
'End Sub

Sub SendEditNpcTo(ByVal index As Long, ByVal NpcNum As Long)
Dim packet As String

    packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).str & SEP_CHAR & Npc(NpcNum).def & SEP_CHAR & Npc(NpcNum).speed & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).HP & SEP_CHAR & Npc(NpcNum).ExpGiven & SEP_CHAR & Npc(NpcNum).Respawn & SEP_CHAR & Npc(NpcNum).QuestID & SEP_CHAR & Npc(NpcNum).opensBank & SEP_CHAR & Npc(NpcNum).opensShop & SEP_CHAR & Npc(NpcNum).type & SEP_CHAR & Npc(NpcNum).Attack_with_Poison & SEP_CHAR & Npc(NpcNum).Poison_length & SEP_CHAR & Npc(NpcNum).Poison_vital & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendShops(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Trim(Shop(i).Name) <> "" Then
            Call SendUpdateShopTo(index, i)
        End If
    Next i
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim packet As String

    packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum)
Dim packet As String

    packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendEditShopTo(ByVal index As Long, ByVal ShopNum As Long)
Dim packet As String
Dim i As Long

    packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        packet = packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
    Next i
    packet = packet & END_CHAR

    Call SendDataTo(index, packet)
End Sub

Sub SendSpells(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim(Spell(i).Name) <> "" Then
            Call SendUpdateSpellTo(index, i)
        End If
    Next i
End Sub

Sub SendPrayers(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim(Prayer(i).Name) <> "" Then
            Call SendUpdatePrayerTo(index, i)
        End If
    Next i
End Sub

Sub SendSound(ByVal index As Long, ByVal sound As String, Optional ByVal blnSpell As Boolean = False)
Dim packet As String
If blnSpell Then
    packet = "playsound" & SEP_CHAR & "magic" & sound & SEP_CHAR & END_CHAR
Else
    packet = "playsound" & SEP_CHAR & sound & SEP_CHAR & END_CHAR
End If
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim packet As String

    packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
Dim packet As String

    packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendUpdatePrayerTo(ByVal index As Long, ByVal PrayerNum As Long)
Dim packet As String

    packet = "UPDATEPRAYER" & SEP_CHAR & PrayerNum & SEP_CHAR & Trim(Prayer(PrayerNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendUpdatePrayerToAll(ByVal PrayerNum As Long)
Dim packet As String

    packet = "UPDATEPRAYER" & SEP_CHAR & PrayerNum & SEP_CHAR & Trim(Prayer(PrayerNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendUpdateQuestToAll(ByVal questNum As Long)
Dim packet As String

    packet = "UPDATEQUEST" & SEP_CHAR & questNum & SEP_CHAR & Trim(Quests(questNum).ID) & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)
End Sub

Sub SendEditSpellTo(ByVal index As Long, ByVal SpellNum As Long)
Dim packet As String

    packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).sound & SEP_CHAR & Spell(SpellNum).manaUse & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendEditPrayerTo(ByVal index As Long, ByVal PrayerNum As Long)
Dim packet As String

    packet = "EDITPRAYER" & SEP_CHAR & PrayerNum & SEP_CHAR & Trim(Prayer(PrayerNum).Name) & SEP_CHAR & Prayer(PrayerNum).ClassReq & SEP_CHAR & Prayer(PrayerNum).LevelReq & SEP_CHAR & Prayer(PrayerNum).type & SEP_CHAR & Prayer(PrayerNum).Data1 & SEP_CHAR & Prayer(PrayerNum).Data2 & SEP_CHAR & Prayer(PrayerNum).Data3 & SEP_CHAR & Prayer(PrayerNum).sound & SEP_CHAR & Prayer(PrayerNum).manaUse & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendEditQuestTo(ByVal index As Long, ByVal questNum As Long)
Dim packet As String

    packet = "EDITQUEST" & SEP_CHAR & questNum & SEP_CHAR & Trim(Quests(questNum).ExpGiven) & SEP_CHAR & Trim(Quests(questNum).FinishQuestMessage) & SEP_CHAR & Trim(Quests(questNum).GetItemQuestMsg) & SEP_CHAR & Trim(Quests(questNum).ItemGiven) & SEP_CHAR & Trim(Quests(questNum).ItemToObtain) & SEP_CHAR & Trim(Quests(questNum).ItemValGiven) & SEP_CHAR & Trim(Quests(questNum).requiredLevel) & SEP_CHAR & Trim(Quests(questNum).StartQuestMsg) & SEP_CHAR & Trim(Quests(questNum).GetItemQuestMsg) & SEP_CHAR & Trim(Quests(questNum).goldGiven) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub


Sub SendEditsignTo(ByVal index As Long, ByVal SignNum As Long)
Dim packet As String

    packet = "EDITSIGN" & SEP_CHAR & SignNum & SEP_CHAR & Trim(Signs(SignNum).header) & SEP_CHAR & Trim(Signs(SignNum).msg) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendTrade(ByVal index As Long, ByVal ShopNum As Long)
Dim packet As String
Dim i As Long, x As Long, y As Long

    packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        packet = packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
        
        ' Item #
        x = Shop(ShopNum).TradeItem(i).GetItem
        
        If Item(x).type = ITEM_TYPE_SPELL Then
            ' Spell class requirement
            y = Spell(Item(x).Data1).ClassReq
            
            If y = 0 Then
                Call PlayerMsg(index, Trim(Item(x).Name) & " can be used by all paths.", RGB_HelpColor)
            Else
                Call PlayerMsg(index, Trim(Item(x).Name) & " can only be used by a " & GetClassName(y - 1) & ".", RGB_HelpColor)
            End If
        End If
    Next i
    packet = packet & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub SendPlayerSpells(ByVal index As Long)
Dim packet As String
Dim i As Long

    packet = "SPELLS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SPELLS
        packet = packet & GetPlayerSpell(index, i) & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub SendPlayerPrayers(ByVal index As Long)
Dim packet As String
Dim i As Long

    packet = "PRAYERS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SPELLS
        packet = packet & GetPlayerPrayer(index, i) & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    
    Call SendDataTo(index, packet)
End Sub

Sub SendWeatherTo(ByVal index As Long)
Dim packet As String

    packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendWeatherToAll()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendWeatherTo(i)
        End If
    Next i
End Sub

Sub SendTimeTo(ByVal index As Long)
Dim packet As String

    packet = "TIME" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SendTimeToAll()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendTimeTo(i)
        End If
    Next i
End Sub

