Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = GAME_NAME & " Server <IP " & frmServer.Socket(0).LocalIP & " Port " & STR(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If IsConnected(index) And Player(index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Function IsLoggedIn(ByVal index As Long) As Boolean
    If IsConnected(index) And Trim(Player(index).Login) <> "" Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim i As Long

    IsMultiAccounts = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And LCase(Trim(Player(i).Login)) = LCase(Trim(Login)) Then
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

Sub SendDataTo(ByVal index As Long, ByVal Data As String)
Dim i As Long, n As Long, startc As Long

    If IsConnected(index) Then
        frmServer.Socket(index).SendData Data
        DoEvents
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

Sub SendDataToAllBut(ByVal index As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
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

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
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

Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim Text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
    Call CloseSocket(index)
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If
    
        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & ".")
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
End Sub

Sub HandleData(ByVal index As Long, ByVal Data As String)
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
        
    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
        
    ' :::::::::::::::::::::::
    ' :: Get players stats ::
    ' :::::::::::::::::::::::
    Dim Packet As String
    If LCase(Parse(0)) = "playershpmp" Then
        Packet = "playershpmp" & SEP_CHAR
            For i = 1 To MAX_PLAYERS
                Packet = Packet & GetPlayerHP(i) & SEP_CHAR & GetPlayerMaxHP(i) & SEP_CHAR
            Next i
        Packet = Packet & END_CHAR
        Call SendDataTo(index, Packet)
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
        
     ' :::::::::::::::::::::::::::
    ' :: Npc hp request packet ::
    ' :::::::::::::::::::::::::::
    'Dim Packet As String
    If LCase(Parse(0)) = "getnpchp" Then
    
    Packet = "npchp" & SEP_CHAR
        For i = 1 To MAX_MAP_NPCS
            n = MapNpc(GetPlayerMap(index), i).Num
            Packet = Packet & MapNpc(GetPlayerMap(index), i).HP & SEP_CHAR & GetNpcMaxHP(n) & SEP_CHAR
        Next i
    Packet = Packet & END_CHAR
    
        Call SendDataTo(index, Packet)
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Guilds Packet ::
    ' :::::::::::::::::::
    
    ' Change Access
    If LCase(Parse(0)) = "guildchangeaccess" Then
    ' Check the requirements.
    If FindPlayer(Parse(1)) = 0 Then
        Call PlayerMsg(index, "Player is offline", White)
        Exit Sub
    End If
    If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(index) Then
        Call PlayerMsg(index, "Player is not in your guild", Red)
        Exit Sub
    End If
    'Set the player's new access level
    Call SetPlayerGuildAccess(FindPlayer(Parse(1)), Parse(2))
    Call SendPlayerData(FindPlayer(Parse(1)))
    'Can have a message here if you'd like
    End If
    ' Disown
    If LCase(Parse(0)) = "guilddisown" Then
    ' Check if all the requirements
        If FindPlayer(Parse(1)) = 0 Then
        Call PlayerMsg(index, "Player is offline", White)
        Exit Sub
    End If
    If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(index) Then
        Call PlayerMsg(index, "Player is not in your guild", Red)
        Exit Sub
    End If
    If GetPlayerGuildAccess(FindPlayer(Parse(1))) > GetPlayerGuildAccess(index) Then
        Call PlayerMsg(index, "Player has a higher guild level than you.", Red)
        Exit Sub
    End If
    'Player checks out, take him out of the guild
    Call SetPlayerGuild(FindPlayer(Parse(1)), "")
    Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
    Call SendPlayerData(FindPlayer(Parse(1)))
    'Can have a message here if you'd like
    End If

    ' Leave Guild
    If LCase(Parse(0)) = "guildleave" Then
    ' Check if they can leave
    If GetPlayerGuild(index) = "" Then
        Call PlayerMsg(index, "You are not in a guild.", Red)
        Exit Sub
    End If
    Call SetPlayerGuild(index, "")
    Call SetPlayerGuildAccess(index, 0)
    Call SendPlayerData(index)
    End If
    
    ' Make A New Guild
    If LCase(Parse(0)) = "makeguild" Then
    ' Check if the Owner is Online
    If FindPlayer(Parse(1)) = 0 Then
        Call PlayerMsg(index, "Player is offline", White)
        Exit Sub
    End If
    ' Check if they are alredy in a guild
        If GetPlayerGuild(FindPlayer(Parse(1))) <> "" Then
        Call PlayerMsg(index, "Player is already in a guild", Red)
        Exit Sub
    End If
    ' If everything is ok then lets make the guild
    Call SetPlayerGuild(FindPlayer(Parse(1)), (Parse(2)))
    Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 3)
    Call SendPlayerData(FindPlayer(Parse(1)))
    End If
    
    
    
    ' Make A Member
    If LCase(Parse(0)) = "guildmember" Then
    ' Check if its possible to admit the member
    If FindPlayer(Parse(1)) = 0 Then
        Call PlayerMsg(index, "Player is offline", White)
        Exit Sub
    End If
    If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(index) Then
        Call PlayerMsg(index, "That player is not in your guild", Red)
        Exit Sub
    End If
    If GetPlayerGuildAccess(FindPlayer(Parse(1))) > 1 Then
        Call PlayerMsg(index, "That player has already been admitted", Red)
        Exit Sub
    End If
    'All has gone well, set the guild access to 1
    Call SetPlayerGuild(FindPlayer(Parse(1)), GetPlayerGuild(index))
    Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 1)
    Call SendPlayerData(FindPlayer(Parse(1)))
    'You can have a message here if you'd like.
    End If
    
    ' Make A Trainie
    If LCase(Parse(0)) = "guildtrainee" Then
    ' Check if its possible to induct member
    If FindPlayer(Parse(1)) = 0 Then
        Call PlayerMsg(index, "Player is offline", White)
        Exit Sub
    End If
    If GetPlayerGuild(FindPlayer(Parse(1))) <> "" Then
        Call PlayerMsg(index, "Player is already in a guild", Red)
        Exit Sub
    End If
    'It is possible, so set the guild to index's guild, and the access level to 0
    Call SetPlayerGuild(FindPlayer(Parse(1)), GetPlayerGuild(index))
    Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
    Call SendPlayerData(FindPlayer(Parse(1)))
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
                If Trim(Player(index).Char(i).Name) <> "" Then
                    Call DeleteName(Player(index).Char(i).Name)
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
            
            ' Prevent Dupeing
            For i = 1 To Len(Name)
                n = Asc(Mid(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(index, "Account duping is not allowed!")
                Exit Sub
                End If
            Next i
        
            ' Check versions
            If Val(Parse(3)) < CLIENT_MAJOR Or Val(Parse(4)) < CLIENT_MINOR Or Val(Parse(5)) < CLIENT_REVISION Then
                Call AlertMsg(index, "Version outdated, please visit " & Trim(GetVar(App.Path & "\Data.ini", "CONFIG", "WebSite")))
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
                
            ' Check Security Code's
            If Parse(6) = "jwehiehfojcvnvnsdinaoiwheoewyriusdyrflsdjncjkxzncisdughfusyfuapsipiuahfpaijnflkjnvjnuahguiryasbdlfkjblsahgfauygewuifaunfauf" And Parse(7) = "ksisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlasndowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas" And Parse(8) = "saiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi" And Val(Parse(9)) = "88978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672" Then
            ' Everything went ok
            Else
                Call AlertMsg(index, "Haha! Stop trying to hack!")
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
            Call SendDataTo(index, Packs)
    
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
        
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Then
                Call AlertMsg(index, "Character name must be at least three characters in length.")
                Exit Sub
            End If
            
            ' Prevent being me
            If LCase(Trim(Name)) = "Liam" Then
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
                Call HackingAttempt(index, "Invalid Class")
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
        
            ' Everything went ok, add the character
            Call AddChar(index, Name, Sex, Class, CharNum)
            Call SavePlayer(index)
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
                Player(index).CharNum = CharNum
                Call JoinGame(index)
            
                CharNum = Player(index).CharNum
                Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
                Call TextAdd(frmServer.txtText, GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".", True)
                Call UpdateCaption
                
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
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Say Text Modification")
                Exit Sub
            End If
        Next i
        
        Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " : " & Msg & "", PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " : " & Msg & "", SayColor)
        Call MapMsg2(GetPlayerMap(index), Msg, index)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "emotemsg" Then
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Emote Text Modification")
                Exit Sub
            End If
        Next i
        
        Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Msg, PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & Msg, EmoteColor)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "broadcastmsg" Then
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Broadcast Text Modification")
                Exit Sub
            End If
        Next i
        
        s = GetPlayerName(index) & ": " & Msg
        Call AddLog(s, PLAYER_LOG)
        Call GlobalMsg(s, BroadcastColor)
        Call TextAdd(frmServer.txtText, s, True)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "globalmsg" Then
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Global Text Modification")
                Exit Sub
            End If
        Next i
        
        If GetPlayerAccess(index) > 0 Then
            s = "(global) " & GetPlayerName(index) & ": " & Msg
            Call AddLog(s, ADMIN_LOG)
            Call GlobalMsg(s, GlobalColor)
            Call TextAdd(frmServer.txtText, s, True)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "adminmsg" Then
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Admin Text Modification")
                Exit Sub
            End If
        Next i
        
        If GetPlayerAccess(index) > 0 Then
            Call AddLog("(admin " & GetPlayerName(index) & ") " & Msg, ADMIN_LOG)
            Call AdminMsg("(admin " & GetPlayerName(index) & ") " & Msg, AdminColor)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "playermsg" Then
        MsgTo = FindPlayer(Parse(1))
        Msg = Parse(2)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Player Msg Text Modification")
                Exit Sub
            End If
        Next i
        
        ' Check if they are trying to talk to themselves
        If MsgTo <> index Then
            If MsgTo > 0 Then
                Call AddLog(GetPlayerName(index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
                Call PlayerMsg(MsgTo, GetPlayerName(index) & " tells you, '" & Msg & "'", TellColor)
                Call PlayerMsg(index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
            Else
                Call PlayerMsg(index, "Player is not online.", White)
            End If
        Else
            Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " begins to mumble to himself, what a wierdo...", PLAYER_LOG)
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " begins to mumble to himself, what a wierdo...", Green)
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playermove" And Player(index).GettingMap = NO Then
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
        If Player(index).CastedSpell = YES Then
            ' Check if they have already casted a spell, and if so we can't let them move
            If GetTickCount > Player(index).AttackTimer + 1000 Then
                Player(index).CastedSpell = NO
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
    If LCase(Parse(0)) = "playerdir" And Player(index).GettingMap = NO Then
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
        CharNum = Player(index).CharNum
        
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
        
        If (GetPlayerInvItemNum(index, InvNum) > 0) And (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
            n = Item(GetPlayerInvItemNum(index, InvNum)).Data2
            
            Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long, n5 As Long
            n1 = Item(GetPlayerInvItemNum(index, InvNum)).StrReq
            n2 = Item(GetPlayerInvItemNum(index, InvNum)).DefReq
            n3 = Item(GetPlayerInvItemNum(index, InvNum)).SpeedReq
            n4 = Item(GetPlayerInvItemNum(index, InvNum)).ClassReq
            n5 = Item(GetPlayerInvItemNum(index, InvNum)).AccessReq
            
            ' Find out what kind of item it is
            Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum <> GetPlayerArmorSlot(index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(index) <> n4 Then
                                Call PlayerMsg(index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Your strength is too low to equip this item!  Required Str (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerArmorSlot(index, InvNum)
                    Else
                        Call SetPlayerArmorSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)

                Case ITEM_TYPE_WEAPON
                    If InvNum <> GetPlayerWeaponSlot(index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(index) <> n4 Then
                                Call PlayerMsg(index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Your strength is too low to equip this item!  Required Str (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerWeaponSlot(index, InvNum)
                    Else
                        Call SetPlayerWeaponSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
                        
                Case ITEM_TYPE_HELMET
                    If InvNum <> GetPlayerHelmetSlot(index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(index) <> n4 Then
                                Call PlayerMsg(index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Your strength is too low to equip this item!  Required Str (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerHelmetSlot(index, InvNum)
                    Else
                        Call SetPlayerHelmetSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
            
                Case ITEM_TYPE_SHIELD
                    If InvNum <> GetPlayerShieldSlot(index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(index) <> n4 Then
                                Call PlayerMsg(index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Your strength is too low to equip this item!  Required Str (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerShieldSlot(index, InvNum)
                    Else
                        Call SetPlayerShieldSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
            
                Case ITEM_TYPE_POTIONADDHP
                    Call SetPlayerHP(index, GetPlayerHP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendHP(index)
                
                Case ITEM_TYPE_POTIONADDMP
                    Call SetPlayerMP(index, GetPlayerMP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendMP(index)
        
                Case ITEM_TYPE_POTIONADDSP
                    Call SetPlayerSP(index, GetPlayerSP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendSP(index)

                Case ITEM_TYPE_POTIONSUBHP
                    Call SetPlayerHP(index, GetPlayerHP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendHP(index)
                
                Case ITEM_TYPE_POTIONSUBMP
                    Call SetPlayerMP(index, GetPlayerMP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendMP(index)
        
                Case ITEM_TYPE_POTIONSUBSP
                    Call SetPlayerSP(index, GetPlayerSP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendSP(index)
                    
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
                    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY Then
                        ' Check if the key they are using matches the map key
                        If GetPlayerInvItemNum(index, InvNum) = Map(GetPlayerMap(index)).Tile(x, y).Data1 Then
                            TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                            TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                            
                            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            If Trim(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) = "" Then
                                Call MapMsg(GetPlayerMap(index), "A door has been unlocked!", White)
                            Else
                                Call MapMsg(GetPlayerMap(index), Trim(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), White)
                            End If
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
                            
                            ' Check if we are supposed to take away the item
                            If Map(GetPlayerMap(index)).Tile(x, y).Data2 = 1 Then
                                Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                Call PlayerMsg(index, "The key disolves.", Yellow)
                            End If
                        End If
                    End If
                    
                    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_DOOR Then
                        TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                        TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                        
                        Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
                    End If
                    
                Case ITEM_TYPE_SPELL
                    ' Get the spell num
                    n = Item(GetPlayerInvItemNum(index, InvNum)).Data1
                    
                    If n > 0 Then
                        ' Make sure they are the right class
                        If Spell(n).ClassReq - 1 = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                            If Spell(n).LevelReq = 0 And Player(index).Char(Player(index).CharNum).Access < 1 Then
                                Call PlayerMsg(index, "This spell can only be used by admins!", BrightRed)
                                Exit Sub
                            End If
                            
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
                                        Call PlayerMsg(index, "You study the spell carefully...", Yellow)
                                        Call PlayerMsg(index, "You have learned a new spell!", White)
                                    Else
                                        Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                        Call PlayerMsg(index, "You have already learned this spell!  The spells crumbles into dust.", BrightRed)
                                    End If
                                Else
                                    Call PlayerMsg(index, "You have learned all that you can learn!", BrightRed)
                                End If
                            Else
                                Call PlayerMsg(index, "You must be level " & i & " to learn this spell.", White)
                            End If
                        Else
                            Call PlayerMsg(index, "This spell can only be learned by a " & GetClassName(Spell(n).ClassReq - 1) & ".", White)
                        End If
                    Else
                        Call PlayerMsg(index, "This scroll is not connected to a spell, please inform an admin!", White)
                    End If
            End Select
            
            Call SendStats(index)
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
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
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                        Else
                            n = GetPlayerDamage(index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                            Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                            Call PlayerMsg(i, GetPlayerName(index) & " swings with enormous might!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(index, i, Damage)
                        Else
                            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call PlayerMsg(index, GetPlayerName(i) & "'s " & Trim(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked your hit!", BrightCyan)
                        Call PlayerMsg(i, "Your " & Trim(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
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
                    Damage = GetPlayerDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).Num).DEF / 2)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                Else
                    n = GetPlayerDamage(index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).Num).DEF / 2)
                    Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(index, i, Damage)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                Else
                    Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
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
        If (PointType < 0) Or (PointType > 3) Then
            Call HackingAttempt(index, "Invalid Point Type")
            Exit Sub
        End If
                
        ' Make sure they have points
        If GetPlayerPOINTS(index) > 0 Then
            ' Take away a stat point
            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "UsingStatPoints " & index & "," & PointType
            Else
                Select Case PointType
                    Case 0
                        Call SetPlayerSTR(index, GetPlayerSTR(index) + 1)
                        Call PlayerMsg(index, "You have gained more strength!", 15)
                    Case 1
                        Call SetPlayerDEF(index, GetPlayerDEF(index) + 1)
                        Call PlayerMsg(index, "You have gained more defense!", 15)
                    Case 2
                        Call SetPlayerMAGI(index, GetPlayerMAGI(index) + 1)
                        Call PlayerMsg(index, "You have gained more magic abilities!", 15)
                    Case 3
                        Call SetPlayerSPEED(index, GetPlayerSPEED(index) + 1)
                        Call PlayerMsg(index, "You have gained more speed!", 15)
                End Select
            End If
        Else
            Call PlayerMsg(index, "You have no skill points to train with!", BrightRed)
        End If
                
        ' Send the update
        Call SendStats(index)
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::::::::::::
    ' :: Player info request packet ::
    ' ::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinforequest" Then
        Name = Parse(1)
        
        i = FindPlayer(Name)
        If i > 0 Then
            Call PlayerMsg(index, "Account: " & Trim(Player(i).Login) & ", Name: " & GetPlayerName(i), BrightGreen)
            If GetPlayerAccess(index) > ADMIN_MONITER Then
                Call PlayerMsg(index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
                Call PlayerMsg(index, "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), BrightGreen)
                Call PlayerMsg(index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
                Call PlayerMsg(index, "STR: " & GetPlayerSTR(i) & "  DEF: " & GetPlayerDEF(i) & "  MAGI: " & GetPlayerMAGI(i) & "  SPEED: " & GetPlayerSPEED(i), BrightGreen)
                n = Int(GetPlayerSTR(i) / 2) + Int(GetPlayerLevel(i) / 2)
                i = Int(GetPlayerDEF(i) / 2) + Int(GetPlayerLevel(i) / 2)
                If n > 100 Then n = 100
                If i > 100 Then i = 100
                Call PlayerMsg(index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", BrightGreen)
            End If
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Warp me to packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "warpmeto" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        n = FindPlayer(Parse(1))
        
        If n <> index Then
            If n > 0 Then
                Call PlayerWarp(index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
                Call PlayerMsg(n, GetPlayerName(index) & " has warped to you.", BrightBlue)
                Call PlayerMsg(index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
                Call AddLog(GetPlayerName(index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(index, "You cannot warp to yourself!", White)
        End If
                
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::
    ' :: Warp to me packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "warptome" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        n = FindPlayer(Parse(1))
        
        If n <> index Then
            If n > 0 Then
                Call PlayerWarp(n, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(index) & ".", BrightBlue)
                Call PlayerMsg(index, GetPlayerName(n) & " has been summoned.", BrightBlue)
                Call AddLog(GetPlayerName(index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(index) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(index, "You cannot warp yourself to yourself!", White)
        End If
        
        Exit Sub
    End If


    ' ::::::::::::::::::::::::
    ' :: Warp to map packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "warpto" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The map
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_MAPS Then
            Call HackingAttempt(index, "Invalid map")
            Exit Sub
        End If
        
        Call PlayerWarp(index, n, GetPlayerX(index), GetPlayerY(index))
        Call PlayerMsg(index, "You have been warped to map #" & n, BrightBlue)
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
                
    ' ::::::::::::::::::::::::::::::
    ' :: Set player sprite packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "setplayersprite" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The sprite
        i = FindPlayer(Parse(1))
        n = Val(Parse(2))
                
        Call SetPlayerSprite(i, n)
        Call SendPlayerData(i)
        Exit Sub
    End If
                
    ' ::::::::::::::::::::::::::
    ' :: Stats request packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "getstats" Then
        Call PlayerMsg(index, "-=- Stats for " & GetPlayerName(index) & " -=-", White)
        Call PlayerMsg(index, "Level: " & GetPlayerLevel(index) & "  Exp: " & GetPlayerExp(index) & "/" & GetPlayerNextLevel(index), White)
        Call PlayerMsg(index, "HP: " & GetPlayerHP(index) & "/" & GetPlayerMaxHP(index) & "  MP: " & GetPlayerMP(index) & "/" & GetPlayerMaxMP(index) & "  SP: " & GetPlayerSP(index) & "/" & GetPlayerMaxSP(index), White)
        Call PlayerMsg(index, "STR: " & GetPlayerSTR(index) & "  DEF: " & GetPlayerDEF(index) & "  MAGI: " & GetPlayerMAGI(index) & "  SPEED: " & GetPlayerSPEED(index), White)
        n = Int(GetPlayerSTR(index) / 2) + Int(GetPlayerLevel(index) / 2)
        i = Int(GetPlayerDEF(index) / 2) + Int(GetPlayerLevel(index) / 2)
        If n > 100 Then n = 100
        If i > 100 Then i = 100
        Call PlayerMsg(index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", White)
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
    If LCase(Parse(0)) = "mapdata" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = 1
        
        MapNum = GetPlayerMap(index)
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
        
        n = n + 12
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
            Map(MapNum).Tile(x, y).Ground = Val(Parse(n))
            Map(MapNum).Tile(x, y).Mask = Val(Parse(n + 1))
            Map(MapNum).Tile(x, y).Anim = Val(Parse(n + 2))
            Map(MapNum).Tile(x, y).Mask2 = Val(Parse(n + 3))
            Map(MapNum).Tile(x, y).M2Anim = Val(Parse(n + 4))
            Map(MapNum).Tile(x, y).Fringe = Val(Parse(n + 5))
            Map(MapNum).Tile(x, y).FAnim = Val(Parse(n + 6))
            Map(MapNum).Tile(x, y).Fringe2 = Val(Parse(n + 7))
            Map(MapNum).Tile(x, y).F2Anim = Val(Parse(n + 8))
            Map(MapNum).Tile(x, y).Type = Val(Parse(n + 9))
            Map(MapNum).Tile(x, y).Data1 = Val(Parse(n + 10))
            Map(MapNum).Tile(x, y).Data2 = Val(Parse(n + 11))
            Map(MapNum).Tile(x, y).Data3 = Val(Parse(n + 12))
            Map(MapNum).Tile(x, y).String1 = Parse(n + 13)
            Map(MapNum).Tile(x, y).String2 = Parse(n + 14)
            Map(MapNum).Tile(x, y).String3 = Parse(n + 15)

            n = n + 16
            Next x
        Next y
       
        For x = 1 To MAX_MAP_NPCS
            Map(MapNum).Npc(x) = Val(Parse(n))
            n = n + 1
            Call ClearMapNpc(x, MapNum)
        Next x
        
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
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Need map yes/no packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "needmap" Then
        ' Get yes/no value
        s = LCase(Parse(1))
                
        If s = "yes" Then
            Call SendMap(index, GetPlayerMap(index))
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            Player(index).GettingMap = NO
            Call SendDataTo(index, "MAPDONE" & SEP_CHAR & END_CHAR)
        Else
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            Player(index).GettingMap = NO
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
        If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
            If Ammount <= 0 Then
                Call HackingAttempt(index, "Trying to drop 0 ammount of currency")
                Exit Sub
            End If
        End If
            
        Call PlayerMapDropItem(index, InvNum, Ammount)
        Call SendStats(index)
        Call SendHP(index)
        Call SendMP(index)
        Call SendSP(index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Respawn map packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "maprespawn" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
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
        
        Call PlayerMsg(index, "Map respawned.", Blue)
        Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Map report packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "mapreport2" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        s = "Free Maps: "
        tMapStart = 1
        tMapEnd = 1
        
        For i = 1 To MAX_MAPS
            If Trim(Map(i).Name) = "" Then
                tMapEnd = tMapEnd + 1
            Else
                If tMapEnd - tMapStart > 0 Then
                    s = s & Trim(STR(tMapStart)) & "-" & Trim(STR(tMapEnd - 1)) & ", "
                End If
                tMapStart = i + 1
                tMapEnd = i + 1
            End If
        Next i
        
        s = s & Trim(STR(tMapStart)) & "-" & Trim(STR(tMapEnd - 1)) & ", "
        s = Mid(s, 1, Len(s) - 2)
        s = s & "."
        
        Call PlayerMsg(index, s, Brown)
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
                    Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(index) & "!", White)
                    Call AddLog(GetPlayerName(index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                    Call AlertMsg(n, "You have been kicked by " & GetPlayerName(index) & "!")
                Else
                    Call PlayerMsg(index, "That is a higher access admin then you!", White)
                End If
            Else
                Call PlayerMsg(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(index, "You cannot kick yourself!", White)
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
            
            Call PlayerMsg(index, n & ": Banned IP " & s & " by " & Name, White)
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
        Call PlayerMsg(index, "Ban list destroyed.", White)
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
                    Call PlayerMsg(index, "That is a higher access admin then you!", White)
                End If
            Else
                Call PlayerMsg(index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(index, "You cannot ban yourself!", White)
        End If
                
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Request edit map packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditmap" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "EDITMAP" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit item packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requestedititem" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "ITEMEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit item packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "edititem" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Item(n).StrReq = Val(Parse(8))
        Item(n).DefReq = Val(Parse(9))
        Item(n).SpeedReq = Val(Parse(10))
        Item(n).ClassReq = Val(Parse(11))
        Item(n).AccessReq = Val(Parse(12))
        
        Item(n).AddHP = Val(Parse(13))
        Item(n).AddMP = Val(Parse(14))
        Item(n).AddSP = Val(Parse(15))
        Item(n).AddStr = Val(Parse(16))
        Item(n).AddDef = Val(Parse(17))
        Item(n).AddMagi = Val(Parse(18))
        Item(n).AddSpeed = Val(Parse(19))
        Item(n).AddEXP = Val(Parse(20))
        
        Item(n).Desc = Parse(21)
        
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
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
        Npc(n).Sprite = Val(Parse(4))
        Npc(n).SpawnSecs = Val(Parse(5))
        Npc(n).Behavior = Val(Parse(6))
        Npc(n).Range = Val(Parse(7))
        Npc(n).STR = Val(Parse(8))
        Npc(n).DEF = Val(Parse(9))
        Npc(n).SPEED = Val(Parse(10))
        Npc(n).MAGI = Val(Parse(11))
        Npc(n).Big = Val(Parse(12))
        Npc(n).MaxHp = Val(Parse(13))
        Npc(n).Exp = Val(Parse(14))
        
        z = 15
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = Val(Parse(z))
            Npc(n).ItemNPC(i).ItemNum = Val(Parse(z + 1))
            Npc(n).ItemNPC(i).ItemValue = Val(Parse(z + 2))
            z = z + 3
        Next i
        
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
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "SPELLEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "editspell" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
    
    ' :::::::::::::::::::::::
    ' :: Save spell packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "savespell") Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
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
        Spell(n).Type = Val(Parse(5))
        Spell(n).Data1 = Val(Parse(6))
        Spell(n).Data2 = Val(Parse(7))
        Spell(n).Data3 = Val(Parse(8))
        Spell(n).MPCost = Val(Parse(9))
        Spell(n).Sound = Val(Parse(10))
                
        ' Save it
        Call SendUpdateSpellToAll(n)
        Call SaveSpell(n)
        Call AddLog(GetPlayerName(index) & " saving spell #" & n & ".", ADMIN_LOG)
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
            If GetPlayerName(index) <> GetPlayerName(n) Then
                If GetPlayerAccess(index) > GetPlayerAccess(n) Then
                    ' Check if player is on
                    If n > 0 Then
                        If GetPlayerAccess(n) <= 0 Then
                            Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
                        End If
                    
                        Call SetPlayerAccess(n, i)
                        Call SendPlayerData(n)
                        Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
                    Else
                        Call PlayerMsg(index, "Player is not online.", White)
                    End If
                Else
                    Call PlayerMsg(index, "Your access level is lower than " & GetPlayerName(n) & "´s.", Red)
                End If
            Else
                Call PlayerMsg(index, "You cant change your access.", Red)
            End If
        Else
            Call PlayerMsg(index, "Invalid access level.", Red)
        End If
                
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Whos online packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "whosonline" Then
        Call SendWhosOnline(index)
        Exit Sub
    End If
    
    ' :::::::::::::::::
    ' :: Online list ::
    ' :::::::::::::::::
    If LCase(Parse(0)) = "onlinelist" Then
        'Stop
        Call SendOnlineList
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
        Call GlobalMsg("MOTD changed to: " & Parse(1), BrightCyan)
        Call AddLog(GetPlayerName(index) & " changed MOTD to: " & Parse(1), ADMIN_LOG)
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
        i = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        
        ' Check if inv full
        x = FindOpenInvSlot(index, Shop(i).TradeItem(n).GetItem)
        If x = 0 Then
            Call PlayerMsg(index, "Trade unsuccessful, inventory full.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have the item
        If HasItem(index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
            Call TakeItem(index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
            Call GiveItem(index, Shop(i).TradeItem(n).GetItem, Shop(i).TradeItem(n).GetValue)
            Call PlayerMsg(index, "The trade was successful!", Yellow)
        Else
            Call PlayerMsg(index, "Trade unsuccessful.", BrightRed)
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
        If Item(GetPlayerInvItemNum(index, n)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(index, n)).Type > ITEM_TYPE_SHIELD Then
            Call PlayerMsg(index, "You can only fix weapons, armors, helmets, and shields.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have a full inventory
        If FindOpenInvSlot(index, GetPlayerInvItemNum(index, n)) <= 0 Then
            Call PlayerMsg(index, "You have no inventory space left!", BrightRed)
            Exit Sub
        End If
        
        ' Now check the rate of pay
        ItemNum = GetPlayerInvItemNum(index, n)
        i = Int(Item(GetPlayerInvItemNum(index, n)).Data2 / 5)
        If i <= 0 Then i = 1
        
        DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(index, n)
        GoldNeeded = Int(DurNeeded * i / 2)
        If GoldNeeded <= 0 Then GoldNeeded = 1
        
        ' Check if they even need it repaired
        If DurNeeded <= 0 Then
            Call PlayerMsg(index, "This item is in perfect condition!", White)
            Exit Sub
        End If
        
        ' Check if they have enough for at least one point
        If HasItem(index, 1) >= i Then
            ' Check if they have enough for a total restoration
            If HasItem(index, 1) >= GoldNeeded Then
                Call TakeItem(index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(index, n, Item(ItemNum).Data1)
                Call PlayerMsg(index, "Item has been totally restored for " & GoldNeeded & " gold!", BrightBlue)
            Else
                ' They dont so restore as much as we can
                DurNeeded = (HasItem(index, 1) / i)
                GoldNeeded = Int(DurNeeded * i / 2)
                If GoldNeeded <= 0 Then GoldNeeded = 1
                
                Call TakeItem(index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(index, n, GetPlayerInvItemDur(index, n) + DurNeeded)
                Call PlayerMsg(index, "Item has been partially fixed for " & GoldNeeded & " gold!", BrightBlue)
            End If
        Else
            Call PlayerMsg(index, "Insufficient gold to fix this item!", BrightRed)
        End If
        Exit Sub
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
                    Call PlayerMsg(index, "You wouldn't stand a chance.", BrightRed)
                Else
                    If GetPlayerLevel(i) > GetPlayerLevel(index) Then
                        Call PlayerMsg(index, "This one seems to have an advantage over you.", Yellow)
                    Else
                        If GetPlayerLevel(i) = GetPlayerLevel(index) Then
                            Call PlayerMsg(index, "This would be an even fight.", White)
                        Else
                            If GetPlayerLevel(index) >= GetPlayerLevel(i) + 5 Then
                                Call PlayerMsg(index, "You could slaughter that player.", BrightBlue)
                            Else
                                If GetPlayerLevel(index) > GetPlayerLevel(i) Then
                                    Call PlayerMsg(index, "You would have an advantage over that player.", Yellow)
                                End If
                            End If
                        End If
                    End If
                End If
            
                ' Change target
                Player(index).Target = i
                Player(index).TargetType = TARGET_TYPE_PLAYER
                Call PlayerMsg(index, "Your target is now " & GetPlayerName(i) & ".", Yellow)
                Exit Sub
            End If
        Next i
        
        ' Check for an item
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(GetPlayerMap(index), i).Num > 0 Then
                If MapItem(GetPlayerMap(index), i).x = x And MapItem(GetPlayerMap(index), i).y = y Then
                    Call PlayerMsg(index, "You see a " & Trim(Item(MapItem(GetPlayerMap(index), i).Num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next i
        
        ' Check for an npc
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(index), i).Num > 0 Then
                If MapNpc(GetPlayerMap(index), i).x = x And MapNpc(GetPlayerMap(index), i).y = y Then
                    ' Change target
                    Player(index).Target = i
                    Player(index).TargetType = TARGET_TYPE_NPC
                    Call PlayerMsg(index, "Your target is now a " & Trim(Npc(MapNpc(GetPlayerMap(index), i).Num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next i
        
        Exit Sub
    End If
    
' ::::::::::::::::::::::::::::::::
' :: Player Chat System Packets ::
' ::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerchat" Then
        n = FindPlayer(Parse(1))
        If n < 1 Then
            Call PlayerMsg(index, "Player is not online.", White)
            Exit Sub
        End If
        If n = index Then
            Exit Sub
        End If
        If Player(index).InChat = 1 Then
            Call PlayerMsg(index, "Your already in a chat with another player!", Pink)
            Exit Sub
        End If

        If Player(n).InChat = 1 Then
            Call PlayerMsg(index, "Player is already in a chat with another player!", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Chat request has been sent to " & GetPlayerName(n) & ".", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " wants you to chat with them.  Type /chat to accept, or /chatdecline to decline.", Pink)
    
        Player(n).ChatPlayer = index
        Player(index).ChatPlayer = n
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "achat" Then
        n = Player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If
        If Player(n).ChatPlayer <> index Then
            Call PlayerMsg(index, "Chat failed.", Pink)
            Exit Sub
        End If
                        
        Call SendDataTo(index, "PPCHATTING" & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "PPCHATTING" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "dchat" Then
        n = Player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Declined chat request.", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " declined your request.", Pink)
        
        Player(index).ChatPlayer = 0
        Player(index).InChat = 0
        Player(n).ChatPlayer = 0
        Player(n).InChat = 0
        Exit Sub
    End If

    If LCase(Parse(0)) = "qchat" Then
        n = Player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If

        Call SendDataTo(index, "qchat" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "qchat" & SEP_CHAR & END_CHAR)
        
        Player(index).ChatPlayer = 0
        Player(index).InChat = 0
        Player(n).ChatPlayer = 0
        Player(n).InChat = 0
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "sendchat" Then
        n = Player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If

        Call SendDataTo(n, "sendchat" & SEP_CHAR & Parse(1) & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
' ::::::::::::::::::::::::::::::::::::
' :: END Player Chat System Packets ::
' ::::::::::::::::::::::::::::::::::::
    
    ' ::::::::::::::::::::::
    ' :: P2P Trade Packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "pptrade" Then
        n = FindPlayer(Parse(1))
        
        ' Check if player is online
        If n < 1 Then
            Call PlayerMsg(index, "Player is not online.", White)
            Exit Sub
        End If
        
        ' Prevent trading with self
        If n = index Then
            Exit Sub
        End If
                
        ' Check if the player is in another trade
        If Player(index).InTrade = 1 Then
            Call PlayerMsg(index, "Your already in a trade with someone else!", Pink)
            Exit Sub
        End If
        
        ' Check where both players are
        If GetPlayerX(n) + 1 <> GetPlayerX(index) Xor GetPlayerY(n) = GetPlayerY(index) Then
            If GetPlayerX(n) - 1 <> GetPlayerX(index) Xor GetPlayerY(n) = GetPlayerY(index) Then
                If GetPlayerX(n) = GetPlayerX(index) Xor GetPlayerY(n) + 1 <> GetPlayerY(index) Then
                    If GetPlayerX(n) = GetPlayerX(index) Xor GetPlayerY(n) - 1 <> GetPlayerY(index) Then
                        Call PlayerMsg(index, "The player needs to be beside you to trade!", Pink)
                        Call PlayerMsg(n, "You need to be beside the player to trade!", Pink)
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        ' Check to see if player is already in a trade
        If Player(n).InTrade = 1 Then
            Call PlayerMsg(index, "Player is already in a trade!", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Trade request has been sent to " & GetPlayerName(n) & ".", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " wants you to trade with them.  Type /accept to accept, or /decline to decline.", Pink)
    
        Player(n).TradePlayer = index
        Player(index).TradePlayer = n
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Accept trade packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "atrade" Then
        n = Player(index).TradePlayer
        
        ' Check if anyone requested a trade
        If n < 1 Then
            Call PlayerMsg(index, "No one requested a trade with you.", Pink)
            Exit Sub
        End If
        
        ' Check if its the right player
        If Player(n).TradePlayer <> index Then
            Call PlayerMsg(index, "Trade failed.", Pink)
            Exit Sub
        End If
        
        ' Check where both players are
        If GetPlayerX(n) + 1 <> GetPlayerX(index) Xor GetPlayerY(n) = GetPlayerY(index) Then
            If GetPlayerX(n) - 1 <> GetPlayerX(index) Xor GetPlayerY(n) = GetPlayerY(index) Then
                If GetPlayerX(n) = GetPlayerX(index) Xor GetPlayerY(n) + 1 <> GetPlayerY(index) Then
                    If GetPlayerX(n) = GetPlayerX(index) Xor GetPlayerY(n) - 1 <> GetPlayerY(index) Then
                        Call PlayerMsg(index, "The player needs to be beside you to trade!", Pink)
                        Call PlayerMsg(n, "You need to be beside the player to trade!", Pink)
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        Call PlayerMsg(index, "You are trading with " & GetPlayerName(n) & "!", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " accepted your trade request!", Pink)
        
        Call SendDataTo(index, "PPTRADING" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "PPTRADING" & SEP_CHAR & END_CHAR)
        
        For i = 1 To MAX_PLAYER_TRADES
            Player(index).Trading(i).InvNum = 0
            Player(index).Trading(i).InvName = ""
            Player(n).Trading(i).InvNum = 0
            Player(n).Trading(i).InvName = ""
        Next i
        
        Player(index).InTrade = 1
        Player(index).TradeItemMax = 0
        Player(index).TradeItemMax2 = 0
        Player(n).InTrade = 1
        Player(n).TradeItemMax = 0
        Player(n).TradeItemMax2 = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Stop trade packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "qtrade" Then
        n = Player(index).TradePlayer
        
        ' Check if anyone trade with player
        If n < 1 Then
            Call PlayerMsg(index, "No one requested a trade with you.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Stopped trading.", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " stopped trading with you!", Pink)
        Call SendDataTo(index, "qtrade" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "qtrade" & SEP_CHAR & END_CHAR)
        
        Player(index).TradePlayer = 0
        Player(index).InTrade = 0
        Player(n).TradePlayer = 0
        Player(n).InTrade = 0
        Call SendDataTo(index, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Decline trade packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "dtrade" Then
        n = Player(index).TradePlayer
        
        ' Check if anyone trade with player
        If n < 1 Then
            Call PlayerMsg(index, "No one requested a trade with you.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Declined trade request.", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " declined your request.", Pink)
        
        Player(index).TradePlayer = 0
        Player(index).InTrade = 0
        Player(n).TradePlayer = 0
        Player(n).InTrade = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Update trade packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "updatetradeinv" Then
        n = Val(Parse(1))
    
        Player(index).Trading(n).InvNum = Val(Parse(2))
        Player(index).Trading(n).InvName = Trim(Parse(3))
        If Player(index).Trading(n).InvNum = 0 Then
            Player(index).TradeItemMax = Player(index).TradeItemMax - 1
            Player(index).TradeOk = 0
            Player(n).TradeOk = 0
            Call SendDataTo(index, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Else
            Player(index).TradeItemMax = Player(index).TradeItemMax + 1
        End If
                
        Call SendDataTo(Player(index).TradePlayer, "updatetradeitem" & SEP_CHAR & n & SEP_CHAR & Player(index).Trading(n).InvNum & SEP_CHAR & Player(index).Trading(n).InvName & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Swap trade items packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "swapitems" Then

        n = Player(index).TradePlayer
        
        If Player(index).TradeOk = 0 Then
            Player(index).TradeOk = 1
            Call SendDataTo(n, "trading" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
        ElseIf Player(index).TradeOk = 1 Then
            Player(index).TradeOk = 0
            Call SendDataTo(n, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        End If
                
        If Player(index).TradeOk = 1 And Player(n).TradeOk = 1 Then
            Player(index).TradeItemMax2 = 0
            Player(n).TradeItemMax2 = 0

            For i = 1 To MAX_INV
                If Player(index).TradeItemMax = Player(index).TradeItemMax2 Then
                    Exit For
                End If
                If GetPlayerInvItemNum(n, i) < 1 Then
                    Player(index).TradeItemMax2 = Player(index).TradeItemMax2 + 1
                End If
            Next i

            For i = 1 To MAX_INV
                If Player(n).TradeItemMax = Player(n).TradeItemMax2 Then
                    Exit For
                End If
                If GetPlayerInvItemNum(index, i) < 1 Then
                    Player(n).TradeItemMax2 = Player(n).TradeItemMax2 + 1
                End If
            Next i
            
            If Player(index).TradeItemMax2 = Player(index).TradeItemMax And Player(n).TradeItemMax2 = Player(n).TradeItemMax Then
                For i = 1 To Player(index).TradeItemMax
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(n, x) < 1 Then
                            Call GiveItem(n, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), 1)
                            Call TakeItem(index, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), 1)
                            Exit For
                        End If
                    Next x
                Next i

                For i = 1 To Player(n).TradeItemMax
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(index, x) < 1 Then
                            Call GiveItem(index, GetPlayerInvItemNum(n, Player(index).Trading(i).InvNum), 1)
                            Call TakeItem(n, GetPlayerInvItemNum(n, Player(index).Trading(i).InvNum), 1)
                            Exit For
                        End If
                    Next x
                Next i

                Call PlayerMsg(n, "Trade Successfull!", BrightGreen)
                Call PlayerMsg(index, "Trade Successfull!", BrightGreen)
                Call SendInventory(n)
                Call SendInventory(index)
            Else
                If Player(index).TradeItemMax2 < Player(index).TradeItemMax Then
                    Call PlayerMsg(index, "Your inventory is full!", BrightRed)
                    Call PlayerMsg(n, GetPlayerName(index) & "'s inventory is full!", BrightRed)
                End If
                If Player(n).TradeItemMax2 < Player(n).TradeItemMax Then
                    Call PlayerMsg(n, "Your inventory is full!", BrightRed)
                    Call PlayerMsg(index, GetPlayerName(n) & "'s inventory is full!", BrightRed)
                End If
            End If
            
            Player(index).TradePlayer = 0
            Player(index).InTrade = 0
            Player(n).TradePlayer = 0
            Player(n).InTrade = 0
            
            Call SendDataTo(index, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Call SendDataTo(index, "qtrade" & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, "qtrade" & SEP_CHAR & END_CHAR)
        End If
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
        If Player(index).InParty = YES Then
            Call PlayerMsg(index, "You are already in a party!", Pink)
            Exit Sub
        End If
        
        If n > 0 Then
            ' Check if its an admin
            If GetPlayerAccess(index) > ADMIN_MONITER Then
                Call PlayerMsg(index, "You can't join a party, you are an admin!", BrightBlue)
                Exit Sub
            End If
        
            If GetPlayerAccess(n) > ADMIN_MONITER Then
                Call PlayerMsg(index, "Admins cannot join parties!", BrightBlue)
                Exit Sub
            End If
            
            ' Make sure they are in right level range
            If GetPlayerLevel(index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(index) - 5 > GetPlayerLevel(n) Then
                Call PlayerMsg(index, "There is more then a 5 level gap between you two, party failed.", Pink)
                Exit Sub
            End If
            
            ' Check to see if player is already in a party
            If Player(n).InParty = NO Then
                Call PlayerMsg(index, "Party request has been sent to " & GetPlayerName(n) & ".", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " wants you to join their party.  Type /join to join, or /leave to decline.", Pink)
            
                Player(index).PartyStarter = YES
                Player(index).PartyPlayer = n
                Player(n).PartyPlayer = index
            Else
                Call PlayerMsg(index, "Player is already in a party!", Pink)
            End If
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Join party packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "joinparty" Then
        n = Player(index).PartyPlayer
        
        If n > 0 Then
            ' Check to make sure they aren't the starter
            If Player(index).PartyStarter = NO Then
                ' Check to make sure that each of there party players match
                If Player(n).PartyPlayer = index Then
                    Call PlayerMsg(index, "You have joined " & GetPlayerName(n) & "'s party!", Pink)
                    Call PlayerMsg(n, GetPlayerName(index) & " has joined your party!", Pink)
                    
                    Player(index).InParty = YES
                    Player(n).InParty = YES
                Else
                    Call PlayerMsg(index, "Party failed.", Pink)
                End If
            Else
                Call PlayerMsg(index, "You have not been invited to join a party!", Pink)
            End If
        Else
            Call PlayerMsg(index, "You have not been invited into a party!", Pink)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Leave party packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "leaveparty" Then
        n = Player(index).PartyPlayer
        
        If n > 0 Then
            If Player(index).InParty = YES Then
                Call PlayerMsg(index, "You have left the party.", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " has left the party.", Pink)
                
                Player(index).PartyPlayer = 0
                Player(index).PartyStarter = NO
                Player(index).InParty = NO
                Player(n).PartyPlayer = 0
                Player(n).PartyStarter = NO
                Player(n).InParty = NO
            Else
                Call PlayerMsg(index, "Declined party request.", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " declined your request.", Pink)
                
                Player(index).PartyPlayer = 0
                Player(index).PartyStarter = NO
                Player(index).InParty = NO
                Player(n).PartyPlayer = 0
                Player(n).PartyStarter = NO
                Player(n).InParty = NO
            End If
        Else
            Call PlayerMsg(index, "You are not in a party!", Pink)
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
    
    ' :::::::::::::::::
    ' :: Cast packet ::
    ' :::::::::::::::::
    If LCase(Parse(0)) = "cast" Then
        ' Spell slot
        n = Val(Parse(1))
        
        Call CastSpell(index, n)
        
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Location packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "requestlocation" Then
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Map: " & GetPlayerMap(index) & ", X: " & GetPlayerX(index) & ", Y: " & GetPlayerY(index), Pink)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Refresh Player Packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "refresh" Then
        Call SendJoinMap(index)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Refresh Player Packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "buysprite" Then
        ' Check if player stepped on sprite changing tile
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
            Call PlayerMsg(index, "You need to be on a sprite tile to buy it!", BrightRed)
            Exit Sub
        End If
        
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 = 0 Then
            Call PlayerMsg(index, "You have bought a new sprite!", BrightGreen)
            Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
            Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 Then
                If Item(GetPlayerInvItemNum(index, i)).Type = ITEM_TYPE_CURRENCY Then
                    If GetPlayerInvItemValue(index, i) >= Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3 Then
                        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3)
                        If GetPlayerInvItemValue(index, i) <= 0 Then
                            Call SetPlayerInvItemNum(index, i, 0)
                        End If
                        Call PlayerMsg(index, "You have bought a new sprite!", BrightGreen)
                        Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                        Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(index)
                    End If
                Else
                    If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i Then
                        Call SetPlayerInvItemNum(index, i, 0)
                        Call PlayerMsg(index, "You have bought a new sprite!", BrightGreen)
                        Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                        Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(index)
                    End If
                End If
                If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i Then
                    Exit Sub
                End If
            End If
        Next i
        
        Call PlayerMsg(index, "You dont have enough to buy this sprite!", BrightRed)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Call the admins packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "calladmins" Then
        If GetPlayerAccess(index) = 0 Then
            Call GlobalMsg(GetPlayerName(index) & " needs an admin!", BrightGreen)
        Else
            Call PlayerMsg(index, "You are an admin!", BrightGreen)
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Call check commands packet ::
    ' ::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "checkcommands" Then
        s = Parse(1)
        If Scripting = 1 Then
            PutVar App.Path & "\Scripts\Command.ini", "TEMP", "Text" & index, Trim(s)
            MyScript.ExecuteStatement "Scripts\Main.txt", "Commands " & index
        Else
            Call PlayerMsg(index, "Thats not a valid command!", 12)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "prompt" Then
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerPrompt " & index & "," & Val(Parse(1)) & "," & Val(Parse(2))
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Request edit emoticon packet ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditemoticon" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "EMOTICONEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If

    If LCase(Parse(0)) = "editemoticon" Then
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse(1))
        
        If n < 0 Or n > MAX_EMOTICONS Then
            Call HackingAttempt(index, "Invalid Emoticon Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing emoticon #" & n & ".", ADMIN_LOG)
        Call SendEditEmoticonTo(index, n)
    End If

    If LCase(Parse(0)) = "saveemoticon" Then
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid Emoticon Index")
            Exit Sub
        End If

        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))

        Call SendUpdateEmoticonToAll(n)
        Call SaveEmoticon(n)
        Call AddLog(GetPlayerName(index) & " saved emoticon #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "checkemoticons" Then
        n = Emoticons(Val(Parse(1))).Pic
        
        Call SendDataToMap(GetPlayerMap(index), "checkemoticons" & SEP_CHAR & index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "mapreport") Then
        
        Packs = "mapreport" & SEP_CHAR
        For i = 1 To MAX_MAPS
            Packs = Packs & Map(i).Name & SEP_CHAR
        Next i
        Packs = Packs & END_CHAR
        
        Call SendDataTo(index, Packs)
        Exit Sub
    End If
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
        
    Call PlayerMsg(index, s, WhoColor)
End Sub

Sub SendOnlineList()
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
End Sub

Sub SendChars(ByVal index As Long)
Dim Packet As String
Dim i As Long
    
    Packet = "ALLCHARS" & SEP_CHAR
    For i = 1 To MAX_CHARS
        Packet = Packet & Trim(Player(index).Char(i).Name) & SEP_CHAR & Trim(Class(Player(index).Char(i).Class).Name) & SEP_CHAR & Player(index).Char(i).Level & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendJoinMap(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = ""
    
    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index And GetPlayerMap(i) = GetPlayerMap(index) Then
            Packet = Packet & "PLAYERDATA" & SEP_CHAR & i & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & GetPlayerSprite(i) & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & GetPlayerX(i) & SEP_CHAR & GetPlayerY(i) & SEP_CHAR & GetPlayerDir(i) & SEP_CHAR & GetPlayerAccess(i) & SEP_CHAR & GetPlayerPK(i) & SEP_CHAR & GetPlayerGuild(i) & SEP_CHAR & GetPlayerGuildAccess(i) & SEP_CHAR & END_CHAR
            Call SendDataTo(index, Packet)
        End If
    Next i
    
    ' Send index's player data to everyone on the map including himself
    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & GetPlayerGuild(index) & SEP_CHAR & GetPlayerGuildAccess(index) & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & GetPlayerGuild(index) & SEP_CHAR & GetPlayerGuildAccess(index) & SEP_CHAR & END_CHAR
    Call SendDataToMapBut(index, MapNum, Packet)
End Sub

Sub SendPlayerData(ByVal index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & SEP_CHAR & GetPlayerGuild(index) & SEP_CHAR & GetPlayerGuildAccess(index) & SEP_CHAR & GetPlayerClass(index) & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long
Dim i As Long

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR
    
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
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
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

Sub SendItems(ByVal index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Trim(Item(i).Name) <> "" Then
            Call SendUpdateItemTo(index, i)
        End If
    Next i
End Sub

Sub SendEmoticons(ByVal index As Long)
Dim Packet As String
Dim i As Long

    For i = 0 To MAX_EMOTICONS
        If Trim(Emoticons(i).Command) <> "" Then
            Call SendUpdateEmoticonTo(index, i)
        End If
    Next i
End Sub

Sub SendNpcs(ByVal index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_NPCS
        If Trim(Npc(i).Name) <> "" Then
            Call SendUpdateNpcTo(index, i)
        End If
    Next i
End Sub

Sub SendInventory(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendWornEquipment(ByVal index As Long)
Dim Packet As String
    
    Packet = "PLAYERWORNEQ" & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendHP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendMP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendSP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendStats(ByVal index As Long)
Dim Packet As String
    
    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & GetPlayerNextLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendClasses(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendNewCharClasses(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR & Class(i).MaleSprite & SEP_CHAR & Class(i).FemaleSprite & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendLeftGame(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & "" & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & END_CHAR
    Call SendDataToAllBut(index, Packet)
End Sub

Sub SendPlayerXY(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditItemTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateEmoticonTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditEmoticonTo(ByVal index As Long, ByVal EmoNum As Long)
Dim Packet As String

    Packet = "EDITEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Trim(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditNpcTo(ByVal index As Long, ByVal NpcNum As Long)
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
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditShopTo(ByVal index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendSpells(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim(Spell(i).Name) <> "" Then
            Call SendUpdateSpellTo(index, i)
        End If
    Next i
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditSpellTo(ByVal index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendTrade(ByVal index As Long, ByVal ShopNum As Long)
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
End Sub

Sub SendPlayerSpells(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "SPELLS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendWeatherTo(ByVal index As Long)
Dim Packet As String

    Packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
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
Dim Packet As String

    Packet = "TIME" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendTimeToAll()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendTimeTo(i)
        End If
    Next i
End Sub

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal index As Long)
Dim Packet As String

    Packet = "MAPMSG2" & SEP_CHAR & Msg & SEP_CHAR & index & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub
