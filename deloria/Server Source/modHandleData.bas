Attribute VB_Name = "modHandleData"
Sub HandleData(ByVal index As Long, ByVal Data As String)
On Error GoTo ErrorHandler
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
Dim Amount As Long
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
Dim Packet As String
Dim TempNum As Long, TempVal As Long
        
    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
                
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Requesting classes for making a character ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "getclasses" Then
        If Not IsPlaying(index) Then
            Call SendNewCharClasses(index)
        End If
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
        
        If FindPlayer(Parse(1)) = index Then
            Call PlayerMsg(index, "You cant change your guild access!", BrightRed)
            Exit Sub
        End If
        
        If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(index) Then
            Call PlayerMsg(index, "Player is not in your guild", BrightRed)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(index) < 3 Then
            Call PlayerMsg(index, "You need to be a higher access to do this!", BrightRed)
            Exit Sub
        End If
        
        'Set the player's new access level
        If Val(Parse(2)) < 0 Then Parse(2) = 0
        If Val(Parse(2)) > 4 Then Parse(2) = 4
        
        If Val(Parse(2)) > GetPlayerGuildAccess(index) Then
            Call PlayerMsg(index, "You cant set a someones guild access higher than your own!", BrightRed)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(index) <= GetPlayerGuildAccess(FindPlayer(Parse(1))) Then
            Call PlayerMsg(index, "You cant change " & GetPlayerName(FindPlayer(Parse(1))) & "'s guild access!", BrightRed)
            Exit Sub
        End If
        
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), Val(Parse(2)))
        Call SendPlayerData(FindPlayer(Parse(1)))
        Call PlayerMsg(FindPlayer(Parse(1)), "Your guild access has been changed to " & Val(Parse(2)) & "!", Yellow)
        Call PlayerMsg(index, "You changed " & GetPlayerName(FindPlayer(Parse(1))) & "'s guild access to " & Val(Parse(2)) & "!", Yellow)
        'Can have a message here if you'd like
        Exit Sub
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
        Exit Sub
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
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "buyguild" Then
        ' Check if they are alredy in a guild
        If GetPlayerGuild(index) <> "" Then
            Call PlayerMsg(index, "You are already in a guild!", BrightRed)
            Exit Sub
        End If
        
        For i = 1 To MAX_INV
            If Trim(LCase(Item(GetPlayerInvItemNum(index, i)).Name)) = "gold" Then
                If GetPlayerInvItemValue(index, i) >= 5000 Then
                    Call TakeItem(index, GetPlayerInvItemNum(index, i), 5000)
                    Call SetPlayerGuild(index, (Parse(1)))
                    Call SetPlayerGuildAccess(index, 4)
                    Call SendPlayerData(index)
                    Call PlayerMsg(index, "You have successfully created a guild!", BrightGreen)
                    Exit Sub
                Else
                    Call PlayerMsg(index, "You need " & 5000 - GetPlayerInvItemValue(index, i) & " more gold to buy a guild!", BrightRed)
                    Exit Sub
                End If
            End If
        Next i
        Call PlayerMsg(index, "You need 5000 Gold to buy a guild!", BrightRed)
        Exit Sub
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
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "kickfromguild" Then
        n = FindPlayer(Parse(1))
        
        If n = 0 Then
            Call PlayerMsg(index, "Player is offline.", White)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(index) <= 2 Then
            Call PlayerMsg(index, "You need be be a higher guild access to kick someone!", Red)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(n) >= GetPlayerGuildAccess(index) Then
            Call PlayerMsg(index, "You cant kick people with the same or higher guild access then you!", Red)
            Exit Sub
        End If
        
        If Trim(GetPlayerGuild(n)) <> Trim(GetPlayerGuild(index)) Then
            Call PlayerMsg(n, "The player needs to be in the same guild as you!", Red)
            Exit Sub
        End If
        
        Call PlayerMsg(n, "You have been kicked from the guild " & Trim(GetPlayerGuild(n)) & " !", Red)
        Call PlayerMsg(index, "You kicked " & Trim(GetPlayerName(n)) & " from the guild!", Red)
        Call SetPlayerGuild(n, "")
        Call SetPlayerGuildAccess(n, 0)
        Call SendPlayerData(n)
        Exit Sub
    End If
                
    If LCase(Parse(0)) = "guildinvite" Then
        If GetPlayerGuild(index) <> "" Then
            Call PlayerMsg(index, "You're already in a guild!", Red)
            Exit Sub
        End If
        If Trim(Player(index).GuildTemp = "") Then
            Call PlayerMsg(index, "No one invited you to a guild!", Red)
            Exit Sub
        End If
        If Val(Parse(1)) = 0 Then
            Call SetPlayerGuild(index, Player(index).GuildTemp)
            Call SetPlayerGuildAccess(index, 0)
            Call SendPlayerData(index)
            Call PlayerMsg(index, "You joined the guild " & Player(index).GuildTemp & "!", BrightGreen)
            Call PlayerMsg(Player(index).GuildInviter, GetPlayerName(index) & " joined your guild!", BrightGreen)
            Player(index).GuildInvitation = False
            Player(index).GuildTemp = ""
            Player(index).GuildInviter = 0
        Else
            Call PlayerMsg(index, "You you declined the invitation from the guild " & Player(index).GuildTemp & "!", BrightRed)
            Call PlayerMsg(Player(index).GuildInviter, GetPlayerName(index) & " declined your guild invitation.", BrightRed)
            Player(index).GuildInvitation = False
            Player(index).GuildTemp = ""
            Player(index).GuildInviter = 0
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "invitetoguild" Then
        i = FindPlayer(Parse(1))
        
        If i = 0 Then
            Call PlayerMsg(index, "Player is offline.", White)
            Exit Sub
        End If
        
        If Player(i).GuildInvitation = True Then
            Call PlayerMsg(index, "Player is already being invited to another guild.", White)
            Exit Sub
        End If
        
        If GetPlayerGuild(i) <> "" Then
            Call PlayerMsg(index, "Player is already in a guild.", White)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(index) <= 0 Then
            Call PlayerMsg(index, "You need to be a higher access to invite players into the guild!", Red)
            Exit Sub
        End If
        
        Player(i).GuildInvitation = True
        Player(i).GuildTemp = GetPlayerGuild(index)
        Player(i).GuildInviter = index
        
        Call PlayerMsg(i, "You have been invited to join the guild " & GetPlayerGuild(index) & ". Do you accept (/guildaccept) or decline (/guilddecline)?", BrightGreen)
        Call PlayerMsg(index, "You have invited " & Trim(GetPlayerName(i)) & " into your guild.", BrightGreen)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "guildchat" Then
        If Trim(GetPlayerGuild(index)) = "" Then
            Call PlayerMsg(index, "Your are not in a guild!", BrightRed)
            Exit Sub
        End If
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) = True Then
                If Trim(GetPlayerGuild(index)) = Trim(GetPlayerGuild(i)) Then
                    Call PlayerMsg(i, GetPlayerName(index) & " (Guild)> " & Parse(1), BrightGreen)
                End If
            End If
        Next i
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "guildwho" Then
        If Trim(GetPlayerGuild(index)) = "" Then
            Call PlayerMsg(index, "Your are not in a guild!", BrightRed)
            Exit Sub
        End If
        s = ""
        n = 0
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) = True Then
                If Trim(GetPlayerGuild(index)) = Trim(GetPlayerGuild(i)) Then
                    If s = "" Then
                        s = s & GetPlayerName(i)
                    Else
                        s = s & ", " & GetPlayerName(i)
                    End If
                    n = n + 1
                End If
            End If
        Next i
        
        Call PlayerMsg(index, "Guild Members Online (" & n & "): " & s, BrightGreen)
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
            
            If FindHDSerial(Player(index).HardDrive) Then
                Call AlertMsg(index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
            End If
            
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
                'Call AlertMsg(index, "Version outdated, please visit " & Trim(GetVar(App.Path & "\Data.ini", "CONFIG", "WebSite")))
                'Exit Sub
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
            If Trim(Parse(6)) = Trim(SEC_CODE1) And Trim(Parse(7)) = Trim(SEC_CODE2) Then
                'everything is ok!
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
            
            If FileExist("banks\" & Trim(Name) & ".ini") = False Then
                For i = 1 To MAX_CHARS
                    For n = 1 To MAX_BANK
                        Call PutVar(App.Path & "\banks\" & Trim(Name) & ".ini", "CHAR" & i, "BankItemNum" & n, STR(Player(index).Char(i).Bank(n).Num))
                        Call PutVar(App.Path & "\banks\" & Trim(Name) & ".ini", "CHAR" & i, "BankItemVal" & n, STR(Player(index).Char(i).Bank(n).Value))
                        Call PutVar(App.Path & "\banks\" & Trim(Name) & ".ini", "CHAR" & i, "BankItemDur" & n, STR(Player(index).Char(i).Bank(n).Dur))
                    Next n
                Next i
            End If
    
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
            If Class < 0 Or Class > 0 Then
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
            
            If Val(Parse(5)) <> 266 And Val(Parse(5)) <> 268 And Val(Parse(5)) <> 270 And Val(Parse(5)) <> 277 And Val(Parse(5)) <> 295 And Val(Parse(5)) <> 267 And Val(Parse(5)) <> 269 And Val(Parse(5)) <> 273 And Val(Parse(5)) <> 275 And Val(Parse(5)) <> 284 Then
                Call AlertMsg(index, "Please select one of the sprites on the list!")
                Exit Sub
            End If
                
            ' Everything went ok, add the character
            Call AddChar(index, Name, Sex, Class, CharNum)
            Player(index).Char(CharNum).Sprite = Val(Parse(5))
            Call SavePlayer(index, CharNum)
            Player(index).Char(CharNum).Sprite = Val(Parse(5))
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
                If frmServer.chkMod.Value = Checked Then
                    If GetPlayerAccess(index) <= 0 Then
                        Call AlertMsg(index, "The server is set to moderators only!")
                        Exit Sub
                    End If
                End If
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
    If LCase(Parse(0)) = "partychat" Then
        Msg = Parse(1)
                
        If Player(index).Party.InParty = YES Then
            If Player(index).Party.Started = YES Then
                n = index
            Else
                n = Player(index).Party.PlayerNums(1)
            End If
            Call PlayerMsg(n, GetPlayerName(index) & " (Party): " & Msg, BrightGreen)
            For i = 1 To MAX_PARTY_MEMS
                If Player(n).Party.PlayerNums(i) > 0 Then
                    Call PlayerMsg(Player(n).Party.PlayerNums(i), GetPlayerName(index) & " (Party): " & Msg, BrightGreen)
                End If
            Next i
        Else
            Call PlayerMsg(index, "You arent in a party!", BrightRed)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "saymsg" Then
        Msg = Parse(1)
                
        Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " : " & Msg & "", PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & ": " & Msg, RGB(191, 191, 191))
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "emotemsg" Then
        Msg = Parse(1)
                
        Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Msg, PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), "*" & GetPlayerName(index) & " " & Msg & "*", EmoteColor)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "broadcastmsg" Then
        Msg = Parse(1)
                
        If GetPlayerAccess(index) = 0 Then
            If MuteBroadcast = True Then
                Call PlayerMsg(index, "Broadcast chat is disabled!", BrightRed)
                Exit Sub
            End If
            If Player(index).Mute = True Then
                Call PlayerMsg(index, "Your broadcast chat is disabled!", BrightRed)
                Exit Sub
            End If
        End If
        
        s = GetPlayerName(index) & ": " & Msg
        Call AddLog(s, PLAYER_LOG)
        If GetPlayerAccess(index) > 0 Then
            Call BroadcastMsg(s, RGB(255, 155, 155))
        Else
            If GetPlayerLevel(index) >= 50 Then
                Call BroadcastMsg(s, RGB(255, 108, 108))
            Else
                Call BroadcastMsg(s, RGB(20, 175, 255))
            End If
        End If
        Call TextAdd(frmServer.txtText, s, True)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "globalmsg" Then
        Msg = Parse(1)
                
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
                
        If GetPlayerAccess(index) > 0 Then
            Call AddLog("(admin " & GetPlayerName(index) & ") " & Msg, ADMIN_LOG)
            Call AdminMsg(GetPlayerName(index) & " (Admin)> " & Msg, BrightGreen)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "playermsg" Then
        MsgTo = FindPlayer(Parse(1))
        Msg = Parse(2)
                
        ' Check if they are trying to talk to themselves
        If MsgTo <> index Then
            If MsgTo > 0 Then
                Call AddLog(GetPlayerName(index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
                Player(MsgTo).Reply = index
                Call PlayerMsg(MsgTo, "Message from " & GetPlayerName(index) & ": """ & Msg & """", RGB(0, 255, 255))
                Call PlayerMsg(index, "Message sent to " & GetPlayerName(MsgTo) & ": """ & Msg & """", RGB(0, 255, 255))
            Else
                Call PlayerMsg(index, "Player is not online.", White)
            End If
        End If
        
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "replymsg" Then
        MsgTo = Player(index).Reply
        Msg = Parse(1)
                
        ' Check if they are trying to talk to themselves
        If MsgTo <> index Then
            If MsgTo > 0 Then
                Call AddLog(GetPlayerName(index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
                Player(MsgTo).Reply = index
                Call PlayerMsg(MsgTo, "Message from " & GetPlayerName(index) & ": """ & Msg & """", RGB(0, 255, 255))
                Call PlayerMsg(index, "Message sent to " & GetPlayerName(MsgTo) & ": """ & Msg & """", RGB(0, 255, 255))
            Else
                Call PlayerMsg(index, "Player is not online.", White)
            End If
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
                    
                    Case ITEM_TYPE_GLOVES
                    If InvNum <> GetPlayerGlovesSlot(index) Then
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
                        Call SetPlayerGlovesSlot(index, InvNum)
                    Else
                        Call SetPlayerGlovesSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
                    
                    Case ITEM_TYPE_BOOTS
                    If InvNum <> GetPlayerBootsSlot(index) Then
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
                        Call SetPlayerBootsSlot(index, InvNum)
                    Else
                        Call SetPlayerBootsSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
                    
            Case ITEM_TYPE_RING
                    If InvNum <> GetPlayerRingSlot(index) Then
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
                        Call SetPlayerRingSlot(index, InvNum)
                    Else
                        Call SetPlayerRingSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
            
            Case ITEM_TYPE_AMULET
                    If InvNum <> GetPlayerAmuletSlot(index) Then
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
                        Call SetPlayerAmuletSlot(index, InvNum)
                    Else
                        Call SetPlayerAmuletSlot(index, 0)
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
                    
                Case ITEM_TYPE_SCROLL
                    Dim Maps As Long
                    Maps = Item(GetPlayerInvItemNum(index, InvNum)).Data1
                    x = Item(GetPlayerInvItemNum(index, InvNum)).Data2
                    y = Item(GetPlayerInvItemNum(index, InvNum)).Data3
                    
                    If Maps > 0 And Maps <= MAX_MAPS Then
                        Call PlayerWarp(index, Maps, x, y)
                        Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 1)
                    End If
                Case ITEM_TYPE_ORB
                    Maps = Item(GetPlayerInvItemNum(index, InvNum)).Data1
                    x = Item(GetPlayerInvItemNum(index, InvNum)).Data2
                    y = Item(GetPlayerInvItemNum(index, InvNum)).Data3
                    
                    If Maps > 0 And Maps <= MAX_MAPS Then
                        Call PlayerWarp(index, Maps, x, y)
                    End If
                Case ITEM_TYPE_BORB
                    If Map(Player(index).Char(Player(index).CharNum).Map).Indoors < 1 Then
                        Player(index).Char(Player(index).CharNum).Binding.Map = GetPlayerMap(index)
                        Player(index).Char(Player(index).CharNum).Binding.x = GetPlayerX(index)
                        Player(index).Char(Player(index).CharNum).Binding.y = GetPlayerY(index)
                        Call PlayerMsg(index, "You have been bound to this spot!", Yellow)
                        Call PlaySound(index, "magic24.wav")
                    Else
                        Call PlayerMsg(index, "You cannot bind indoors!", BrightRed)
                    End If
                Case ITEM_TYPE_GGORB
                    If Player(index).Party.InParty = YES Then
                        If Player(index).Party.Started = YES Then
                            For i = 1 To MAX_PARTY_MEMS
                                If Player(index).Party.PlayerNums(i) > 0 Then
                                    Call PlayerWarp(Player(index).Party.PlayerNums(i), GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                                    Call PlayerMsg(index, "You have been group gathered to " & GetPlayerName(index) & "!", Yellow)
                                End If
                            Next i
                        Else
                            Call PlayerWarp(Player(index).Party.PlayerNums(1), GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                            Call PlayerMsg(Player(index).Party.PlayerNums(1), "You have been group gathered to " & GetPlayerName(index) & "!", Yellow)
                            For i = 1 To MAX_PARTY_MEMS
                                If Player(index).Party.PlayerNums(i) > 0 And Player(index).Party.PlayerNums(i) <> index Then
                                    Call PlayerWarp(Player(Player(index).Party.PlayerNums(1)).Party.PlayerNums(i), GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                                    Call PlayerMsg(index, "You have been group gathered to " & GetPlayerName(index) & "!", Yellow)
                                End If
                            Next i
                            
                        End If
                        Call PlayerMsg(index, "Your group has been gathered to you!", BrightRed)
                    Else
                        Call PlayerMsg(index, "You are not in a party!", BrightRed)
                    End If
            End Select
            
            Call SendStats(index)
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
        End If
        Call SendInventory(index)
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
                    Damage = GetPlayerDamage(index) - GetPlayerProtection(i)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                    
                    If Damage > 0 Then
                        If GetPlayerSP(index) > 0 Then
                            Call AttackPlayer(index, i, Damage)
                            Call SetPlayerSP(index, GetPlayerSP(index) - 1)
                            Call SendSP(index)
                        End If
                    Else
                        Call BattleMsg(index, "You missed " & GetPlayerName(i) & "!", BrightRed, index)
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
                Damage = GetPlayerDamage(index) - Npc(MapNpc(GetPlayerMap(index), i).Num).DEF
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                
                If Damage > 0 Then
                    If GetPlayerSP(index) > 0 Then
                        Call AttackNpc(index, i, Damage)
                        Call SetPlayerSP(index, GetPlayerSP(index) - 1)
                        Call SendSP(index)
                    End If
                Else
                    Call BattleMsg(index, "You missed the " & Npc(MapNpc(GetPlayerMap(index), i).Num).Name & "!", BrightRed, index)
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
        If (PointType < 0) Or (PointType > 4) Then
            Call HackingAttempt(index, "Invalid Point Type")
            Exit Sub
        End If
                
        ' Make sure they have points
        If GetPlayerPOINTS(index) > 0 Then
            ' Take away a stat point
            Select Case PointType
                Case 0
                    If GetPlayerBaseSTR(index) >= 100 Then Exit Sub
                    Call SetPlayerSTR(index, GetPlayerBaseSTR(index) + 1)
                Case 1
                    If GetPlayerBaseDEF(index) >= 100 Then Exit Sub
                    Call SetPlayerDEF(index, GetPlayerBaseDEF(index) + 1)
                Case 2
                    If GetPlayerBaseMAGI(index) >= 100 Then Exit Sub
                    Call SetPlayerMAGI(index, GetPlayerBaseMAGI(index) + 1)
                    Call SendMP(index)
                Case 3
                    If GetPlayerBaseSPEED(index) >= 255 Then Exit Sub
                    Call SetPlayerSPEED(index, GetPlayerBaseSPEED(index) + 1)
                    Call SendSP(index)
                Case 4
                    If GetPlayerBaseVIT(index) >= 255 Then Exit Sub
                    Call SetPlayerVIT(index, GetPlayerBaseVIT(index) + 1)
                    Call SendHP(index)
            End Select
            
            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)
        End If
        
        ' Send the update
        Call PlayerPoints(index)
        Call SendStats(index)
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::::::::::::
    ' :: Player info request packet ::
    ' ::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinforequest" Then
        'Name = Parse(1)
        
        'i = FindPlayer(Name)
        'If i > 0 Then
            'Call PlayerMsg(index, "Account: " & Trim(Player(i).Login) & ", Name: " & GetPlayerName(i), BrightGreen)
            'If GetPlayerAccess(index) > ADMIN_MONITER Then
                'n = Int(GetPlayerSTR(i) / 2) + Int(GetPlayerLevel(i) / 2)
                'i = Int(GetPlayerDEF(i) / 2) + Int(GetPlayerLevel(i) / 2)
                'If n > 100 Then n = 100
                'If i > 100 Then i = 100
                'Call PlayerMsg(index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", BrightGreen)
            'End If
        'Else
            'Call PlayerMsg(index, "Player is not online.", White)
        'End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Warp me to packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "warpmeto" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < 1 Then
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
        If GetPlayerAccess(index) < 1 Then
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
        If GetPlayerAccess(index) < 1 Then
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
        'n = Int(GetPlayerSTR(index) / 2) + Int(GetPlayerLevel(index) / 2)
        'i = Int(GetPlayerDEF(index) / 2) + Int(GetPlayerLevel(index) / 2)
        'If n > 100 Then n = 100
        'If i > 100 Then i = 100
        'Call PlayerMsg(index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", White)
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
        Map(MapNum).Indoors = Val(Parse(n + 12))
        
        n = n + 13
        
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
        
        'Call SendDataTo(index, "reloadmaps" & SEP_CHAR & MapNum & SEP_CHAR & END_CHAR)
        
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
        Amount = Val(Parse(2))
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_INV Then
            Call HackingAttempt(index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
            If Amount <= 0 Then
                Call PlayerMsg(index, "You must drop more than 0!", BrightRed)
                Exit Sub
            End If
            
            If Amount > GetPlayerInvItemValue(index, InvNum) Then
                Call PlayerMsg(index, "You dont have that much to drop!", BrightRed)
                Exit Sub
            End If
        End If
        
        ' Prevent hacking
        If Item(GetPlayerInvItemNum(index, InvNum)).Type <> ITEM_TYPE_CURRENCY Then
            If Amount > GetPlayerInvItemValue(index, InvNum) Then
                Call HackingAttempt(index, "Item amount modification")
                Exit Sub
            End If
        End If
        
        If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_ORB Or Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_BORB Or Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_GGORB Then
            Call PlayerMsg(index, "You cant drop orbs!", BrightRed)
            Exit Sub
        End If
            
        Call PlayerMapDropItem(index, InvNum, Amount)
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
        
    Dim Name2 As String
    ' :::::::::::::::::::::
    ' :: Ban list packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "banlist" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MONITER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = 1
        f = FreeFile
        Open App.Path & "\banlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            Input #f, Name
            Input #f, Name2
            
            Call PlayerMsg(index, n & "> " & Name2 & "(" & s & "): banned by " & Name, White)
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
        Call Kill(App.Path & "\banHDlist.txt")
        Call PlayerMsg(index, "Ban list destroyed.", White)
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::
    ' :: Ban player packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "banplayer" Then
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MONITER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player index
        n = FindPlayer(Parse(1))
        
        If n <> index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(index) Then
                    Call BanIndex(n, index)
                    Call BanHDIndex(n, Player(n).HardDrive, index)
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
        Exit Sub
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
        Exit Sub
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
        Exit Sub
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
        Exit Sub
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
        
        If i <= 0 Then Exit Sub
        
        ' Check if inv full
        x = FindOpenInvSlot(index, Shop(i).TradeItem(n).GetItem)
        If x = 0 Then
            Call BattleMsg(index, "Trade unsuccessful, inventory full.", Yellow, index)
            Exit Sub
        End If
        
        ' Check if they have the item
        If HasItem(index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
            Call TakeItem(index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
            Call GiveItem(index, Shop(i).TradeItem(n).GetItem, Shop(i).TradeItem(n).GetValue)
            Call BattleMsg(index, "The trade was successful!", Yellow, index)
        Else
            Call BattleMsg(index, "Trade unsuccessful.", Yellow, index)
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
        If Item(GetPlayerInvItemNum(index, n)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(index, n)).Type > ITEM_TYPE_SHIELD Or Item(GetPlayerInvItemNum(index, n)).Type > ITEM_TYPE_BOOTS Or Item(GetPlayerInvItemNum(index, n)).Type > ITEM_TYPE_GLOVES Or Item(GetPlayerInvItemNum(index, n)).Type > ITEM_TYPE_RING Or Item(GetPlayerInvItemNum(index, n)).Type > ITEM_TYPE_AMULET Then
            Call PlayerMsg(index, "You can only fix weapons, armors, helmets, shields, boots, gloves, rings and amulets.", BrightRed)
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
                            
                ' Change target
                Player(index).Target = i
                Player(index).TargetType = TARGET_TYPE_PLAYER
                Call BattleMsg(index, "You see " & GetPlayerName(i) & " - " & "Level: " & GetPlayerLevel(i) & " - {" & GetPlayerGuild(i) & "}", Yellow, index)
                Exit Sub
            End If
        Next i
        
        ' Check for an item
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(GetPlayerMap(index), i).Num > 0 Then
                If MapItem(GetPlayerMap(index), i).x = x And MapItem(GetPlayerMap(index), i).y = y Then
                    Call BattleMsg(index, "You see a " & Trim(Item(MapItem(GetPlayerMap(index), i).Num).Name) & ".", Yellow, index)
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
                    Call BattleMsg(index, "Your target is now a " & Trim(Npc(MapNpc(GetPlayerMap(index), i).Num).Name) & ".", Yellow, index)
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
                        
        Call SendDataTo(index, "PPCHATTING" & SEP_CHAR & n & SEP_CHAR & GetPlayerName(n) & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "PPCHATTING" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & END_CHAR)
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

Dim Something As Boolean

Something = False

If Something = True Then
' :::::::::::::::::::
' :: Trade packets ::
' :::::::::::::::::::
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
        Dim CanTrade As Boolean
        CanTrade = False
        
        If GetPlayerX(index) = GetPlayerX(n) And GetPlayerY(index) + 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) = GetPlayerX(n) And GetPlayerY(index) - 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) + 1 = GetPlayerX(n) And GetPlayerY(index) = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) - 1 = GetPlayerX(n) And GetPlayerY(index) = GetPlayerY(n) Then CanTrade = True
            
        If CanTrade = True Then
            ' Check to see if player is already in a trade
            If Player(n).InTrade = 1 Then
                Call PlayerMsg(index, "Player is already in a trade!", Pink)
                Exit Sub
            End If
            
            Call PlayerMsg(index, "Trade request has been sent to " & GetPlayerName(n) & ".", Pink)
            Call PlayerMsg(n, GetPlayerName(index) & " wants you to trade with them.  Type /accept to accept, or /decline to decline.", Pink)
        
            Player(n).TradePlayer = index
            Player(index).TradePlayer = n
        Else
            Call PlayerMsg(index, "You need to be beside the player to trade!", Pink)
            Call PlayerMsg(n, "The player needs to be beside you to trade!", Pink)
        End If
        Exit Sub
    End If
    
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
        CanTrade = False
        
        If GetPlayerX(index) = GetPlayerX(n) And GetPlayerY(index) + 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) = GetPlayerX(n) And GetPlayerY(index) - 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) + 1 = GetPlayerX(n) And GetPlayerY(index) = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) - 1 = GetPlayerX(n) And GetPlayerY(index) = GetPlayerY(n) Then CanTrade = True
            
        If CanTrade = True Then
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
        Else
            Call PlayerMsg(index, "The player needs to be beside you to trade!", Pink)
            Call PlayerMsg(n, "You need to be beside the player to trade!", Pink)
        End If
        Exit Sub
    End If

    If LCase(Parse(0)) = "qtrade" Then
        n = Player(index).TradePlayer
        
        ' Check if anyone trade with player
        If n < 1 Then
            Call PlayerMsg(index, "No one requested a trade with you.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Stopped trading.", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " stopped trading with you!", Pink)

        Player(index).TradeOk = 0
        Player(n).TradeOk = 0
        Player(index).TradePlayer = 0
        Player(index).InTrade = 0
        Player(n).TradePlayer = 0
        Player(n).InTrade = 0
        Call SendDataTo(index, "qtrade" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "qtrade" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If

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
                For i = 1 To MAX_PLAYER_TRADES
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(n, x) < 1 Then
                            If Player(index).Trading(i).InvNum > 0 Then
                                Call GiveItem(n, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), 1)
                                Call TakeItem(index, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), 1)
                                Exit For
                            End If
                        End If
                    Next x
                Next i

                For i = 1 To MAX_PLAYER_TRADES
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(index, x) < 1 Then
                            If Player(n).Trading(i).InvNum > 0 Then
                                Call GiveItem(index, GetPlayerInvItemNum(n, Player(n).Trading(i).InvNum), 1)
                                Call TakeItem(n, GetPlayerInvItemNum(n, Player(n).Trading(i).InvNum), 1)
                                Exit For
                            End If
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
            Player(index).TradeOk = 0
            Player(n).TradePlayer = 0
            Player(n).InTrade = 0
            Player(n).TradeOk = 0
            Call SendDataTo(index, "qtrade" & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, "qtrade" & SEP_CHAR & END_CHAR)
        End If
        Exit Sub
    End If
End If
' :::::::::::::::::::::::
' :: End Trade packets ::
' :::::::::::::::::::::::

    ' ::::::::::::::::::
    ' :: Party packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "party" Then
        n = FindPlayer(Parse(1))
        
        ' Prevent partying with self
        If n = index Then
            Call PlayerMsg(index, "You cannot party with yourself!", Pink)
            Exit Sub
        End If
                
        ' Check for a previous party and if so drop it
        For i = 1 To MAX_PARTY_MEMS
            If Player(index).Party.PlayerNums(i) > 0 Then
                Call PlayerMsg(index, "You are already in a party!", Pink)
                Exit Sub
            End If
            If Player(n).Party.PlayerNums(i) > 0 Then
                Call PlayerMsg(index, "The person is already in a party!", Pink)
                Exit Sub
            End If
        Next i
        
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
            If Player(n).Party.InParty = NO Then
                Call PlayerMsg(index, "Party request has been sent to " & GetPlayerName(n) & ".", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " wants you to join their party.  Type /join to join, or /leave to decline.", Pink)
            
                Player(index).Party.Started = YES
                Player(index).Party.PlayerNums(1) = n
                Player(n).Party.PlayerNums(1) = index
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
        n = Player(index).Party.PlayerNums(1)
        
        If n > 0 Then
            ' Check to make sure they aren't the starter
            If Player(index).Party.Started = NO Then
                ' Check to make sure that each of their party players match
                For i = 1 To MAX_PARTY_MEMS
                    If Player(n).Party.PlayerNums(i) = index Then
                        Call PlayerMsg(index, "You have joined " & GetPlayerName(n) & "'s party!", Pink)
                        Call PlayerMsg(n, GetPlayerName(index) & " has joined your party!", Pink)
                        
                        Player(index).Party.InParty = YES
                        Player(n).Party.InParty = YES
                    ElseIf Player(n).Party.PlayerNums(i) > 0 Then
                        Call PlayerMsg(Player(n).Party.PlayerNums(i), GetPlayerName(index) & " has joined your party!", Pink)
                    End If
                Next i
                If Player(index).Party.InParty = NO Then
                    Call PlayerMsg(index, "Party failed.", Pink)
                End If
                Call SendParty(index)
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
        If Player(index).Party.Started = YES Then
            n = index
        Else
            n = Player(index).Party.PlayerNums(1)
        End If
            
        If n > 0 Then
            If Player(index).Party.InParty = YES Then
                If Player(index).Party.Started = NO Then
                    Call PlayerMsg(index, "You have left the party.", Pink)
                    Call PlayerMsg(n, GetPlayerName(index) & " has left the party.", Pink)
                    For i = 1 To MAX_PARTY_MEMS
                        If Player(n).Party.PlayerNums(i) = index Then
                            Player(n).Party.PlayerNums(i) = 0
                        ElseIf Player(n).Party.PlayerNums(i) > 0 Then
                            Call PlayerMsg(Player(n).Party.PlayerNums(i), GetPlayerName(index) & " has left the party.", Pink)
                        End If
                    Next i
                    Call SendParty(Player(index).Party.PlayerNums(1))
                    Player(index).Party.PlayerNums(1) = 0
                    Player(index).Party.Started = NO
                    Player(index).Party.InParty = NO
                    Call SendParty(index)
                    For i = 1 To MAX_PARTY_MEMS
                        If Player(n).Party.PlayerNums(i) > 0 Then
                            Exit Sub
                        End If
                    Next i
                    Player(n).Party.InParty = NO
                    Player(n).Party.Started = NO
                    Call SendParty(n)
                    Call PlayerMsg(n, "The party has now been disbanded.", Pink)
                Else
                    Call PlayerMsg(index, "You have left the party.", Pink)
                    For i = 1 To MAX_PARTY_MEMS
                        If Player(index).Party.PlayerNums(i) > 0 Then
                            Call PlayerMsg(Player(index).Party.PlayerNums(i), GetPlayerName(index) & " has left the party. It will now be disbanded.", Pink)
                            Player(Player(index).Party.PlayerNums(i)).Party.PlayerNums(1) = 0
                            Player(Player(index).Party.PlayerNums(i)).Party.Started = NO
                            Player(Player(index).Party.PlayerNums(i)).Party.InParty = NO
                            Call SendParty(Player(index).Party.PlayerNums(i))
                        End If
                        Player(index).Party.PlayerNums(i) = 0
                    Next i
                    Player(index).Party.Started = NO
                    Player(index).Party.InParty = NO
                    Call SendParty(index)
                End If
            Else
                If Player(index).Party.Started = NO Then
                    Call PlayerMsg(index, "Declined party request.", Pink)
                    For i = 1 To MAX_PARTY_MEMS
                        If Player(n).Party.PlayerNums(i) = index Then
                            Player(n).Party.PlayerNums(i) = 0
                        End If
                    Next i
                    Player(index).Party.PlayerNums(1) = 0
                    Player(index).Party.Started = NO
                    Player(index).Party.InParty = NO
                    Call SendParty(index)
                    For i = 1 To MAX_PARTY_MEMS
                        If Player(n).Party.PlayerNums(i) > 0 Then
                            Exit Sub
                        End If
                    Next i
                    Call PlayerMsg(n, GetPlayerName(index) & " has left or refused the party.", Pink)
                    Player(n).Party.InParty = NO
                    Player(n).Party.Started = NO
                    For i = 1 To MAX_PARTY_MEMS
                        Player(n).Party.PlayerNums(i) = 0
                    Next i
                    Call SendParty(n)
                Else
                    For i = 1 To MAX_PARTY_MEMS
                        If Player(n).Party.PlayerNums(i) > 0 Then
                            Call PlayerMsg(Player(index).Party.PlayerNums(i), GetPlayerName(index) & " has disbanded the party.", Pink)
                            Player(index).Party.PlayerNums(i) = 0
                            Player(Player(index).Party.PlayerNums(i)).Party.InParty = NO
                            Player(Player(index).Party.PlayerNums(i)).Party.InParty = NO
                            Call SendParty(Player(index).Party.PlayerNums(i))
                        End If
                    Next i
                    Call PlayerMsg(index, "Disbanded party.", Pink)
                End If
            End If
        Else
            Call PlayerMsg(index, "You are not in a party!", Pink)
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Invite to party packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "invite" Then
        n = FindPlayer(Parse(1))
        Dim FreeSpace As Boolean
        FreeSpace = False
        
        ' Prevent partying with self
        If n = index Then
            Exit Sub
        End If
        
        ' Check if user is in a party
        If Player(index).Party.InParty = NO Then
            Call PlayerMsg(index, "You are not in a party!", Pink)
            Exit Sub
        End If
                
        ' Check if they are the starter
        If Player(index).Party.Started = NO Then
            Call PlayerMsg(index, "You are not the starter of this party!", Pink)
            Exit Sub
        End If
        
        ' Check if the party is full
        For i = 1 To MAX_PARTY_MEMS
            If Player(index).Party.PlayerNums(i) = 0 Then
                FreeSpace = True
                Exit For
            End If
        Next i
        If FreeSpace = False Then
            Call PlayerMsg(index, "Your party is full!", Pink)
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
            If Player(n).Party.InParty = NO Then
                Call PlayerMsg(index, "Party request has been sent to " & GetPlayerName(n) & ".", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " wants you to join their party.  Type /join to join, or /leave to decline.", Pink)
                For i = 1 To MAX_PARTY_MEMS
                    If Player(index).Party.PlayerNums(i) = 0 Then
                        Player(index).Party.PlayerNums(i) = n
                        GoTo NoFor
                    End If
                Next i
NoFor:
                Player(n).Party.PlayerNums(1) = index
            Else
                Call PlayerMsg(index, "Player is already in a g!", Pink)
            End If
        Else
            Call PlayerMsg(index, "Player is not online.", White)
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
        Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
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
                    If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerBootsSlot(index) <> i And GetPlayerGlovesSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerAmuletSlot(index) <> i Then
                        Call SetPlayerInvItemNum(index, i, 0)
                        Call PlayerMsg(index, "You have bought a new sprite!", BrightGreen)
                        Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                        Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(index)
                    End If
                End If
                If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerBootsSlot(index) <> i And GetPlayerGlovesSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerAmuletSlot(index) <> i Then
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
        Exit Sub
    End If

    If LCase(Parse(0)) = "saveemoticon" Then
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        If n < 0 Or n > MAX_EMOTICONS Then
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
    
    If LCase(Parse(0)) = "traininghouse" Then
        If Map(GetPlayerMap(index)).Moral = MAP_MORAL_TRAINING Then
            If GetPlayerPOINTS(index) > 0 Then
                Call SendDataTo(index, "traininghouse" & SEP_CHAR & END_CHAR)
                Call SendStats(index)
                Call PlayerPoints(index)
            Else
                Call PlayerMsg(index, "You need more train skill points!", BrightRed)
            End If
        Else
            Call PlayerMsg(index, "You need to be in one of the training houses!", BrightRed)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "masswarp" Then
        If GetPlayerAccess(index) >= ADMIN_CREATOR Then
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) = True Then
                    If GetPlayerAccess(i) <= ADMIN_MONITER Then
                        Call PlayerWarp(i, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                    End If
                End If
            Next i
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "mutebroadcast" Then
        If GetPlayerAccess(index) > 0 Then
            If MuteBroadcast = True Then
                Call GlobalMsg("Broadcast chat enabled!", BrightGreen)
                MuteBroadcast = False
            Else
                Call GlobalMsg("Broadcast chat disabled!", BrightGreen)
                MuteBroadcast = True
            End If
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "muteplayer" Then
        n = FindPlayer(Parse(1))
                
        If n = index Then
            Call PlayerMsg(index, "You cant mute yourself.", White)
            Exit Sub
        End If
        
        If n > 0 Then
            If GetPlayerAccess(index) <= 0 Then Exit Sub
            If GetPlayerAccess(n) <= 0 Then
                If Player(n).Mute = True Then
                    Call PlayerMsg(index, "You have unmuted " & GetPlayerName(n) & "!", White)
                    Call PlayerMsg(n, GetPlayerName(index) & " has unmuted you!", White)
                    Player(n).Mute = False
                Else
                    Call PlayerMsg(index, "You have muted " & GetPlayerName(n) & "!", BrightRed)
                    Call PlayerMsg(n, GetPlayerName(index) & " has muted you!", BrightRed)
                    Player(n).Mute = True
                End If
            Else
                Call PlayerMsg(index, "You can only mute players.", White)
            End If
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "unjailplayer" Then
        n = FindPlayer(Parse(1))
        
        If n = 0 Then
            Call PlayerMsg(index, "Player is offline.", White)
            Exit Sub
        End If
        
        If GetPlayerAccess(index) <= 0 Then
            Call PlayerMsg(index, "You need to be a higher access to release someone from jail!", BrightRed)
            Exit Sub
        End If
            
        i = Rand(3, 1)
        Call GlobalMsg(GetPlayerName(n) & " has been released from jail!", Yellow)
        Call PlayerWarp(n, 19, 15, 27)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "jailplayer" Then
        n = FindPlayer(Parse(1))
        
        If n = 0 Then
            Call PlayerMsg(index, "Player is offline.", White)
            Exit Sub
        End If
        
        If GetPlayerAccess(index) <= 0 Then
            Call PlayerMsg(index, "You need to be a higher access to jail someone!", BrightRed)
            Exit Sub
        End If
            
        i = Rand(3, 1)
        Call GlobalMsg(GetPlayerName(n) & " has been jailed!", Yellow)
        
        If i = 1 Then
            Call PlayerWarp(n, 19, 10, 22)
        ElseIf i = 2 Then
            Call PlayerWarp(n, 19, 15, 22)
        Else
            Call PlayerWarp(n, 19, 20, 22)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "gmtime" Then
        GameTime = Val(Parse(1))
        Call SendTimeToAll
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "weather" Then
        GameWeather = Val(Parse(1))
        Call SendWeatherToAll
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "intensity" Then
        RainIntensity = Val(Parse(1))
        frmServer.lblRainIntensity.Caption = "Intensity: " & RainIntensity
        frmServer.scrlRainIntensity.Value = RainIntensity
        Call SendWeatherToAll
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "changepass" Then
        Dim OldPass As String
        Msg = Trim(Parse(1))
        OldPass = Trim(GetVar(App.Path & "\accounts\" & GetPlayerLogin(index) & ".ini", "GENERAL", "Password"))
        If LCase(Trim(OldPass)) = LCase(Trim(Msg)) Then
            Call PlayerMsg(index, "Your current password is already " & Trim(Msg) & "!", 12)
        ElseIf LCase(OldPass) <> LCase(Msg) Then
            Call PlayerMsg(index, "Your new password is now " & Trim(Msg) & "!", 10)
            Player(index).Password = Trim(Msg)
            Call PutVar(App.Path & "\accounts\" & GetPlayerLogin(index) & ".ini", "GENERAL", "Password", LCase(Trim(TextSay)))
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "changename" Then
        If GetPlayerAccess(index) < ADMIN_CREATOR Then
            Call PlayerMsg(index, "You need to be an admin to do this!", BrightRed)
            Exit Sub
        End If
        
        i = FindPlayer(Trim(Parse(1)))
        
        If i > 0 Then
            Call PlayerMsg(index, "Successfully changed " & GetPlayerName(i) & " to " & Trim(Parse(2)) & "!", Pink)
            Call SetPlayerName(i, Trim(Parse(2)))
            Call PlayerMsg(i, "Your name has been changed to " & Trim(Parse(2)) & "!", Pink)
            Call SendPlayerData(i)
        Else
            Call PlayerMsg(index, "Player not online.", White)
            Exit Sub
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "configuremaps" Then
        Player(index).HardDrive = Val(Parse(1))
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "bankdeposit" Then
        x = GetPlayerInvItemNum(index, Val(Parse(1)))
        i = FindOpenBankSlot(index, x)
        If i = 0 Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Bank full!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        If Val(Parse(2)) > GetPlayerInvItemValue(index, Val(Parse(1))) Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You cant deposit more than you have!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        If GetPlayerWeaponSlot(index) = Val(Parse(1)) Or GetPlayerArmorSlot(index) = Val(Parse(1)) Or GetPlayerShieldSlot(index) = Val(Parse(1)) Or GetPlayerHelmetSlot(index) = Val(Parse(1)) Or GetPlayerBootsSlot(index) = Val(Parse(1)) Or GetPlayerGlovesSlot(index) = Val(Parse(1)) Or GetPlayerRingSlot(index) = Val(Parse(1)) Or GetPlayerAmuletSlot(index) = Val(Parse(1)) Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You cant deposit worn equipment!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        If Item(x).Type = ITEM_TYPE_CURRENCY Then
            If Val(Parse(2)) <= 0 Then
                Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You must deposit more than 0!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
        
        Call TakeItem(index, x, Val(Parse(2)))
        Call GiveBankItem(index, x, Val(Parse(2)), i)
        
        Call SendBank(index)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "bankwithdraw" Then
        i = GetPlayerBankItemNum(index, Val(Parse(1)))
        TempVal = Val(Parse(2))
        x = FindOpenInvSlot(index, i)
        If x = 0 Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Inventory full!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        If Val(Parse(2)) > GetPlayerBankItemValue(index, Val(Parse(1))) Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You cant withdraw more than you have!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
                
        If Item(i).Type = ITEM_TYPE_CURRENCY Then
            If Val(Parse(2)) <= 0 Then
                Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You must withdraw more than 0!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If

            If Trim(LCase(Item(GetPlayerInvItemNum(index, x)).Name)) <> "gold" Then
                If GetPlayerInvItemValue(index, x) + Val(Parse(2)) > 100 Then
                    TempVal = 100 - GetPlayerInvItemValue(index, x)
                End If
            End If
        End If
                
        Call GiveItem(index, i, TempVal)
        Call TakeBankItem(index, i, TempVal)
        
        Call SendBank(index)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "playermove" And Player(index).GettingMap = YES Then Exit Sub
    If LCase(Parse(0)) = "playerdir" And Player(index).GettingMap = YES Then Exit Sub
    
Call AdminMsg(GetPlayerName(index) & " is trying to hack with parse '" & Parse(0) & "'", BrightGreen)
Call HackingAttempt(index, "Hacking Attempt!")
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modHamdleData.bas", "HandleData", Err.Number, Err.Description
  Call TextAdd(frmServer.txtErrorLog, ":: Report This To GSD ::", True)
  Call TextAdd(frmServer.txtErrorLog, Parse(0) & " - " & " has an error!", True)
  Call TextAdd(frmServer.txtErrorLog, ":: End Report ::", True)
End Sub
