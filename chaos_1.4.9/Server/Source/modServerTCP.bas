Attribute VB_Name = "modServerTCP"

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit

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

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim I As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For I = 1 To MAX_PLAYERS

        If IsPlaying(I) And GetPlayerAccess(I) > 0 Then
            Call SendDataTo(I, Packet)
        End If
    Next
End Sub

Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
    Call CloseSocket(Index)
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

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If
        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & " for (" & Reason & ").")
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
Dim MsgTo As Long
Dim Dir As Long
Dim InvNum As Long
Dim Amount As Long
Dim Damage As Long
Dim PointType As Long
Dim Movement As Long
Dim I As Long, N As Long, x As Long, y As Long, f As Long
Dim MapNum As Long
Dim s As String
Dim ShopNum As Long, ItemNum As Long
Dim DurNeeded As Long, GoldNeeded As Long
Dim z As Long
Dim Packet As String
Dim o As Long
Dim TempNum As Long, TempVal As Long
On Error GoTo hell

    Parse = Split(Data, SEP_CHAR)

    ' Parse's Without Being Online
    If Not IsPlaying(Index) Then

        Select Case LCase$(Parse(0))

            Case "getinfo"
                Call SendInfo(Index)
                Call SendNewsTo(Index)
                Exit Sub

            Case "gatglasses"
                Call SendNewCharClasses(Index)
                Exit Sub

            Case "newfaccountied"
            Dim Email As String
            Dim Vault As String
                If Not IsLoggedIn(Index) Then
                    Name = Parse(1)
                    Password = Parse(2)
                    Email = Parse(3)
                    Vault = Parse(4)
                    
                    For I = 1 To Len(Name)
                        N = Asc(Mid$(Name, I, 1))

                        If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                        Else
                            Call PlainMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.", 1)
                            Exit Sub
                        End If
                    Next

                    If Not AccountExist(Name) Then
                        Call AddAccount(Index, Name, Password, Email, Vault)
                        Call TextAdd(frmServer.txtText(0), "Account " & Name & " has been created.", True)
                        Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                        Call PlainMsg(Index, "Your account has been created!", 1)
                    Else
                        Call PlainMsg(Index, "Sorry, that account name is already taken!", 1)
                    End If
                End If
                Exit Sub

            Case "logination"

                If Not IsLoggedIn(Index) Then
                    Name = Parse(1)
                    Password = Parse(2)
                    For I = 1 To Len(Name)
                        N = Asc(Mid$(Name, I, 1))

                        If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                        Else
                            Call PlainMsg(Index, "Account duping is not allowed!", 3)
                            Exit Sub
                        End If
                    Next
                    Dim Encryptor As clsCRijndael

                    ' I like being creative with variable names
                    Dim temp1() As Byte
                    Dim temp2() As Byte
                    Dim temp3() As Byte
                    Dim temp4 As String
                    Dim temp5 As String

                    Set Encryptor = New clsCRijndael
                    temp1 = Parse(6)
                    temp2 = Parse(3) & "." & Parse(4) & "." & Parse(5)
                    temp3 = Encryptor.EncryptData(temp1, temp2)
                    temp4 = ""
                    For I = 0 To UBound(temp3)
                        temp4 = temp4 & Right$("0" & Hex$(temp3(I)), 2)
                    Next
                    temp1 = SEC_CODE
                    temp2 = CLIENT_MAJOR & "." & CLIENT_MINOR & "." & CLIENT_REVISION
                    temp3 = Encryptor.EncryptData(temp1, temp2)
                    temp5 = ""
                    For I = 0 To UBound(temp3)
                        temp5 = temp5 & Right$("0" & Hex$(temp3(I)), 2)
                    Next

                    If temp4 <> temp5 Then
                        Call SendDataTo(Index, "sound" & SEP_CHAR & "ANewVersionHasBeenReleased" & SEP_CHAR & END_CHAR)
                        Call PlainMsg(Index, "Version outdated, please visit " & Trim$(GetVar(App.Path & "\Data.ini", "CONFIG", "WebSite")), 3)
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
                    Packs = Packs & MAX_SPEECH & SEP_CHAR
                    Packs = Packs & MAX_ELEMENTS & SEP_CHAR
                    Packs = Packs & PAPERDOLL & SEP_CHAR
                    Packs = Packs & SIZE_X & SEP_CHAR
                    Packs = Packs & SIZE_Y & SEP_CHAR
                    Packs = Packs & END_CHAR
                    Call SendDataTo(Index, Packs)
                    Call LoadPlayer(Index, Name)
                    Call SendChars(Index)
                    Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
                    Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", True)
                End If
                Exit Sub

            Case "addachara"
            Dim RacePath As Long
                Name = Parse(1)
                Sex = Val(Parse(2))
                Class = Val(Parse(3))
                CharNum = Val(Parse(4))
                RacePath = Val(Parse(5))
                
                For I = 1 To Len(Name)
                    N = Asc(Mid$(Name, I, 1))

                    If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                    Else
                        Call PlainMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.", 4)
                        Exit Sub
                    End If
                Next

                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(Index, "Invalid CharNum")
                    Exit Sub
                End If

                If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
                    Call HackingAttempt(Index, "Invalid Sex")
                    Exit Sub
                End If

                If Class < 1 Or Class > Max_Classes Then
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
                Call AddChar(Index, Name, Sex, Class, CharNum, RacePath)
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
                        Open App.Path & "\main\accounts\charlist.txt" For Append As #f
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

    Select Case LCase$(Parse(0))

            ' :::::::::::::::::::
            ' :: Guilds Packet ::
            ' :::::::::::::::::::
            ' Access
        Case "guildchangeaccess"

            ' Check the requirements.
            If FindPlayer(Parse(1)) = 0 Then
                Call PlayerMsg(Index, "Player is offline", White)
                Exit Sub
            End If

            If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(Index) Then
                Call PlayerMsg(Index, "Player is not in your guild", Red)
                Exit Sub
            End If

            'Set the player's new access level
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
            Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 5)
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
            
            ' :::::::::::::::::::::::
            ' :: Bug Report packet ::
            ' :::::::::::::::::::::::
     Case "bugreport"
    Dim BugReport As String
    Dim Message As String
    If GetPlayerLevel(Index) >= 1 Then
    Message = Trim(Parse(1))
    'BugReport = Time & " " & GetPlayerName(Index) & ": " & Message
            Call AddLog(GetPlayerName(Index) & ": " & Message & "", BUG_LOG)
            Call PlayerMsg(Index, "Thank you for reporting this bug, " & GetPlayerName(Index), White)
    Exit Sub
    End If
    
             ' :::::::::::::::::::::::
            ' :: Suggestion Report packet ::
            ' :::::::::::::::::::::::
     Case "suggestionreport"
    'Dim Message As String
    If GetPlayerLevel(Index) >= 1 Then
    Message = Trim(Parse(1))
            Call AddLog(GetPlayerName(Index) & ": " & Message & "", SUGGESTION_LOG)
            Call PlayerMsg(Index, "Thank you for reporting your Suggestions, " & GetPlayerName(Index), White)
    Exit Sub
    End If

            ' ::::::::::::::::::::
            ' :: Social packets ::
            ' ::::::::::::::::::::
        Case "saymsg"
            Msg = Parse(1)

            ' Prevent hacking
            For I = 1 To Len(Msg)

                If Asc(Mid$(Msg, I, 1)) < 32 Or Asc(Mid$(Msg, I, 1)) > 126 Then
                    Call HackingAttempt(Index, "Say Text Modification")
                    Exit Sub
                End If
            Next

            If frmServer.chkM.Value = Unchecked Then
                If GetPlayerAccess(Index) <= 0 Then
                    Call PlayerMsg(Index, "Map messages have been disabled by the server!", BrightRed)
                    Exit Sub
                End If
            End If
           
            If LANGUAGEFILTER = 1 Then
            'Check for swearing
            If SwearCheck(Msg) = True Then
            Call PlayerMsg(Index, "Please use appropriate language.", Red)
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

                If Asc(Mid$(Msg, I, 1)) < 32 Or Asc(Mid$(Msg, I, 1)) > 126 Then
                    Call HackingAttempt(Index, "Emote Text Modification")
                    Exit Sub
                End If
            Next

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

                If Asc(Mid$(Msg, I, 1)) < 32 Or Asc(Mid$(Msg, I, 1)) > 126 Then
                    Call HackingAttempt(Index, "Broadcast Text Modification")
                    Exit Sub
                End If
            Next

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

                If Asc(Mid$(Msg, I, 1)) < 32 Or Asc(Mid$(Msg, I, 1)) > 126 Then
                    Call HackingAttempt(Index, "Global Text Modification")
                    Exit Sub
                End If
            Next

            If frmServer.chkG.Value = Unchecked Then
                If GetPlayerAccess(Index) <= 0 Then
                    Call PlayerMsg(Index, "Global messages have been disabled by the server!", BrightRed)
                    Exit Sub
                End If
            End If

            If Player(Index).Mute = True Then Exit Sub
            If GetPlayerAccess(Index) > 0 Then
                s = "(global) " & GetPlayerName(Index) & ": " & Msg
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

                If Asc(Mid$(Msg, I, 1)) < 32 Or Asc(Mid$(Msg, I, 1)) > 126 Then
                    Call HackingAttempt(Index, "Admin Text Modification")
                    Exit Sub
                End If
            Next

            If frmServer.chkA.Value = Unchecked Then
                Call PlayerMsg(Index, "Admin messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If

            If GetPlayerAccess(Index) > 0 Then
                Call AddLog("(admin " & GetPlayerName(Index) & ") " & Msg, ADMIN_LOG)
                Call AdminMsg("(admin " & GetPlayerName(Index) & ") " & Msg, AdminColor)
            End If
            TextAdd frmServer.txtText(5), GetPlayerName(Index) & ": " & Msg, True
            Exit Sub

        Case "playermsg"
            MsgTo = FindPlayer(Parse(1))
            Msg = Parse(2)

            ' Prevent hacking
            For I = 1 To Len(Msg)

                If Asc(Mid$(Msg, I, 1)) < 32 Or Asc(Mid$(Msg, I, 1)) > 126 Then
                    Call HackingAttempt(Index, "Player Msg Text Modification")
                    Exit Sub
                End If
            Next

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
                    Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & Msg & "'", TellColor)
                    Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
                Else
                    Call PlayerMsg(Index, "Player is not online.", White)
                End If
            Else
                Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " begins to mumble to himself, what a wierdo...", PLAYER_LOG)
                Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " begins to mumble to himself, what a wierdo...", Green)
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
            
    Case "poison"
          Call PoisonActive(Index)
          Exit Sub
          
    Case "disease"
          Call DiseaseActive(Index)
          Exit Sub
            
            ' :::::::::::::::::::::::::::
    ' :: Use Guild Deed Packet ::
    ' :::::::::::::::::::::::::::
    Case "useguilddeed"
    Dim GuildName As String
        GuildName = Trim(Parse(1))
        InvNum = Val(Parse(2))
        CharNum = Player(Index).CharNum
        
              'Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
                If Not FindGuild(GuildName) Then
                    f = FreeFile
                    Open App.Path & "\main\accounts\Guilds.txt" For Append As #f
                        Print #f, (GuildName)
                    Close #f
        
        If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_GUILDDEED Then
            If Player(Index).Char(CharNum).Guild = "" Or Player(Index).Char(CharNum).Guildaccess = 0 Then
               Call SetPlayerGuild(Index, GuildName)
               Call SetPlayerGuildAccess(Index, 4)
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
                Call SendPlayerData(Index)
               Call PlayerMsg(Index, "Your now the leader of " & GuildName & "!", BrightBlue)
                Call GlobalMsg("" & GetPlayerName(Index) & " Is now the leader of " & GuildName & "!", BrightBlue)
                'Call AddLog(GetPlayerName(index) & " created the guild " & GuildName & ". with a guild deed", GUILD_LOG)
            Else
                Call PlayerMsg(Index, "You are already in a guild!", BrightRed)
            End If
        Else
            Call PlayerMsg(Index, "You need an Guild Deed to make a guild!", BrightRed)
        End If
        Else
        Call PlayerMsg(Index, "Theres Already a Guild Named That!", BrightRed)
        End If
        Exit Sub
        
        Case "useitem"
            InvNum = Val(Parse(1))
            CharNum = Player(Index).CharNum
            Call PerformUseItem(Index, InvNum, CharNum)
            Exit Sub
        

            ' ::::::::::::::::::::::::::
            ' :: Player attack packet ::
            ' ::::::::::::::::::::::::::
        Case "attack"
            If GetPlayerWeaponSlot(Index) > 0 Then
                If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 > 0 Then
                    Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & END_CHAR)
                    Exit Sub
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
                                Damage = GetPlayerDamage(Index) - GetPlayerProtection(I) + (Rnd * 5) - 2
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & SEP_CHAR & END_CHAR)
                            Else
                                N = GetPlayerDamage(Index)
                                Damage = N + Int(Rnd * Int(N / 2)) + 1 - GetPlayerProtection(I) + (Rnd * 5) - 2
                                Call BattleMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan, 0)
                                Call BattleMsg(I, GetPlayerName(Index) & " swings with enormous might!", BrightCyan, 1)

                                'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                                'Call PlayerMsg(I, GetPlayerName(index) & " swings with enormous might!", BrightCyan)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Blow0" & SEP_CHAR & END_CHAR)
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

                            'Call PlayerMsg(index, GetPlayerName(I) & "'s " & Trim$(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked your hit!", BrightCyan)
                            'Call PlayerMsg(I, "Your " & Trim$(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                        Exit Sub
                    End If
                End If
            Next

            ' Try to attack a npc
            For I = 1 To MAX_MAP_NPCS

                ' Can we attack the npc?
                If CanAttackNpc(Index, I) Then

                    ' Get the damage we can do
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), I).num).DEF / 2) + (Rnd * 5) - 2
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & SEP_CHAR & END_CHAR)
                    Else
                        N = GetPlayerDamage(Index)
                        Damage = N + Int(Rnd * Int(N / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), I).num).DEF / 2) + (Rnd * 5) - 2
                        Call BattleMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan, 0)

                        'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Blow0" & SEP_CHAR & END_CHAR)
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
            Next
            
            
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                 If Player(I).CorpseMap = GetPlayerMap(Index) Then
                    If CanReachCorpse(Index, I) = True Then
                    Call PlayerMsg(Index, "You look into " & GetPlayerName(I) & "'s corpse.", Yellow)
                    Call SendUseCorpseTo(Index, I)
                    Exit Sub
                    End If
                 End If
                End If
            Next I
            Exit Sub
            
            Case "takecorpseloot"
            Call PickUpCorpseLoot(Index, Val(Parse(1)), Val(Parse(2)))
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
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "main\Scripts\Main.txt", "UsingStatPoints " & Index & "," & PointType
                Else

                    Select Case PointType

                        Case 0
                            Call SetPlayerstr(Index, GetPlayerstr(Index) + 1)
                            Call BattleMsg(Index, "You have gained more strength!", 15, 0)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "strengthRaised" & SEP_CHAR & END_CHAR)

                        Case 1
                            Call SetPlayerDEF(Index, GetPlayerDEF(Index) + 1)
                            Call BattleMsg(Index, "You have gained more defense!", 15, 0)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "DefenseRaised" & SEP_CHAR & END_CHAR)

                        Case 2
                            Call SetPlayerMAGI(Index, GetPlayerMAGI(Index) + 1)
                            Call BattleMsg(Index, "You have gained more magic abilities!", 15, 0)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "MagicRaised" & SEP_CHAR & END_CHAR)

                        Case 3
                            Call SetPlayerSPEED(Index, GetPlayerSPEED(Index) + 1)
                            Call BattleMsg(Index, "You have gained more speed!", 15, 0)
                            Call SendDataTo(Index, "sound" & SEP_CHAR & "SpeedRaised" & SEP_CHAR & END_CHAR)
                    End Select
                    Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)
                End If
            Else
                Call BattleMsg(Index, "You have no skill points to train with!", BrightRed, 0)
            End If
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            Call SendFP(Index)
            Call SendStats(Index)
            Exit Sub

            ' ::::::::::::::::::::::::::::::::
            ' :: Player info request packet ::
            ' ::::::::::::::::::::::::::::::::
        Case "playerinforequest"
            Name = Parse(1)
            I = FindPlayer(Name)

            If I > 0 Then
                Call PlayerMsg(Index, "Account: " & Trim$(Player(I).Login) & ", Name: " & GetPlayerName(I), BrightGreen)

                If GetPlayerAccess(Index) > ADMIN_MONITER Then
                    Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(I) & " -=-", BrightGreen)
                    Call PlayerMsg(Index, "Level: " & GetPlayerLevel(I) & "  Exp: " & GetPlayerExp(I) & "/" & GetPlayerNextLevel(I), BrightGreen)
                    Call PlayerMsg(Index, "HP: " & GetPlayerHP(I) & "/" & GetPlayerMaxHP(I) & "  MP: " & GetPlayerMP(I) & "/" & GetPlayerMaxMP(I) & "  SP: " & GetPlayerSP(I) & "/" & GetPlayerMaxSP(I), BrightGreen)
                    Call PlayerMsg(Index, "str: " & GetPlayerstr(I) & "  DEF: " & GetPlayerDEF(I) & "  MAGI: " & GetPlayerMAGI(I) & "  SPEED: " & GetPlayerSPEED(I), BrightGreen)
                    N = Int(GetPlayerstr(I) / 2) + Int(GetPlayerLevel(I) / 2)
                    I = Int(GetPlayerDEF(I) / 2) + Int(GetPlayerLevel(I) / 2)

                    If N > 100 Then N = 100
                    If I > 100 Then I = 100
                    Call PlayerMsg(Index, "Critical Hit Chance: " & N & "%, Block Chance: " & I & "%", BrightGreen)
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
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' The sprite
            N = Val(Parse(1))
            Call SetPlayerSprite(Index, N)
            Call SendPlayerData(Index)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Set player sprite packet ::
            ' ::::::::::::::::::::::::::::::
        Case "setplayersprite"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' The sprite
            I = FindPlayer(Parse(1))
            N = Val(Parse(2))
            Call SetPlayerSprite(I, N)
            Call SendPlayerData(I)
            Exit Sub

            ' ::::::::::::::::::::::::::
            ' :: Stats request packet ::
            ' ::::::::::::::::::::::::::
        Case "getstats"
            Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(Index) & " -=-", White)
            Call PlayerMsg(Index, "Level: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), White)
            Call PlayerMsg(Index, "HP: " & GetPlayerHP(Index) & "/" & GetPlayerMaxHP(Index) & "  MP: " & GetPlayerMP(Index) & "/" & GetPlayerMaxMP(Index) & "  SP: " & GetPlayerSP(Index) & "/" & GetPlayerMaxSP(Index), White)
            Call PlayerMsg(Index, "str: " & GetPlayerstr(Index) & "  DEF: " & GetPlayerDEF(Index) & "  MAGI: " & GetPlayerMAGI(Index) & "  SPEED: " & GetPlayerSPEED(Index), White)
            N = Int(GetPlayerstr(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            I = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)

            If N > 100 Then N = 100
            If I > 100 Then I = 100
            Call PlayerMsg(Index, "Critical Hit Chance: " & N & "%, Block Chance: " & I & "%", White)
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
            N = 1
            MapNum = GetPlayerMap(Index)
            Call ClearMap(MapNum)
            Map(MapNum).Name = Parse(N + 1)
            Map(MapNum).Revision = Val(Parse(N + 2)) + 1
            Map(MapNum).Moral = Val(Parse(N + 3))
            Map(MapNum).Up = Val(Parse(N + 4))
            Map(MapNum).Down = Val(Parse(N + 5))
            Map(MapNum).Left = Val(Parse(N + 6))
            Map(MapNum).Right = Val(Parse(N + 7))
            Map(MapNum).Music = Parse(N + 8)
            Map(MapNum).BootMap = Val(Parse(N + 9))
            Map(MapNum).BootX = Val(Parse(N + 10))
            Map(MapNum).BootY = Val(Parse(N + 11))
            Map(MapNum).Indoors = Val(Parse(N + 12))
            N = N + 13
            I = GetPlayerMap(Index)
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Ground = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).GroundSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Mask = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).MaskSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Anim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).AnimSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Fringe = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).FringeSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Type = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Data1 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Data2 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Data3 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).String1 = Parse(N)
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).String2 = Parse(N)
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).String3 = Parse(N)
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Mask2 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Mask2Set = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).M2Anim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).M2AnimSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).FAnim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).FAnimSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Fringe2 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Fringe2Set = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).Light = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).F2Anim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(I).Tile(x, y).F2AnimSet = Val(Parse(N))
                        N = N + 1
                    End If
                    N = N + 1
                Next
            Next
            For x = 1 To MAX_MAP_NPCS
                Map(MapNum).Npc(x) = Val(Parse(N))
                Map(MapNum).NpcSpawn(x).Used = Val(Parse(N + 1))
                Map(MapNum).NpcSpawn(x).x = Val(Parse(N + 2))
                Map(MapNum).NpcSpawn(x).y = Val(Parse(N + 3))
                N = N + 4
                Call ClearMapNpc(x, MapNum)
            Next

            ' Clear out it all
            For I = 1 To MAX_MAP_ITEMS
                Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), I).x, MapItem(GetPlayerMap(Index), I).y)
                Call ClearMapItem(I, GetPlayerMap(Index))
            Next

            ' Save the map
            Call SaveMap(MapNum)
            
            Call ResetMapGrid(MapNum)

            ' Respawn
            Call SpawnMapItems(GetPlayerMap(Index))

            ' Respawn NPCS
            Call SpawnMapNpcs(GetPlayerMap(Index))

            ' Refresh map for everyone online
            For I = 1 To MAX_PLAYERS

                If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
                    Call SendDataTo(I, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(I) & SEP_CHAR & Map(GetPlayerMap(I)).Revision & SEP_CHAR & END_CHAR)

                    Call PlayerWarp(I, MapNum, GetPlayerX(I), GetPlayerY(I))
                End If
            Next
            Exit Sub
            
            Case "sellitem"
     Dim SellItemNum As Long
     Dim SellItemSlot As Integer
     
     SellItemNum = Parse(1)
     SellItemSlot = Parse(2)
   
     If GetPlayerWeaponSlot(Index) = Val(Parse(1)) Or GetPlayerArmorSlot(Index) = Val(Parse(1)) Or GetPlayerShieldSlot(Index) = Val(Parse(1)) Or GetPlayerHelmetSlot(Index) = Val(Parse(1)) Or GetPlayerLegsSlot(Index) = Val(Parse(1)) Or GetPlayerBootsSlot(Index) = Val(Parse(1)) Or GetPlayerGlovesSlot(Index) = Val(Parse(1)) Or GetPlayerRing1Slot(Index) = Val(Parse(1)) Or GetPlayerRing2Slot(Index) = Val(Parse(1)) Or GetPlayerAmuletSlot(Index) = Val(Parse(1)) Then
         Call PlayerMsg(Index, "You cannot sell worn items.", Red)
         Exit Sub
     End If

              Call TakeItem(Index, SellItemNum, 1)
       Call GiveItem(Index, 1, Item(SellItemNum).Price)
               Call SendDataTo(Index, "updatesell" & SEP_CHAR & END_CHAR)
       Exit Sub


            ' ::::::::::::::::::::::::::::
            ' :: Need map yes/no packet ::
            ' ::::::::::::::::::::::::::::
        Case "needmap"

            ' Get yes/no value
            s = LCase$(Parse(1))

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
            If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then

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
            If Item(GetPlayerInvItemNum(Index, InvNum)).Type <> ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
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
            Call SendFP(Index)
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
                Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), I).x, MapItem(GetPlayerMap(Index), I).y)
                Call ClearMapItem(I, GetPlayerMap(Index))
            Next

            ' Respawn
            Call SpawnMapItems(GetPlayerMap(Index))

            ' Respawn NPCS
            Call SpawnMapNpcs(GetPlayerMap(Index))
            ' Reset grid
            Call ResetMapGrid(GetPlayerMap(Index))
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
            N = FindPlayer(Parse(1))

            If N <> Index Then
                If N > 0 Then
                    If GetPlayerAccess(N) <= GetPlayerAccess(Index) Then
                        Call GlobalMsg(GetPlayerName(N) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(Index) & "!", White)
                        Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(N) & ".", ADMIN_LOG)
                        Call AlertMsg(N, "You have been kicked by " & GetPlayerName(Index) & "!")
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
            N = 1
            f = FreeFile
            Open App.Path & "\banlist.txt" For Input As #f
            Do While Not EOF(f)
                Input #f, s
                Input #f, Name
                Call PlayerMsg(Index, N & ": Banned IP " & s & " by " & Name, White)
                N = N + 1
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
            N = FindPlayer(Parse(1))

            If N <> Index Then
                If N > 0 Then
                    If GetPlayerAccess(N) <= GetPlayerAccess(Index) Then
                        Call BanIndex(N, Index)
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
            
            Case "acceptquest"
        Dim CurrentQuestNum As Long
        Dim CurrentQuestNpc As Long
        
        CurrentQuestNum = Val(Parse(1))
        CurrentQuestNpc = Val(Parse(2))
        Call PlayerMsg(Index, "packet recieved -accept quest-" & CurrentQuestNum & " - " & CurrentQuestNpc, Green)
        Call ActuallyStartQuest(CurrentQuestNum, Index, CurrentQuestNpc)
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
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_ITEMS Then
                Call HackingAttempt(Index, "Invalid Item Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(Index) & " editing item #" & N & ".", ADMIN_LOG)
            Call SendEditItemTo(Index, N)
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
            N = Val(Parse(1))

            If N < 0 Or N > MAX_ITEMS Then
                Call HackingAttempt(Index, "Invalid Item Index")
                Exit Sub
            End If

            ' Update the item
            Item(N).Name = Parse(2)
            Item(N).Pic = Val(Parse(3))
            Item(N).Type = Val(Parse(4))
            Item(N).Data1 = Val(Parse(5))
            Item(N).Data2 = Val(Parse(6))
            Item(N).Data3 = Val(Parse(7))
            Item(N).StrReq = Val(Parse(8))
            Item(N).DefReq = Val(Parse(9))
            Item(N).SpeedReq = Val(Parse(10))
            Item(N).MagicReq = Val(Parse(11))
            Item(N).ClassReq = Val(Parse(12))
            Item(N).AccessReq = Val(Parse(13))
            Item(N).AddHP = Val(Parse(14))
            Item(N).AddMP = Val(Parse(15))
            Item(N).AddSP = Val(Parse(16))
            Item(N).AddStr = Val(Parse(17))
            Item(N).AddDef = Val(Parse(18))
            Item(N).AddMagi = Val(Parse(19))
            Item(N).AddSpeed = Val(Parse(20))
            Item(N).AddEXP = Val(Parse(21))
            Item(N).Desc = Parse(22)
            Item(N).AttackSpeed = Val(Parse(23))
            Item(N).Price = Val(Parse(24))
            Item(N).Stackable = Val(Parse(25))
            Item(N).Bound = Val(Parse(26))
            Item(N).LevelReq = Val(Parse(27))
            Item(N).Element = Val(Parse(28))
            Item(N).StamRemove = Val(Parse(29))
            Item(N).Rarity = Parse(30)
            Item(N).BowsReq = Val(Parse(31))
            Item(N).LargeBladesReq = Val(Parse(32))
            Item(N).SmallBladesReq = Val(Parse(33))
            Item(N).BluntWeaponsReq = Val(Parse(34))
            Item(N).PoleArmsReq = Val(Parse(35))
            Item(N).AxesReq = Val(Parse(36))
            Item(N).ThrownReq = Val(Parse(37))
            Item(N).XbowsReq = Val(Parse(38))
            Item(N).LBA = Val(Parse(39))
            Item(N).SBA = Val(Parse(40))
            Item(N).BWA = Val(Parse(41))
            Item(N).PAA = Val(Parse(42))
            Item(N).AA = Val(Parse(43))
            Item(N).TWA = Val(Parse(44))
            Item(N).XBA = Val(Parse(45))
            Item(N).BA = Val(Parse(46))
            Item(N).Poison = Val(Parse(47))
            Item(N).Disease = Val(Parse(48))
            Item(N).AilmentDamage = Val(Parse(49))
            Item(N).AilmentMS = Val(Parse(50))
            Item(N).AilmentInterval = Val(Parse(51))

            ' Save it
            Call SendUpdateItemToAll(N)
            Call SaveItem(N)
            Call AddLog(GetPlayerName(Index) & " saved item #" & N & ".", ADMIN_LOG)
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
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_NPCS Then
                Call HackingAttempt(Index, "Invalid NPC Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(Index) & " editing npc #" & N & ".", ADMIN_LOG)
            Call SendEditNpcTo(Index, N)
            
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
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_NPCS Then
                Call HackingAttempt(Index, "Invalid NPC Index")
                Exit Sub
            End If

            ' Update the npc
            Npc(N).Name = Parse(2)
            Npc(N).AttackSay = Parse(3)
            Npc(N).Sprite = Val(Parse(4))
            Npc(N).SpawnSecs = Val(Parse(5))
            Npc(N).Behavior = Val(Parse(6))
            Npc(N).Range = Val(Parse(7))
            Npc(N).STR = Val(Parse(8))
            Npc(N).DEF = Val(Parse(9))
            Npc(N).Speed = Val(Parse(10))
            Npc(N).Magi = Val(Parse(11))
            Npc(N).Big = Val(Parse(12))
            Npc(N).MaxHp = Val(Parse(13))
            Npc(N).Exp = Val(Parse(14))
            Npc(N).SpawnTime = Val(Parse(15))
            Npc(N).Speech = Val(Parse(16))
            Npc(N).Element = Val(Parse(17))
            Npc(N).Poison = Val(Parse(18))
            Npc(N).AP = Val(Parse(19))
            Npc(N).Disease = Val(Parse(20))
            Npc(N).Quest = Val(Parse(21))
            Npc(N).NpcDIR = Val(Parse(22))
            Npc(N).AilmentDamage = Val(Parse(23))
            Npc(N).AilmentInterval = Val(Parse(24))
            Npc(N).AilmentMS = Val(Parse(25))
            Npc(N).Spell = Val(Parse(26))
            z = 27
            For I = 1 To MAX_NPC_DROPS
                Npc(N).ItemNPC(I).Chance = Val(Parse(z))
                Npc(N).ItemNPC(I).ItemNum = Val(Parse(z + 1))
                Npc(N).ItemNPC(I).ItemValue = Val(Parse(z + 2))
                z = z + 3
            Next

            ' Save it
            Call SendUpdateNpcToAll(N)
            Call SaveNpc(N)
            Call AddLog(GetPlayerName(Index) & " saved npc #" & N & ".", ADMIN_LOG)
            Exit Sub
            
            Case "requesteditquest"
       Call callrequstedEditQuest(Index)
        Exit Sub

    Case "editquest"
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        N = Val(Parse(1))
        If N < 0 Or N > MAX_QUESTS Then
            Call HackingAttempt(Index, "Invalid Quest Index")
            Exit Sub
        End If
        Call AddLog(GetPlayerName(Index) & " editing quest #" & N & ".", ADMIN_LOG)
        Call SendEditQuestTo(Index, N)
        Exit Sub

    Case "savequest"
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        N = Val(Parse(1))
        If N < 0 Or N > MAX_QUESTS Then
            Call HackingAttempt(Index, "Invalid Quests Index")
            Exit Sub
        End If
        Debug.Print Parse(5) & Parse(6)
        Quest(N).Name = Parse(2)
        Quest(N).After = Parse(3)
        Quest(N).Before = Parse(4)
        Quest(N).ClassIsReq = Val(Parse(5))
        Quest(N).ClassReq = Val(Parse(6))
        Quest(N).During = Parse(7)
        Quest(N).End = Parse(8)
        Quest(N).ItemReq = Val(Parse(9))
        Quest(N).ItemVal = Val(Parse(10))
        Quest(N).LevelIsReq = Val(Parse(11))
        Quest(N).LevelReq = Val(Parse(12))
        Quest(N).NotHasItem = Parse(13)
        Quest(N).RewardNum = Val(Parse(14))
        Quest(N).RewardVal = Val(Parse(15))
        Quest(N).Start = Parse(16)
        Quest(N).StartItem = Val(Parse(17))
        Quest(N).StartOn = Val(Parse(18))
        Quest(N).Startval = Val(Parse(19))
        Quest(N).QuestExpReward = Val(Parse(20))
        Call SendUpdateQuestToAll(N)
        Call SaveQuest(N)
        Call AddLog(GetPlayerName(Index) & " saved quest #" & N & ".", ADMIN_LOG)
        Exit Sub

    Case "questdone"
      Call GiveRewardItem(Index, Quest(Val(Parse(1))).RewardNum, Quest(Val(Parse(1))).RewardVal, Val(Parse(3)))
            Exit Sub
            
    Case "vault"
            Dim VaultPass As String
            VaultPass = Parse(1)
               Call VaultVerify(Index, VaultPass)
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
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_SHOPS Then
                Call HackingAttempt(Index, "Invalid Shop Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(Index) & " editing shop #" & N & ".", ADMIN_LOG)
            Call SendEditShopTo(Index, N)
            Exit Sub

        Case "addfriend"
            Name = Trim$(Parse(1))

            If Not FindChar(Name) Then
                Call PlayerMsg(Index, "No such player exists!", Blue)
                Exit Sub
            End If

            If Name = GetPlayerName(Index) Then
                Call PlayerMsg(Index, "You can't add yourself!", Blue)
                Exit Sub
            End If
            For I = 1 To MAX_FRIENDS

                If Player(Index).Char(Player(Index).CharNum).Friends(I) = Name Then
                    Call PlayerMsg(Index, "You already have that user as a friend!", Blue)
                    Exit Sub
                End If
            Next
            For I = 1 To MAX_FRIENDS

                If Player(Index).Char(Player(Index).CharNum).Friends(I) = "" Then
                    Player(Index).Char(Player(Index).CharNum).Friends(I) = Name
                    Call PlayerMsg(Index, "Friend added.", Blue)
                    Call SendFriendListTo(Index)
                    Exit Sub
                End If
            Next
            Call PlayerMsg(Index, "Sorry, but you have too many friends already.", Blue)
            Exit Sub

        Case "removefriend"
            Name = Trim$(Parse(1))
            For I = 1 To MAX_FRIENDS

                If Player(Index).Char(Player(Index).CharNum).Friends(I) = Name Then
                    Player(Index).Char(Player(Index).CharNum).Friends(I) = ""
                    Call PlayerMsg(Index, "Friend removed.", Blue)
                    Call SendFriendListTo(Index)
                    Exit Sub
                End If
            Next
            Call PlayerMsg(Index, "That person isn't on your friend list!", Blue)
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
            N = 6
            For z = 1 To 6
                For I = 1 To MAX_TRADES
                    Shop(ShopNum).TradeItem(z).Value(I).GiveItem = Val(Parse(N))
                    Shop(ShopNum).TradeItem(z).Value(I).GiveValue = Val(Parse(N + 1))
                    Shop(ShopNum).TradeItem(z).Value(I).GetItem = Val(Parse(N + 2))
                    Shop(ShopNum).TradeItem(z).Value(I).GetValue = Val(Parse(N + 3))
                    N = N + 4
                Next
            Next

            ' Save it
            Call SendUpdateShopToAll(ShopNum)
            Call SaveShop(ShopNum)
            Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Request edit main packet ::
            ' ::::::::::::::::::::::::::::::
        Case "requesteditmain"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_CREATOR Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            
            f = FreeFile
            Open App.Path & "\main\Scripts\Main.txt" For Input As #f
                Call SendDataTo(Index, "MAINEDITOR" & SEP_CHAR & Input$(LOF(f), f) & SEP_CHAR & END_CHAR)
            Close #f
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
        frmServer.GameTimeSpeed.text = 0
        TimeDisable = True
        frmServer.Command69.Caption = "Enable Time"
    Else
        Gamespeed = 1
        frmServer.GameTimeSpeed.text = 1
        TimeDisable = False
        frmServer.Command69.Caption = "Disable Time"
    End If
            
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
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_SPELLS Then
                Call HackingAttempt(Index, "Invalid Spell Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(Index) & " editing spell #" & N & ".", ADMIN_LOG)
            Call SendEditSpellTo(Index, N)
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
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_SPELLS Then
                Call HackingAttempt(Index, "Invalid Spell Index")
                Exit Sub
            End If

            ' Update the spell
            Spell(N).Name = Parse(2)
            Spell(N).ClassReq = Val(Parse(3))
            Spell(N).LevelReq = Val(Parse(4))
            Spell(N).Type = Val(Parse(5))
            Spell(N).Data1 = Val(Parse(6))
            Spell(N).Data2 = Val(Parse(7))
            Spell(N).Data3 = Val(Parse(8))
            Spell(N).MPCost = Val(Parse(9))
            Spell(N).sound = Val(Parse(10))
            Spell(N).Range = Val(Parse(11))
            Spell(N).SpellAnim = Val(Parse(12))
            Spell(N).SpellTime = Val(Parse(13))
            Spell(N).SpellDone = Val(Parse(14))
            Spell(N).AE = Val(Parse(15))
            Spell(N).Pic = Val(Parse(16))
            Spell(N).Element = Val(Parse(17))

            ' Save it
            Call SendUpdateSpellToAll(N)
            Call SaveSpell(N)
            Call AddLog(GetPlayerName(Index) & " saving spell #" & N & ".", ADMIN_LOG)
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
            N = FindPlayer(Parse(1))

            ' The access
            I = Val(Parse(2))

            ' Check for invalid access level
            If I >= 0 Or I <= 3 Then
                If GetPlayerName(Index) <> GetPlayerName(N) Then
                    If GetPlayerAccess(Index) > GetPlayerAccess(N) Then

                        ' Check if player is on
                        If N > 0 Then
                            If GetPlayerAccess(N) <= 0 Then
                                Call GlobalMsg(GetPlayerName(N) & " has been blessed with administrative access.", BrightBlue)
                            End If
                            Call SetPlayerAccess(N, I)
                            Call SendPlayerData(N)
                            Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(N) & "'s access.", ADMIN_LOG)
                        Else
                            Call PlayerMsg(Index, "Player is not online.", White)
                        End If
                    Else
                        Call PlayerMsg(Index, "Your access level is lower than " & GetPlayerName(N) & "´s.", Red)
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
            
            Case "57"
        If GetPlayerParty(Index) = 0 Then
        Call PlayerMsg(Index, "Hacking attempt.", Yellow)
        Exit Sub
        End If
        If Player(Index).TargetType = TARGET_TYPE_NPC Then
        Call PlayerMsg(Index, "Instead of being a loser trying to make friends with NPCs, why don't you go out and get some real friends?", Yellow)
        Exit Sub
        End If
        If IsPlaying(Player(Index).Target) = False Then
        Call PlayerMsg(Index, "This player is not playing at the moment.", Yellow)
        Exit Sub
        End If
        If Index = Player(Index).Target Then
        Call PlayerMsg(Index, "You cannot invite yourself into your own party... since well... you're already in it? Maybe you should try inviting others...", Yellow)
        Exit Sub
        End If
        If Not GetPlayerParty(Player(Index).Target) = 0 Then
        Call PlayerMsg(Index, "This player is already in another party at the moment.", Yellow)
        Exit Sub
        End If
        If GetPlayerLevel(Index) > (GetPlayerLevel(Player(Index).Target) + 3) Or GetPlayerLevel(Index) < (GetPlayerLevel(Player(Index).Target) - 3) Then
        Call PlayerMsg(Index, "You cannot invite anyone three levels above or below you.", Yellow)
        Exit Sub
        End If
        If Not Party(GetPlayerParty(Index)).Leader = Index Then
        Call PlayerMsg(Index, "Only the leader of the party can invite people to join it.", Yellow)
        Else
         Call InvitePlayerToParty(Index, Player(Index).Target)
       End If
        Exit Sub

' :::::::::::::::::::
' :: Invite packet ::
' :::::::::::::::::::
   Case "58"
   N = FindPlayer(Parse(1))
   If GetPlayerMap(N) <> GetPlayerMap(Index) Then
   Call PlayerMsg(Index, "The player you're inviting must be on the same map as you to receive an invitation.", Yellow)
   Exit Sub
   End If
   Call InvitePlayerToParty(Index, N)
   Exit Sub
   
   Case "f6"
   If GetPlayerParty(Index) = 0 Then
   Call PlayerMsg(Index, "You are not in a party!", Red)
   Exit Sub
   End If
   If Party(GetPlayerParty(Index)).Member(Val(Parse(1))) = 0 Then
   Call PlayerMsg(Index, "This party member does not exist", Red)
   Exit Sub
   End If
   
   If GetPlayerMap(Party(GetPlayerParty(Index)).Member(Val(Parse(1)))) = GetPlayerMap(Index) Then
   Player(Index).Target = Party(GetPlayerParty(Index)).Member(Val(Parse(1)))
   Player(Index).TargetType = TARGET_TYPE_PLAYER
   Call PlayerMsg(Index, "Your target is now " & GetPlayerName(Party(GetPlayerParty(Index)).Member(Val(Parse(1)))) & ".", Green)
   End If
   Exit Sub
   
   
   ' ::::::::::::::::::::::::
    ' :: Leave party packet ::
    ' ::::::::::::::::::::::::
    Case "64"
    ' Check if the player was in the party, if so, remove them.
    If GetPlayerParty(Index) > 0 Then
    Call PartyRemoval(Index, GetPlayerParty(Index), Trim$(GetPlayerName(Index)))
    Exit Sub
    Else
    Call PlayerMsg(Index, "You are not in a party at the moment.", Red)
    Exit Sub
    End If
    
    Case "62"
    ' A player has accepted the invitation.
    
    If GetPlayerInvited(Index) = 0 Then
    Call PlayerMsg(Index, "You cannot a party since you were never invited to one.", Red)
    Exit Sub
    End If
    
    If Party(GetPlayerInvited(Index)).Created = False And GetPlayerInvited(Index) > 0 Then
    Call PlayerMsg(Index, "You cannot join this party since it no longer exists.", Red)
    Player(Index).Char(Player(Index).CharNum).PartyInvitedTo = 0
    Player(Index).Char(Player(Index).CharNum).PartyInvitedToBy = ""
    Exit Sub
    End If
    
    If Not Trim$(Player(Index).Char(Player(Index).CharNum).PartyInvitedToBy) = Trim$(GetPlayerName(Party(GetPlayerInvited(Index)).Leader)) Then
    Call PlayerMsg(Index, "The person who invited you from this party is no longer the leader, thus this invitation is invalid.", Red)
    Player(Index).Char(Player(Index).CharNum).PartyInvitedTo = 0
    Player(Index).Char(Player(Index).CharNum).PartyInvitedToBy = ""
    Exit Sub
    End If
    
    If Not GetPlayerMap(Index) = GetPlayerMap(Party(GetPlayerInvited(Index)).Leader) Then
    Call PlayerMsg(Index, "You must be on the same map as the party leader in order to accept an invitation.", Red)
    Exit Sub
    End If
    
    If Party(GetPlayerInvited(Index)).Created = True Then
       Call SetPlayerParty(Index, GetPlayerInvited(Index))
       Player(Index).Char(Player(Index).CharNum).PartyInvitedTo = 0
        Player(Index).Char(Player(Index).CharNum).PartyInvitedToBy = ""
        Exit Sub
    End If
    
    
            
            Case "c7"
        If GetPlayerParty(Index) = 0 Then
        Call PlayerMsg(Index, "Hacking attempt.", Red)
        Exit Sub
        End If
        If Player(Index).TargetType = TARGET_TYPE_NPC Then
        Call PlayerMsg(Index, "Why would an NPC even join the likes of you in the first place?", Yellow)
        Exit Sub
        End If
        If IsPlaying(Player(Index).Target) = False Then
        Call PlayerMsg(Index, "This player is not playing at the moment.", Yellow)
        Exit Sub
        End If
        If Not GetPlayerParty(Player(Index).Target) = GetPlayerParty(Index) Then
        Call PlayerMsg(Index, "This player is not in your party at the moment.", Yellow)
        Exit Sub
        End If
        If Index = Player(Index).Target Then
        Call PlayerMsg(Index, "You cannot remove yourself from the party with this option. To leave, type either /leave or click the leave button in the party menu.", Yellow)
        Exit Sub
        End If
        If Not Party(GetPlayerParty(Index)).Leader = Index Then
        Call PlayerMsg(Index, "Only the leader of the party can remove members from it.", Yellow)
        Else
       ' Call PartyMsg(GetPlayerParty(Index), "The leader of the party, " & GetPlayerName(Index) & ", has removed " & GetPlayerName(Player(Index).Target) & " from the party.", Yellow)
        Call PartyRemoval(Player(Index).Target, GetPlayerParty(Index), GetPlayerName(Player(Index).Target))
       End If
        Exit Sub
    
    
    Case "i2"
        Call CreateParty(Index)
        Exit Sub

        Case "onlinelist"
            Call SendOnlineList
            Exit Sub

        Case "setmotd"

            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            Call SpecialPutVar(App.Path & "\motd.ini", "MOTD", "Msg", Parse(1))
            Call GlobalMsg("MOTD changed to: " & Parse(1), BrightCyan)
            Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Parse(1), ADMIN_LOG)
            Exit Sub

        Case "traderequest"
        ' Trade num
        N = Val(Parse(1))
        z = Val(Parse(2))
        
        ' Prevent hacking
        If (N < 1) Or (N > 6) Then
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
        'Check if its furniture
        If Item(Shop(I).TradeItem(N).Value(z).GetItem).Type = ITEM_TYPE_FURNITURE Then
        If HasItem(Index, Shop(I).TradeItem(N).Value(z).GiveItem) >= Shop(I).TradeItem(N).Value(z).GiveValue And Player(Index).Char(Player(Index).CharNum).Hands = 0 Then
        Call TakeItem(Index, Shop(I).TradeItem(N).Value(z).GiveItem, Shop(I).TradeItem(N).Value(z).GiveValue)
        Player(Index).Char(Player(Index).CharNum).Hands = Shop(I).TradeItem(N).Value(z).GetItem
        Call SendDataTo(Index, "Sethands" & SEP_CHAR & Shop(I).TradeItem(N).Value(z).GetItem & SEP_CHAR & END_CHAR)
        Call PlayerMsg(Index, "The trade was successful!", Yellow)
        Exit Sub
                Else
            Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
            Exit Sub
        End If
        End If
        
        ' Check if inv full
        If I <= 0 Then Exit Sub
        x = FindOpenInvSlot(Index, Shop(I).TradeItem(N).Value(z).GetItem)
        If x = 0 Then
            Call PlayerMsg(Index, "Trade unsuccessful, inventory full.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have the item
        If HasItem(Index, Shop(I).TradeItem(N).Value(z).GiveItem) >= Shop(I).TradeItem(N).Value(z).GiveValue Then
            Call TakeItem(Index, Shop(I).TradeItem(N).Value(z).GiveItem, Shop(I).TradeItem(N).Value(z).GiveValue)
            Call GiveItem(Index, Shop(I).TradeItem(N).Value(z).GetItem, Shop(I).TradeItem(N).Value(z).GetValue)
            Call PlayerMsg(Index, "The trade was successful!", Yellow)
        Else
            Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
        End If
        Exit Sub

        Case "fixitem"

            ' Inv num
            N = Val(Parse(1))

            ' Make sure its a equipable item
            If Item(GetPlayerInvItemNum(Index, N)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(Index, N)).Type > ITEM_TYPE_SHIELD Or Item(GetPlayerInvItemNum(Index, N)).Type > ITEM_TYPE_LEGS Or Item(GetPlayerInvItemNum(Index, N)).Type > ITEM_TYPE_BOOTS Or Item(GetPlayerInvItemNum(Index, N)).Type > ITEM_TYPE_GLOVES Or Item(GetPlayerInvItemNum(Index, N)).Type > ITEM_TYPE_RING1 Or Item(GetPlayerInvItemNum(Index, N)).Type > ITEM_TYPE_RING2 Or Item(GetPlayerInvItemNum(Index, N)).Type > ITEM_TYPE_AMULET Then
                Call PlayerMsg(Index, "You can only fix weapons, armors, helmets, and shields, legs, boots, gloves, rings And amulets.", BrightRed)
                Exit Sub
            End If

            ' Check if they have a full inventory
            If FindOpenInvSlot(Index, GetPlayerInvItemNum(Index, N)) <= 0 Then
                Call PlayerMsg(Index, "You have no inventory space left!", BrightRed)
                Exit Sub
            End If

            ' Check if you can actually repair the item
            If Item(ItemNum).Data1 < 0 Then
                Call PlayerMsg(Index, "This item isn't repairable!", BrightRed)
                Exit Sub
            End If

            ' Now check the rate of pay
            ItemNum = GetPlayerInvItemNum(Index, N)
            I = Int(Item(GetPlayerInvItemNum(Index, N)).Data2 / 5)

            If I <= 0 Then I = 1
            DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(Index, N)
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
                    Call SetPlayerInvItemDur(Index, N, Item(ItemNum).Data1 * -1)
                    Call PlayerMsg(Index, "Item has been totally restored for " & GoldNeeded & " gold!", BrightBlue)
                Else

                    ' They dont so restore as much as we can
                    DurNeeded = (HasItem(Index, 1) / I)
                    GoldNeeded = Int(DurNeeded * I / 2)

                    If GoldNeeded <= 0 Then GoldNeeded = 1
                    Call TakeItem(Index, 1, GoldNeeded)
                    Call SetPlayerInvItemDur(Index, N, GetPlayerInvItemDur(Index, N) + DurNeeded)
                    Call PlayerMsg(Index, "Item has been partially fixed for " & GoldNeeded & " gold!", BrightBlue)
                End If
            Else
                Call PlayerMsg(Index, "Insufficient gold to fix this item!", BrightRed)
            End If
            Exit Sub

        Case "search"
            x = Val(Parse(1))
            y = Val(Parse(2))

            ' Prevent subscript out of range
            If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then
                Exit Sub
            End If

            ' Check for a player
            For I = 1 To MAX_PLAYERS

                If IsPlaying(I) And GetPlayerMap(Index) = GetPlayerMap(I) And GetPlayerX(I) = x And GetPlayerY(I) = y Then

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
            Next

            ' Check for an npc
            For I = 1 To MAX_MAP_NPCS

                If MapNpc(GetPlayerMap(Index), I).num > 0 Then
                    If MapNpc(GetPlayerMap(Index), I).x = x And MapNpc(GetPlayerMap(Index), I).y = y Then

                        ' Change target
                        Player(Index).Target = I
                        Player(Index).TargetType = TARGET_TYPE_NPC
                        Call PlayerMsg(Index, "Your target is now a " & Trim$(Npc(MapNpc(GetPlayerMap(Index), I).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If
            Next

            ' Check for an item
            For I = 1 To MAX_MAP_ITEMS

                If MapItem(GetPlayerMap(Index), I).num > 0 Then
                    If MapItem(GetPlayerMap(Index), I).x = x And MapItem(GetPlayerMap(Index), I).y = y Then
                        Call PlayerMsg(Index, "You see a " & Trim$(Item(MapItem(GetPlayerMap(Index), I).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If
            Next
            Exit Sub

        Case "playerchat"
            N = FindPlayer(Parse(1))

            If N < 1 Then
                Call PlayerMsg(Index, "Player is not online.", White)
                Exit Sub
            End If

            If N = Index Then
                Exit Sub
            End If

            If Player(Index).InChat = 1 Then
                Call PlayerMsg(Index, "Your already in a chat with another player!", Pink)
                Exit Sub
            End If

            If Player(N).InChat = 1 Then
                Call PlayerMsg(Index, "Player is already in a chat with another player!", Pink)
                Exit Sub
            End If
            Call PlayerMsg(Index, "Chat request has been sent to " & GetPlayerName(N) & ".", Pink)
            Call PlayerMsg(N, GetPlayerName(Index) & " wants you to chat with them.  Type /chat to accept, or /chatdecline to decline.", Pink)
            Player(N).ChatPlayer = Index
            Player(Index).ChatPlayer = N
            Exit Sub

        Case "achat"
            N = Player(Index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
                Exit Sub
            End If

            If Player(N).ChatPlayer <> Index Then
                Call PlayerMsg(Index, "Chat failed.", Pink)
                Exit Sub
            End If
            Call SendDataTo(Index, "PPCHATTING" & SEP_CHAR & N & SEP_CHAR & END_CHAR)
            Call SendDataTo(N, "PPCHATTING" & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
            Exit Sub
            
        Case "forgetspell"
' Spell slot
N = CLng(Parse(1))

' Prevent subscript out of range
If N <= 0 Or N > MAX_PLAYER_SPELLS Then
Call HackingAttempt(Index, "Invalid Spell Slot")
Exit Sub
End If

With Player(Index).Char(Player(Index).CharNum)
If .Spell(N) = 0 Then
Call PlayerMsg(Index, "No spell here.", Red)

Else
Call PlayerMsg(Index, "You have forgotten the spell """ & Trim$(Spell(.Spell(N)).Name) & """", Green)

.Spell(N) = 0
Call SendSpells(Index)
End If
End With
Exit Sub

        Case "dchat"
            N = Player(Index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
                Exit Sub
            End If
            Call PlayerMsg(Index, "Declined chat request.", Pink)
            Call PlayerMsg(N, GetPlayerName(Index) & " declined your request.", Pink)
            Player(Index).ChatPlayer = 0
            Player(Index).InChat = 0
            Player(N).ChatPlayer = 0
            Player(N).InChat = 0
            Exit Sub

        Case "qchat"
            N = Player(Index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
                Exit Sub
            End If
            Call SendDataTo(Index, "qchat" & SEP_CHAR & END_CHAR)
            Call SendDataTo(N, "qchat" & SEP_CHAR & END_CHAR)
            Player(Index).ChatPlayer = 0
            Player(Index).InChat = 0
            Player(N).ChatPlayer = 0
            Player(N).InChat = 0
            Exit Sub

        Case "sendchat"
            N = Player(Index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
                Exit Sub
            End If
            Call SendDataTo(N, "sendchat" & SEP_CHAR & Parse(1) & SEP_CHAR & Index & SEP_CHAR & END_CHAR)
            Exit Sub

        Case "pptrade"
            N = FindPlayer(Parse(1))

            ' Check if player is online
            If N < 1 Then
                Call PlayerMsg(Index, "Player is not online.", White)
                Exit Sub
            End If

            ' Prevent trading with self
            If N = Index Then
                Exit Sub
            End If

            ' Check if the player is in another trade
            If Player(Index).InTrade = 1 Then
                Call PlayerMsg(Index, "Your already in a trade with someone else!", Pink)
                Exit Sub
            End If

            For I = 0 To 3
                If DirToX(GetPlayerX(Index), I) = GetPlayerX(N) And DirToY(GetPlayerY(Index), I) = GetPlayerY(N) Then
                    ' Check to see if player is already in a trade
                    If Player(N).InTrade = 1 Then
                        Call PlayerMsg(Index, "Player is already in a trade!", Pink)
                        Exit Sub
                    End If
                    Call PlayerMsg(Index, "Trade request has been sent to " & GetPlayerName(N) & ".", Pink)
                    Call PlayerMsg(N, GetPlayerName(Index) & " wants you to trade with them.  Type /accept to accept, or /decline to decline.", Pink)
                    Player(N).TradePlayer = Index
                    Player(Index).TradePlayer = N
                    Exit Sub
                End If
            Next
            
            Call PlayerMsg(Index, "You need to be beside the player to trade!", Pink)
            Call PlayerMsg(N, "The player needs to be beside you to trade!", Pink)
            Exit Sub

        Case "atrade"
            N = Player(Index).TradePlayer

            ' Check if anyone requested a trade
            If N < 1 Then
                Call PlayerMsg(Index, "No one requested a trade with you.", Pink)
                Exit Sub
            End If

            ' Check if its the right player
            If Player(N).TradePlayer <> Index Then
                Call PlayerMsg(Index, "Trade failed.", Pink)
                Exit Sub
            End If

            ' Check where both players are
            For I = 0 To 3
                If DirToX(GetPlayerX(Index), I) = GetPlayerX(N) And DirToY(GetPlayerY(Index), I) = GetPlayerY(N) Then
                    Call PlayerMsg(Index, "You are trading with " & GetPlayerName(N) & "!", Pink)
                    Call PlayerMsg(N, GetPlayerName(Index) & " accepted your trade request!", Pink)
                    Call SendDataTo(Index, "PPTRADING" & SEP_CHAR & END_CHAR)
                    Call SendDataTo(N, "PPTRADING" & SEP_CHAR & END_CHAR)
                    For o = 1 To MAX_PLAYER_TRADES
                        Player(Index).Trading(o).InvNum = 0
                        Player(Index).Trading(o).InvName = ""
                        Player(N).Trading(o).InvNum = 0
                        Player(N).Trading(o).InvName = ""
                    Next
                    Player(Index).InTrade = 1
                    Player(Index).TradeItemMax = 0
                    Player(Index).TradeItemMax2 = 0
                    Player(N).InTrade = 1
                    Player(N).TradeItemMax = 0
                    Player(N).TradeItemMax2 = 0
                    Exit Sub
                End If
            Next
            
            Call PlayerMsg(Index, "The player needs to be beside you to trade!", Pink)
            Call PlayerMsg(N, "You need to be beside the player to trade!", Pink)
            Exit Sub

        Case "qtrade"
            N = Player(Index).TradePlayer

            ' Check if anyone trade with player
            If N < 1 Then
                Call PlayerMsg(Index, "No one requested a trade with you.", Pink)
                Exit Sub
            End If
            Call PlayerMsg(Index, "Stopped trading.", Pink)
            Call PlayerMsg(N, GetPlayerName(Index) & " stopped trading with you!", Pink)
            Player(Index).TradeOk = 0
            Player(N).TradeOk = 0
            Player(Index).TradePlayer = 0
            Player(Index).InTrade = 0
            Player(N).TradePlayer = 0
            Player(N).InTrade = 0
            Call SendDataTo(Index, "qtrade" & SEP_CHAR & END_CHAR)
            Call SendDataTo(N, "qtrade" & SEP_CHAR & END_CHAR)
            Exit Sub

        Case "dtrade"
            N = Player(Index).TradePlayer

            ' Check if anyone trade with player
            If N < 1 Then
                Call PlayerMsg(Index, "No one requested a trade with you.", Pink)
                Exit Sub
            End If
            Call PlayerMsg(Index, "Declined trade request.", Pink)
            Call PlayerMsg(N, GetPlayerName(Index) & " declined your request.", Pink)
            Player(Index).TradePlayer = 0
            Player(Index).InTrade = 0
            Player(N).TradePlayer = 0
            Player(N).InTrade = 0
            Exit Sub

        Case "updatetradeinv"
            N = Val(Parse(1))
            Player(Index).Trading(N).InvNum = Val(Parse(2))
            Player(Index).Trading(N).InvName = Trim$(Parse(3))
            Player(Index).Trading(N).InvVal = Val(Parse(4))
            If Player(Index).Trading(N).InvNum = 0 Then
                Player(Index).TradeItemMax = Player(Index).TradeItemMax - 1
                Player(Index).TradeOk = 0
                Player(N).TradeOk = 0
                Call SendDataTo(Index, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                Call SendDataTo(N, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Else
                Player(Index).TradeItemMax = Player(Index).TradeItemMax + 1
            End If
            Call SendDataTo(Player(Index).TradePlayer, "updatetradeitem" & SEP_CHAR & N & SEP_CHAR & Player(Index).Trading(N).InvNum & SEP_CHAR & Player(Index).Trading(N).InvName & SEP_CHAR & Player(Index).Trading(N).InvVal & SEP_CHAR & END_CHAR)
            Exit Sub


        Case "swapitems"
            Dim cur As Long
            
            N = Player(Index).TradePlayer

            If Player(Index).TradeOk = 0 Then
                Player(Index).TradeOk = 1
                Call SendDataTo(N, "trading" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            ElseIf Player(Index).TradeOk = 1 Then
                Player(Index).TradeOk = 0
                Call SendDataTo(N, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            End If

            If Player(Index).TradeOk = 1 And Player(N).TradeOk = 1 Then
                Player(Index).TradeItemMax2 = 0
                Player(N).TradeItemMax2 = 0
                For I = 1 To MAX_INV

                    If Player(Index).TradeItemMax = Player(Index).TradeItemMax2 Then
                        Exit For
                    End If

                    If GetPlayerInvItemNum(N, I) < 1 Then
                        Player(Index).TradeItemMax2 = Player(Index).TradeItemMax2 + 1
                    End If
                Next
                For I = 1 To MAX_INV

                    If Player(N).TradeItemMax = Player(N).TradeItemMax2 Then
                        Exit For
                    End If

                    If GetPlayerInvItemNum(Index, I) < 1 Then
                        Player(N).TradeItemMax2 = Player(N).TradeItemMax2 + 1
                    End If
                Next

                If Player(Index).TradeItemMax2 = Player(Index).TradeItemMax And Player(N).TradeItemMax2 = Player(N).TradeItemMax Then
                    For I = 1 To MAX_PLAYER_TRADES
                        For x = 1 To MAX_INV

                            If GetPlayerInvItemNum(N, x) < 1 Then
                                If Player(Index).Trading(I).InvNum > 0 Then
                                    If Player(Index).Trading(I).InvVal > 0 Then
                                    Call GiveItem(N, GetPlayerInvItemNum(Index, Player(Index).Trading(I).InvNum), Player(Index).Trading(I).InvVal)
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, Player(Index).Trading(I).InvNum), Player(Index).Trading(I).InvVal)
                                    Else
                                    Call GiveItem(N, GetPlayerInvItemNum(Index, Player(Index).Trading(I).InvNum), 1)
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, Player(Index).Trading(I).InvNum), 1)
                                    End If
                                    Exit For
                                End If
                            End If
                        Next
                    Next
                    For I = 1 To MAX_PLAYER_TRADES
                        For x = 1 To MAX_INV

                            If GetPlayerInvItemNum(Index, x) < 1 Then
                                If Player(N).Trading(I).InvNum > 0 Then
                                    If Player(N).Trading(I).InvVal > 0 Then
                                    Call GiveItem(Index, GetPlayerInvItemNum(N, Player(N).Trading(I).InvNum), Player(N).Trading(I).InvVal)
                                    Call TakeItem(N, GetPlayerInvItemNum(N, Player(N).Trading(I).InvNum), Player(N).Trading(I).InvVal)
                                    Else
                                    Call GiveItem(Index, GetPlayerInvItemNum(N, Player(N).Trading(I).InvNum), 1)
                                    Call TakeItem(N, GetPlayerInvItemNum(N, Player(N).Trading(I).InvNum), 1)
                                    End If
                                    Exit For
                                End If
                            End If
                        Next
                    Next
                    Call PlayerMsg(N, "Trade Successfull!", BrightGreen)
                    Call PlayerMsg(Index, "Trade Successfull!", BrightGreen)
                    Call SendInventory(N)
                    Call SendInventory(Index)
                Else

                    If Player(Index).TradeItemMax2 < Player(Index).TradeItemMax Then
                        Call PlayerMsg(Index, "Your inventory is full!", BrightRed)
                        Call PlayerMsg(N, GetPlayerName(Index) & "'s inventory is full!", BrightRed)
                    End If

                    If Player(N).TradeItemMax2 < Player(N).TradeItemMax Then
                        Call PlayerMsg(N, "Your inventory is full!", BrightRed)
                        Call PlayerMsg(Index, GetPlayerName(N) & "'s inventory is full!", BrightRed)
                    End If
                End If
                Player(Index).TradePlayer = 0
                Player(Index).InTrade = 0
                Player(Index).TradeOk = 0
                Player(N).TradePlayer = 0
                Player(N).InTrade = 0
                Player(N).TradeOk = 0
                Call SendDataTo(Index, "qtrade" & SEP_CHAR & END_CHAR)
                Call SendDataTo(N, "qtrade" & SEP_CHAR & END_CHAR)
            End If
            Exit Sub


        Case "party"
        N = FindPlayer(Parse(1))
        
        ' Prevent partying with self
        If N = Index Then
            Exit Sub
        End If
                
        ' Check for a previous party and if so drop it
        If Player(Index).InParty = YES Then
            Call PlayerMsg(Index, "You are already in a party!", Pink)
            Exit Sub
        End If
        
        If N > 0 Then
            ' Check if its an admin
            If GetPlayerAccess(Index) > ADMIN_MONITER Then
                Call PlayerMsg(Index, "You can't join a party, you are an admin!", BrightBlue)
                Exit Sub
            End If
        
            If GetPlayerAccess(N) > ADMIN_MONITER Then
                Call PlayerMsg(Index, "Admins cannot join parties!", BrightBlue)
                Exit Sub
            End If
            
            ' Make sure they are in right level range
            If GetPlayerLevel(Index) + 5 < GetPlayerLevel(N) Or GetPlayerLevel(Index) - 5 > GetPlayerLevel(N) Then
                Call PlayerMsg(Index, "There is more then a 5 level gap between you two, party failed.", Pink)
                Exit Sub
            End If
            
            ' Check to see if player is already in a party
            If Player(N).InParty = NO Then
                Call PlayerMsg(Index, "Party request has been sent to " & GetPlayerName(N) & ".", Pink)
                Call PlayerMsg(N, GetPlayerName(Index) & " wants you to join their party.  Type /join to join, or /leave to decline.", Pink)
            
                Player(Index).PartyStarter = YES
                Player(Index).PartyPlayer = N
                Player(N).PartyPlayer = Index
            Else
                Call PlayerMsg(Index, "Player is already in a party!", Pink)
            End If
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If
        Exit Sub

    ' :::::::::::::::::::::::
    ' :: Join party packet ::
    ' :::::::::::::::::::::::
    Case "joinparty"
        N = Player(Index).PartyPlayer
        
        If N > 0 Then
            ' Check to make sure they aren't the starter
            If Player(Index).PartyStarter = NO Then
                ' Check to make sure that each of there party players match
                If Player(N).PartyPlayer = Index Then
                    Call PlayerMsg(Index, "You have joined " & GetPlayerName(N) & "'s party!", Pink)
                    Call PlayerMsg(N, GetPlayerName(Index) & " has joined your party!", Pink)
                    
                    Player(Index).InParty = YES
                    Player(N).InParty = YES
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
   
    Case "leaveparty"
        N = Player(Index).PartyPlayer
        
        If N > 0 Then
            If Player(Index).InParty = YES Then
                Call PlayerMsg(Index, "You have left the party.", Pink)
                Call PlayerMsg(N, GetPlayerName(Index) & " has left the party.", Pink)
                
                Player(Index).PartyPlayer = 0
                Player(Index).PartyStarter = NO
                Player(Index).InParty = NO
                Player(N).PartyPlayer = 0
                Player(N).PartyStarter = NO
                Player(N).InParty = NO
            Else
                Call PlayerMsg(Index, "Declined party request.", Pink)
                Call PlayerMsg(N, GetPlayerName(Index) & " declined your request.", Pink)
                
                Player(Index).PartyPlayer = 0
                Player(Index).PartyStarter = NO
                Player(Index).InParty = NO
                Player(N).PartyPlayer = 0
                Player(N).PartyStarter = NO
                Player(N).InParty = NO
            End If
        Else
            Call PlayerMsg(Index, "You are not in a party!", Pink)
        End If
        Exit Sub

        Case "partychat"

            If Player(Index).PartyID > 0 Then
                For I = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Index).PartyID).Member(I) <> 0 Then Call PlayerMsg(Party(Player(Index).PartyID).Member(I), GetPlayerName(Index) & "-" & Parse(1), PartyColor)
                Next
            Else
                Call PlayerMsg(Index, "You are not in a party!", Pink)
            End If
            Exit Sub

        Case "guildchat"

            If GetPlayerGuild(Index) <> "" Then
                For I = 1 To MAX_PLAYERS

                    If GetPlayerGuild(Index) = GetPlayerGuild(I) Then Call PlayerMsg(I, GetPlayerGuild(Index) & "-" & GetPlayerName(Index) & ": " & Parse(1), GuildColor)
                Next
            Else
                Call PlayerMsg(Index, "You are not in a guild!", Pink)
            End If
            Exit Sub

        Case "newmain"

            If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
                Dim temp As String

                f = FreeFile
                Open App.Path & "\main\Scripts\Main.txt" For Input As #f
                temp = Input$(LOF(f), f)
                Close #f
                f = FreeFile
                Open App.Path & "\main\Scripts\Backup.txt" For Output As #f
                Print #f, temp
                Close #f
                f = FreeFile
                Open App.Path & "\main\Scripts\Main.txt" For Output As #f
                Print #f, Parse(1)
                Close #f

                If SCRIPTING = 1 Then
                    Set MyScript = Nothing
                    Set clsScriptCommands = Nothing
                    Set MyScript = New clsSadScript
                    Set clsScriptCommands = New clsCommands
                    MyScript.ReadInCode App.Path & "\main\Scripts\Main.txt", "main\Scripts\Main.txt", MyScript.SControl, False
                    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
                    Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
                    Call PlayerMsg(Index, "Scripts reloaded.", White)
                End If
                Call AddLog(GetPlayerName(Index) & " updated the script.", ADMIN_LOG)
            End If
            Exit Sub

        Case "requestbackupmain"

            If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
                Dim nothertemp As String

                f = FreeFile
                Open App.Path & "\main\Scripts\Backup.txt" For Input As #f
                nothertemp = Input$(LOF(f), f)
                Close #f
                f = FreeFile
                Open App.Path & "\main\Scripts\Main.txt" For Output As #f
                Print #f, nothertemp
                Close #f

                If SCRIPTING = 1 Then
                    Set MyScript = Nothing
                    Set clsScriptCommands = Nothing
                    Set MyScript = New clsSadScript
                    Set clsScriptCommands = New clsCommands
                    MyScript.ReadInCode App.Path & "\main\Scripts\Main.txt", "main\Scripts\Main.txt", MyScript.SControl, False
                    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
                    Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
                    Call PlayerMsg(Index, "Scripts reloaded.", White)
                End If
                Call AddLog(GetPlayerName(Index) & " used the backup script.", ADMIN_LOG)
            End If
            Exit Sub

        Case "spells"
            Call SendPlayerSpells(Index)
            Exit Sub
            
            Case "setattribute"
        Call ScriptSetAttribute(Val(Parse(1)), Val(Parse(2)), Val(Parse(3)), Val(Parse(4)), Val(Parse(5)), Val(Parse(6)), Val(Parse(7)), Val(Parse(8)), Val(Parse(9)), Val(Parse(10)))
        Exit Sub

        Case "cast"
            N = Val(Parse(1))
            Call CastSpell(Index, N)
            Exit Sub

        Case "requestlocation"

            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            Call PlayerMsg(Index, "Map: " & GetPlayerMap(Index) & ", X: " & GetPlayerX(Index) & ", Y: " & GetPlayerY(Index), Pink)
            Exit Sub

        Case "refresh"
            Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), False)
            Call PlayerMsg(Index, "Map refreshed.", White)
            Exit Sub

        Case "killpet"
            Player(Index).Pet.Alive = NO
            Player(Index).Pet.Sprite = 0
            Call TakeFromGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y)
            Packet = "PETDATA" & SEP_CHAR
            Packet = Packet & Index & SEP_CHAR
            Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
            Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
            Packet = Packet & Player(Index).Pet.x & SEP_CHAR
            Packet = Packet & Player(Index).Pet.y & SEP_CHAR
            Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
            Packet = Packet & Player(Index).Pet.Sprite & SEP_CHAR
            Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
            Packet = Packet & Player(Index).Pet.Level * 5 & SEP_CHAR
            Packet = Packet & END_CHAR
            Call SendDataToMap(GetPlayerMap(Index), Packet)
            Exit Sub

        Case "petmoveselect"
            x = Val(Parse(1))
            y = Val(Parse(2))
            Player(Index).Pet.MapToGo = GetPlayerMap(Index)
            Player(Index).Pet.Target = 0
            Player(Index).Pet.XToGo = x
            Player(Index).Pet.YToGo = y
            Player(Index).Pet.AttackTimer = GetTickCount
            For I = 1 To MAX_PLAYERS

                If IsPlaying(I) Then
                    If GetPlayerMap(I) = Player(Index).Pet.Map Then
                        If GetPlayerX(I) = x And GetPlayerY(I) = y Then
                            Player(Index).Pet.TargetType = TARGET_TYPE_PLAYER
                            Player(Index).Pet.Target = I
                            Call PlayerMsg(Index, "Your pet's target is now " & Trim$(GetPlayerName(I)) & ".", Yellow)
                            Exit Sub
                        End If
                    End If
                End If
            Next
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(Player(Index).Pet.Map, I).num > 0 Then
                    If MapNpc(Player(Index).Pet.Map, I).x = x And MapNpc(Player(Index).Pet.Map, I).y = y Then
                        Player(Index).Pet.TargetType = TARGET_TYPE_NPC
                        Player(Index).Pet.Target = I
                        Call PlayerMsg(Index, "Your pet's target is now a " & Trim$(Npc(MapNpc(Player(Index).Pet.Map, I).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If
            Next
            Call PlayerMsg(Index, "Pet is moving to (" & x & "," & y & ")", Yellow)
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

                        If GetPlayerWeaponSlot(Index) <> I And GetPlayerArmorSlot(Index) <> I And GetPlayerShieldSlot(Index) <> I And GetPlayerHelmetSlot(Index) <> I And GetPlayerLegsSlot(Index) <> I And GetPlayerBootsSlot(Index) <> I And GetPlayerGlovesSlot(Index) <> I And GetPlayerRing1Slot(Index) <> I And GetPlayerRing2Slot(Index) <> I And GetPlayerAmuletSlot(Index) <> I Then
                            Call SetPlayerInvItemNum(Index, I, 0)
                            Call PlayerMsg(Index, "You have bought a new sprite!", BrightGreen)
                            Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                            Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & END_CHAR)
                            Call SendInventory(Index)
                        End If
                    End If

                    If GetPlayerWeaponSlot(Index) <> I And GetPlayerArmorSlot(Index) <> I And GetPlayerShieldSlot(Index) <> I And GetPlayerHelmetSlot(Index) <> I And GetPlayerLegsSlot(Index) <> I And GetPlayerBootsSlot(Index) <> I And GetPlayerGlovesSlot(Index) <> I And GetPlayerRing1Slot(Index) <> I And GetPlayerRing2Slot(Index) <> I And GetPlayerAmuletSlot(Index) <> I Then
                        Exit Sub
                    End If
                End If
            Next
            Call PlayerMsg(Index, "You dont have enough to buy this sprite!", BrightRed)
            Exit Sub
            
            Case "clearowner"
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        Call PlayerMsg(Index, "Owner cleared!", BrightRed)
        Map(GetPlayerMap(Index)).Owner = ""
        Map(GetPlayerMap(Index)).Name = "Unowned House"
        Map(GetPlayerMap(Index)).Revision = Map(GetPlayerMap(Index)).Revision + 1
        Call SaveMap(GetPlayerMap(Index))
        Call SendDataToAll("maphouseupdate" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & (Map(GetPlayerMap(Index)).Owner) & SEP_CHAR & (Map(GetPlayerMap(Index)).Name) & SEP_CHAR & END_CHAR)
        Exit Sub
        
        Case "buyhouse"
        Call PlayerBuyHouse(Index)
        Exit Sub
        
        Case "stopkillquest"
      'Call StopKillQuest(Index)
            Exit Sub
        
        Case "sellhouse"
        Call PlayerBuyHouse(Index)
        Exit Sub

        Case "checkcommands"
            s = Parse(1)

            If SCRIPTING = 1 Then
                PutVar App.Path & "\main\Scripts\Command.ini", "TEMP", "Text" & Index, Trim$(s)
                MyScript.ExecuteStatement "main\Scripts\Main.txt", "Commands " & Index
            Else
                Call PlayerMsg(Index, "Thats not a valid command!", 12)
            End If
            Exit Sub

        Case "prompt"

            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "main\Scripts\Main.txt", "PlayerPrompt " & Index & "," & Val(Parse(1)) & "," & Val(Parse(2))
            End If
            Exit Sub

        Case "requesteditarrow"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            Call SendDataTo(Index, "ARROWEDITOR" & SEP_CHAR & END_CHAR)
            Exit Sub

        Case "editarrow"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_ARROWS Then
                Call HackingAttempt(Index, "Invalid arrow Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(Index) & " editing arrow #" & N & ".", ADMIN_LOG)
            Call SendEditArrowTo(Index, N)
            Exit Sub

        Case "savearrow"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_ITEMS Then
                Call HackingAttempt(Index, "Invalid arrow Index")
                Exit Sub
            End If
            Arrows(N).Name = Parse(2)
            Arrows(N).Pic = Val(Parse(3))
            Arrows(N).Range = Val(Parse(4))
            Call SendUpdateArrowToAll(N)
            Call SaveArrow(N)
            Call AddLog(GetPlayerName(Index) & " saved arrow #" & N & ".", ADMIN_LOG)
            Exit Sub
            
        Case "gofishing"
    Dim Tool As String
    Tool = GetPlayerWeaponSlot(Index)
    If Tool = 0 Then
    Call PlayerMsg(Index, "You don't have a " & Parse(4) & " equipped", 15)
    ElseIf GetPlayerInvItemNum(Index, Tool) = Val(Parse(1)) Then
       Call GoFishing(Index, Val(Parse(3)), RandomNo(MAX_LEVEL + RandomNo(200)), Parse(2))
   Else
      Call PlayerMsg(Index, "You don't have a " & Parse(4) & " equipped", 15)
    End If
    Exit Sub
    Case "gomining"
    Tool = GetPlayerWeaponSlot(Index)
    If Tool = 0 Then
    Call PlayerMsg(Index, "You don't have a " & Parse(4) & " equipped", 15)
    ElseIf GetPlayerInvItemNum(Index, Tool) = Val(Parse(1)) Then
       Call GoMining(Index, Val(Parse(3)), RandomNo(MAX_LEVEL + RandomNo(200)), Parse(2))
   Else
      Call PlayerMsg(Index, "You don't have a " & Parse(4) & " equipped", 15)
    End If
    Exit Sub
    Case "goljacking"
    Tool = GetPlayerWeaponSlot(Index)
    If Tool = 0 Then
    Call PlayerMsg(Index, "You don't have a " & Parse(4) & " equipped", 15)
    ElseIf GetPlayerInvItemNum(Index, Tool) = Val(Parse(1)) Then
       Call GoLJacking(Index, Val(Parse(3)), RandomNo(MAX_LEVEL + RandomNo(200)), Parse(2))
   Else
      Call PlayerMsg(Index, "You don't have a " & Parse(4) & " equipped", 15)
    End If
    Exit Sub

        Case "checkarrows"
            N = Arrows(Val(Parse(1))).Pic
            Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & N & SEP_CHAR & END_CHAR)
            Exit Sub

        Case "speechscript"

            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "main\Scripts\Main.txt", "ScriptedTile " & Index & "," & Parse(1)
            End If
            Exit Sub

        Case "requesteditspeech"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            Call SendDataTo(Index, "SPEECHEDITOR" & SEP_CHAR & END_CHAR)
            Exit Sub

        Case "editspeech"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_SPEECH Then
                Call HackingAttempt(Index, "Invalid Speech Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(Index) & " editing speech #" & N & ".", ADMIN_LOG)
            Call SendEditSpeechTo(Index, N)
            Exit Sub

        Case "savespeech"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_SPEECH Then
                Call HackingAttempt(Index, "Invalid Speech Index")
                Exit Sub
            End If
            Speech(N).Name = Parse(2)
            Dim p As Long

            p = 3
            For I = 0 To MAX_SPEECH_OPTIONS
                Speech(N).num(I).Exit = Val(Parse(p))
                Speech(N).num(I).text = Parse(p + 1)
                Speech(N).num(I).SaidBy = Val(Parse(p + 2))
                Speech(N).num(I).Respond = Val(Parse(p + 3))
                Speech(N).num(I).Script = Val(Parse(p + 4))
                p = p + 5
                For o = 1 To 3
                    Speech(N).num(I).Responces(o).Exit = Val(Parse(p))
                    Speech(N).num(I).Responces(o).GoTo = Val(Parse(p + 1))
                    Speech(N).num(I).Responces(o).text = Parse(p + 2)
                    p = p + 3
                Next
            Next
            Call SaveSpeech(N)
            Call SendSpeechToAll(N)
            Call AddLog(GetPlayerName(Index) & " saved speech #" & N & ".", ADMIN_LOG)
            Exit Sub

        Case "needspeech"
            Call SendSpeechTo(Index, Val(Parse(1)))
            Exit Sub

        Case "requesteditemoticon"

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            Call SendDataTo(Index, "EMOTICONEDITOR" & SEP_CHAR & END_CHAR)
            Exit Sub
            
         Case "requesteditelement"
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "ELEMENTEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub

        Case "editemoticon"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_EMOTICONS Then
                Call HackingAttempt(Index, "Invalid Emoticon Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(Index) & " editing emoticon #" & N & ".", ADMIN_LOG)
            Call SendEditEmoticonTo(Index, N)
            Exit Sub
        
        Case "editelement"
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If

        N = Val(Parse(1))
        
        If N < 0 Or N > MAX_ELEMENTS Then
            Call HackingAttempt(Index, "Invalid Element Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing element #" & N & ".", ADMIN_LOG)
        Call SendEditElementTo(Index, N)
        Exit Sub

        Case "saveemoticon"

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_EMOTICONS Then
                Call HackingAttempt(Index, "Invalid Emoticon Index")
                Exit Sub
            End If
            Emoticons(N).Type = Val(Parse(2))
            Emoticons(N).Command = Parse(3)
            Emoticons(N).Pic = Val(Parse(4))
            Emoticons(N).sound = Parse(5)
            Call SendUpdateEmoticonToAll(N)
            Call SaveEmoticon(N)
            Call AddLog(GetPlayerName(Index) & " saved emoticon #" & N & ".", ADMIN_LOG)
            Exit Sub
            
        Case "saveelement"
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        N = Val(Parse(1))
        If N < 0 Or N > MAX_ELEMENTS Then
            Call HackingAttempt(Index, "Invalid Element Index")
            Exit Sub
        End If

        Element(N).Name = Parse(2)
        Element(N).Strong = Val(Parse(3))
        Element(N).Weak = Val(Parse(4))
        
        Call SendUpdateElementToAll(N)
        Call SaveElement(N)
        Call AddLog(GetPlayerName(Index) & " saved element #" & N & ".", ADMIN_LOG)
        Exit Sub

        Case "checkemoticons"
            Call SendDataToMap(GetPlayerMap(Index), "checkemoticons" & SEP_CHAR & Index & SEP_CHAR & Emoticons(Val(Parse(1))).Type & SEP_CHAR & Emoticons(Val(Parse(1))).Pic & SEP_CHAR & Emoticons(Val(Parse(1))).sound & SEP_CHAR & END_CHAR)
            Exit Sub
            
        Case "sethands"
        Player(Index).Char(Player(Index).CharNum).Hands = Val(Parse(1))
        Exit Sub

        Case "mapreport"
            Packs = "mapreport" & SEP_CHAR
            For I = 1 To MAX_MAPS
                Packs = Packs & Map(I).Name & SEP_CHAR
            Next
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

        Case "warptome"
            N = FindPlayer(Parse(1))

            If N > 0 Then
                Call PlayerWarp(N, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Else
                Call PlayerMsg(Index, "Player not online!", BrightRed)
            End If
            Exit Sub

        Case "warpplayer"

            If Val(Parse(1)) > MAX_MAPS Or Val(Parse(1)) < 1 Then
                If FindPlayer(Trim$(Parse(1))) <> 0 Then
                    Call PlayerWarp(Index, GetPlayerMap(FindPlayer(Trim$(Parse(1)))), GetPlayerX(FindPlayer(Trim$(Parse(1)))), GetPlayerY(FindPlayer(Trim$(Parse(1)))))

                    If Player(Index).Pet.Alive = YES Then
                        Player(Index).Pet.Map = GetPlayerMap(Index)
                        Player(Index).Pet.x = GetPlayerX(Index)
                        Player(Index).Pet.y = GetPlayerY(Index)
                    End If
                Else
                    Call PlayerMsg(Index, "'" & Parse(1) & "' is not a valid map number or an online player's name!", BrightRed)
                    Exit Sub
                End If
            Else
                Call PlayerWarp(Index, Val(Parse(1)), GetPlayerX(Index), GetPlayerY(Index))

                If Player(Index).Pet.Alive = YES Then
                    Player(Index).Pet.Map = GetPlayerMap(Index)
                    Player(Index).Pet.x = GetPlayerX(Index)
                    Player(Index).Pet.y = GetPlayerY(Index)
                End If
            End If
            Exit Sub

        Case "arrowhit"
            N = Val(Parse(1))
            z = Val(Parse(2))
            x = Val(Parse(3))
            y = Val(Parse(4))

            If N = TARGET_TYPE_PLAYER Then

                ' Make sure we dont try to attack ourselves
                If z <> Index Then

                    ' Can we attack the player?
                    If CanAttackPlayerWithArrow(Index, z) Then
                        If Not CanPlayerBlockHit(z) Then

                            ' Get the damage we can do
                            If Not CanPlayerCriticalHit(Index) Then
                                Damage = GetPlayerDamage(Index) - GetPlayerProtection(z) + (Rnd * 5) - 2
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & SEP_CHAR & END_CHAR)
                            Else
                                N = GetPlayerDamage(Index)
                                Damage = N + Int(Rnd * Int(N / 2)) + 1 - GetPlayerProtection(z) + (Rnd * 5) - 2
                                Call BattleMsg(Index, "You feel a surge of energy upon shooting!", BrightCyan, 0)
                                Call BattleMsg(z, GetPlayerName(Index) & " shoots with amazing accuracy!", BrightCyan, 1)

                                'Call PlayerMsg(index, "You feel a surge of energy upon shooting!", BrightCyan)
                                'Call PlayerMsg(z, GetPlayerName(index) & " shoots with amazing accuracy!", BrightCyan)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Blow0" & SEP_CHAR & END_CHAR)
                            End If

                            If Damage > 0 Then
                                Call AttackPlayer(Index, z, Damage)
                            Else
                                Call BattleMsg(Index, "Your attack does nothing.", BrightRed, 0)
                                Call BattleMsg(z, GetPlayerName(z) & "'s attack did nothing.", BrightRed, 1)

                                'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                            End If
                        Else
                            Call BattleMsg(Index, GetPlayerName(z) & " blocked your hit!", BrightCyan, 0)
                            Call BattleMsg(z, "You blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan, 1)

                            'Call PlayerMsg(index, GetPlayerName(z) & "'s " & Trim$(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked your hit!", BrightCyan)
                            'Call PlayerMsg(z, "Your " & Trim$(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                        Exit Sub
                    End If
                End If
            ElseIf N = TARGET_TYPE_NPC Then

                ' Can we attack the npc?
                If CanAttackNpcWithArrow(Index, z) Then

                    ' Get the damage we can do
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), z).num).DEF / 2) + (Rnd * 5) - 2
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & SEP_CHAR & END_CHAR)
                    Else
                        N = GetPlayerDamage(Index)
                        Damage = N + Int(Rnd * Int(N / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), z).num).DEF / 2) + (Rnd * 5) - 2
                        Call BattleMsg(Index, "You feel a surge of energy upon shooting!", BrightCyan, 0)

                        'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "Blow0" & SEP_CHAR & END_CHAR)
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
    
    Select Case LCase$(Parse(0))
        Case "bankdeposit"
            x = GetPlayerInvItemNum(Index, Val(Parse(1)))
            I = FindOpenBankSlot(Index, x)
            If I = 0 Then
                Call SendDataTo(Index, "bankmsg" & SEP_CHAR & "Bank full!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
       
            If Val(Parse(2)) > GetPlayerInvItemValue(Index, Val(Parse(1))) Then
                Call SendDataTo(Index, "bankmsg" & SEP_CHAR & "You cant deposit more than you have!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
       
            If GetPlayerWeaponSlot(Index) = Val(Parse(1)) Or GetPlayerArmorSlot(Index) = Val(Parse(1)) Or GetPlayerShieldSlot(Index) = Val(Parse(1)) Or GetPlayerHelmetSlot(Index) = Val(Parse(1)) Or GetPlayerLegsSlot(Index) = Val(Parse(1)) Or GetPlayerBootsSlot(Index) = Val(Parse(1)) Or GetPlayerGlovesSlot(Index) = Val(Parse(1)) Or GetPlayerRing1Slot(Index) = Val(Parse(1)) Or GetPlayerRing2Slot(Index) = Val(Parse(1)) Or GetPlayerAmuletSlot(Index) = Val(Parse(1)) Then
                Call SendDataTo(Index, "bankmsg" & SEP_CHAR & "You cant deposit worn equipment!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
       
            If Item(x).Type = ITEM_TYPE_CURRENCY Or Item(x).Stackable = 1 Then
                If Val(Parse(2)) <= 0 Then
                    Call SendDataTo(Index, "bankmsg" & SEP_CHAR & "You must deposit more than 0!" & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
            End If
       
            Call TakeItem(Index, x, Val(Parse(2)))
            Call GiveBankItem(Index, x, Val(Parse(2)), I)
       
            Call SendBank(Index)
            Exit Sub
   
        Case "bankwithdraw"
            I = GetPlayerBankItemNum(Index, Val(Parse(1)))
            TempVal = Val(Parse(2))
            x = FindOpenInvSlot(Index, I)
            If x = 0 Then
                Call SendDataTo(Index, "bankmsg" & SEP_CHAR & "Inventory full!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
       
            If Val(Parse(2)) > GetPlayerBankItemValue(Index, Val(Parse(1))) Then
                Call SendDataTo(Index, "bankmsg" & SEP_CHAR & "You cant withdraw more than you have!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
               
            If Item(I).Type = ITEM_TYPE_CURRENCY Then
                If Val(Parse(2)) <= 0 Then
                    Call SendDataTo(Index, "bankmsg" & SEP_CHAR & "You must withdraw more than 0!" & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If

                If Trim(LCase(Item(GetPlayerInvItemNum(Index, x)).Name)) <> "gold" Then
                    If GetPlayerInvItemValue(Index, x) + Val(Parse(2)) > 100 Then
                        TempVal = 100 - GetPlayerInvItemValue(Index, x)
                    End If
                End If
            End If
               
            Call GiveItem(Index, I, TempVal)
            Call TakeBankItem(Index, I, TempVal)
       
            Call SendBank(Index)
            Exit Sub
    End Select
    
    Call HackingAttempt(Index, "Invalid packet. (" & Parse(0) & ")")
hell:
f = FreeFile
Open App.Path & "\errors.txt" For Append As #f
   Print #f, ("RTE:" & Err & " " & Time & " " & Date & " " & GetPlayerIP(Index) & " " & Data)
Close #f
Exit Sub
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
            Packet = Mid$(Player(Index).Buffer, 1, Start - 1)
            Player(Index).Buffer = Mid$(Player(Index).Buffer, Start + 1, Len(Player(Index).Buffer))
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
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If

    Loop
    Close #f
End Function

Function IsConnected(ByVal Index As Long) As Boolean

    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If

End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean

    If IsConnected(Index) And Trim$(Player(Index).Login) <> "" Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim I As Long

    IsMultiAccounts = False
    For I = 1 To MAX_PLAYERS

        If IsConnected(I) And LCase$(Trim$(Player(I).Login)) = LCase$(Trim$(Login)) Then
            IsMultiAccounts = True
            Exit Function
        End If
    Next
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

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal Index As Long)
Dim Packet As String

    Packet = "MAPMSG2" & SEP_CHAR & Msg & SEP_CHAR & Index & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub PlainMsg(ByVal Index As Long, ByVal Msg As String, ByVal num As Long)
Dim Packet As String

    Packet = "PLAINMSG" & SEP_CHAR & Msg & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendArrows(ByVal Index As Long)
Dim I As Long

    For I = 1 To MAX_ARROWS
        Call SendUpdateArrowTo(Index, I)
    Next
End Sub

Sub SendChars(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "ALLCHARS" & SEP_CHAR
    For I = 1 To MAX_CHARS
        Packet = Packet & Trim(Player(Index).Char(I).Name) & SEP_CHAR & Trim(Class(Player(Index).Char(I).Class).Name) & SEP_CHAR & Player(Index).Char(I).Level & SEP_CHAR & Player(Index).Char(I).Sprite & SEP_CHAR & Player(Index).Char(I).HelmetLogin & SEP_CHAR & Player(Index).Char(I).LegsLogin & SEP_CHAR & Player(Index).Char(I).ArmorLogin & SEP_CHAR & Player(Index).Char(I).ShieldLogin & SEP_CHAR & Player(Index).Char(I).WeaponLogin & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendClasses(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For I = 1 To Max_Classes
        Packet = Packet & GetClassName(I) & SEP_CHAR & GetClassMaxHP(I) & SEP_CHAR & GetClassMaxMP(I) & SEP_CHAR & GetClassMaxSP(I) & SEP_CHAR & Class(I).STR & SEP_CHAR & Class(I).DEF & SEP_CHAR & Class(I).Speed & SEP_CHAR & Class(I).Magi & SEP_CHAR & Class(I).Locked & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)

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
    Next
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
Dim I As Long

    For I = 1 To MAX_PLAYERS

        If IsPlaying(I) And I <> Index Then
            Call SendDataTo(I, Data)
        End If
    Next
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
Dim I As Long

    For I = 1 To MAX_PLAYERS

        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
Dim I As Long

    For I = 1 To MAX_PLAYERS

        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum And I <> Index Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next
End Sub

Sub SendEditArrowTo(ByVal Index As Long, ByVal EmoNum As Long)
Dim Packet As String

    Packet = "EDITArrow" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Arrows(EmoNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditEmoticonTo(ByVal Index As Long, ByVal EmoNum As Long)
Dim Packet As String

    Packet = "EDITEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Emoticons(EmoNum).Type & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & Emoticons(EmoNum).sound & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateElementToAll(ByVal ElementNum As Long)
Dim Packet As String

    Packet = "UPDATEELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & SEP_CHAR & END_CHAR
 Call SendDataToAll(Packet)
End Sub

Sub SendUpdateElementTo(ByVal Index As Long, ByVal ElementNum As Long)
Dim Packet As String

    Packet = "UPDATEELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & SEP_CHAR & END_CHAR
  Call SendDataTo(Index, Packet)
End Sub

Sub SendEditElementTo(ByVal Index As Long, ByVal ElementNum As Long)
Dim Packet As String

    Packet = "EDITELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & SEP_CHAR & END_CHAR
 Call SendDataTo(Index, Packet)
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound & SEP_CHAR & Item(ItemNum).LevelReq & SEP_CHAR & Item(ItemNum).Element & SEP_CHAR & Item(ItemNum).StamRemove & SEP_CHAR & Item(ItemNum).Rarity & SEP_CHAR & Item(ItemNum).BowsReq & SEP_CHAR & Item(ItemNum).LargeBladesReq & SEP_CHAR & Item(ItemNum).SmallBladesReq & SEP_CHAR & Item(ItemNum).BluntWeaponsReq & SEP_CHAR & Item(ItemNum).PoleArmsReq & SEP_CHAR & Item(ItemNum).AxesReq & SEP_CHAR & Item(ItemNum).ThrownReq & SEP_CHAR & Item(ItemNum).XbowsReq & SEP_CHAR & Item(ItemNum).LBA & SEP_CHAR & Item(ItemNum).SBA & SEP_CHAR & Item(ItemNum).BWA
    Packet = Packet & SEP_CHAR & Item(ItemNum).PAA & SEP_CHAR & Item(ItemNum).AA & SEP_CHAR & Item(ItemNum).TWA & SEP_CHAR & Item(ItemNum).XBA & SEP_CHAR & Item(ItemNum).BA & SEP_CHAR & Item(ItemNum).Poison & SEP_CHAR & Item(ItemNum).Disease & SEP_CHAR & Item(ItemNum).AilmentDamage & SEP_CHAR & Item(ItemNum).AilmentMS & SEP_CHAR & Item(ItemNum).AilmentInterval
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal npcnum As Long)
Dim Packet As String
Dim I As Long

    Packet = "EDITNPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(Npc(npcnum).Name) & SEP_CHAR & Trim$(Npc(npcnum).AttackSay) & SEP_CHAR & Npc(npcnum).Sprite & SEP_CHAR & Npc(npcnum).SpawnSecs & SEP_CHAR & Npc(npcnum).Behavior & SEP_CHAR & Npc(npcnum).Range & SEP_CHAR & Npc(npcnum).STR & SEP_CHAR & Npc(npcnum).DEF & SEP_CHAR & Npc(npcnum).Speed & SEP_CHAR & Npc(npcnum).Magi & SEP_CHAR & Npc(npcnum).Big & SEP_CHAR & Npc(npcnum).MaxHp & SEP_CHAR & Npc(npcnum).Exp & SEP_CHAR & Npc(npcnum).SpawnTime & SEP_CHAR & Npc(npcnum).Speech & SEP_CHAR & Npc(npcnum).Element & SEP_CHAR & Npc(npcnum).Poison & SEP_CHAR & Npc(npcnum).AP & SEP_CHAR & Npc(npcnum).Disease & SEP_CHAR & Npc(npcnum).Quest & SEP_CHAR & Npc(npcnum).NpcDIR & SEP_CHAR & Npc(npcnum).AilmentDamage & SEP_CHAR & Npc(npcnum).AilmentInterval & SEP_CHAR & Npc(npcnum).AilmentMS & SEP_CHAR & Npc(npcnum).Spell & SEP_CHAR
    For I = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(npcnum).ItemNPC(I).Chance
        Packet = Packet & SEP_CHAR & Npc(npcnum).ItemNPC(I).ItemNum
        Packet = Packet & SEP_CHAR & Npc(npcnum).ItemNPC(I).ItemValue & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim I As Long, z As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For I = 1 To 6
        For z = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(I).Value(z).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(z).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(z).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(z).GetValue & SEP_CHAR
        Next
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSpeechTo(ByVal Index As Long, ByVal SpcNum As Long)
Dim Packet As String
Dim I, o As Long

    Packet = "EDITSPEECH" & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR
    For I = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(I).Exit & SEP_CHAR & Speech(SpcNum).num(I).text & SEP_CHAR & Speech(SpcNum).num(I).SaidBy & SEP_CHAR & Speech(SpcNum).num(I).Respond & SEP_CHAR & Speech(SpcNum).num(I).Script & SEP_CHAR
        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(I).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(I).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(I).Responces(o).text & SEP_CHAR
        Next
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal spellnum As Long)
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & spellnum & SEP_CHAR & Trim$(Spell(spellnum).Name) & SEP_CHAR & Spell(spellnum).ClassReq & SEP_CHAR & Spell(spellnum).LevelReq & SEP_CHAR & Spell(spellnum).Type & SEP_CHAR & Spell(spellnum).Data1 & SEP_CHAR & Spell(spellnum).Data2 & SEP_CHAR & Spell(spellnum).Data3 & SEP_CHAR & Spell(spellnum).MPCost & SEP_CHAR & Spell(spellnum).sound & SEP_CHAR & Spell(spellnum).Range & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & Spell(spellnum).AE & SEP_CHAR & Spell(spellnum).Pic & SEP_CHAR & Spell(spellnum).Element & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEmoticons(ByVal Index As Long)
Dim I As Long

    For I = 0 To MAX_EMOTICONS

        If Trim$(Emoticons(I).Command) <> "" Then
            Call SendUpdateEmoticonTo(Index, I)
        End If
    Next
End Sub

Sub SendFriendListTo(ByVal Index As Long)
Dim Packet As String
Dim N As Long

    Packet = "FRIENDLIST" & SEP_CHAR
    For N = 1 To MAX_FRIENDS

        If FindPlayer(Player(Index).Char(Player(Index).CharNum).Friends(N)) And Player(Index).Char(Player(Index).CharNum).Friends(N) <> "" Then
            Packet = Packet & Player(Index).Char(Player(Index).CharNum).Friends(N) & SEP_CHAR
        End If
    Next
    Packet = Packet & NEXT_CHAR & SEP_CHAR
    For N = 1 To MAX_FRIENDS
        Packet = Packet & Player(Index).Char(Player(Index).CharNum).Friends(N) & SEP_CHAR
    Next
    Packet = Packet & NEXT_CHAR & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendFriendListToNeeded(ByVal Name As String)
Dim I, o As Long

    For I = I To MAX_PLAYERS

        If IsPlaying(I) Then
            For o = 1 To MAX_FRIENDS

                If Trim$(Player(I).Char(Player(I).CharNum).Friends(I)) = Name Then
                    Call SendFriendListTo(I)
                End If
            Next
        End If
    Next
End Sub

Sub SendHP(ByVal Index As Long)
Dim Packet As String

    Packet = "playerhp" & SEP_CHAR & Index & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
    Packet = "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendFP(ByVal Index As Long)
Dim Packet As String

    Packet = "playerfp" & SEP_CHAR & Index & SEP_CHAR & GetPlayerMaxFP(Index) & SEP_CHAR & GetPlayerFP(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendInfo(ByVal Index As Long)
Dim Packet As String

    Packet = "INFO" & SEP_CHAR & TotalOnlinePlayers & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendBank(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "PLAYERBANK" & SEP_CHAR
    For I = 1 To MAX_BANK
        Packet = Packet & GetPlayerBankItemNum(Index, I) & SEP_CHAR & GetPlayerBankItemValue(Index, I) & SEP_CHAR & GetPlayerBankItemDur(Index, I) & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
   
    Call SendDataTo(Index, Packet)
End Sub

Sub SendBankUpdate(ByVal Index As Long, ByVal BankSlot As Long)
Dim Packet As String
   
    Packet = "PLAYERBANKUPDATE" & SEP_CHAR & BankSlot & SEP_CHAR & GetPlayerBankItemNum(Index, BankSlot) & SEP_CHAR & GetPlayerBankItemValue(Index, BankSlot) & SEP_CHAR & GetPlayerBankItemDur(Index, BankSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendInventory(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "PLAYERINV" & SEP_CHAR & Index & SEP_CHAR
    For I = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, I) & SEP_CHAR & GetPlayerInvItemValue(Index, I) & SEP_CHAR & GetPlayerInvItemDur(Index, I) & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim Packet As String

    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & Index & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendItems(ByVal Index As Long)
Dim I As Long

    For I = 1 To MAX_ITEMS

        If Trim$(Item(I).Name) <> "" Then
            Call SendUpdateItemTo(Index, I)
        End If
    Next
End Sub

Sub SendElements(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    For I = 0 To MAX_ELEMENTS
            Call SendUpdateElementTo(Index, I)
    Next I
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
            Packet = Packet & GetPlayerAlignment(I) & SEP_CHAR
            Packet = Packet & END_CHAR
            Call SendDataTo(Index, Packet)

            If Player(I).Pet.Alive = YES Then
                Packet = "PETDATA" & SEP_CHAR
                Packet = Packet & I & SEP_CHAR
                Packet = Packet & Player(I).Pet.Alive & SEP_CHAR
                Packet = Packet & Player(I).Pet.Map & SEP_CHAR
                Packet = Packet & Player(I).Pet.x & SEP_CHAR
                Packet = Packet & Player(I).Pet.y & SEP_CHAR
                Packet = Packet & Player(I).Pet.Dir & SEP_CHAR
                Packet = Packet & Player(I).Pet.Sprite & SEP_CHAR
                Packet = Packet & Player(I).Pet.HP & SEP_CHAR
                Packet = Packet & Player(I).Pet.Level * 5 & SEP_CHAR
                Packet = Packet & END_CHAR
                Call SendDataTo(Index, Packet)
            End If
        End If
    Next

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
    Packet = Packet & GetPlayerAlignment(Index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)

    If Player(Index).Pet.Alive = YES Then
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & Index & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Index).Pet.x & SEP_CHAR
        Packet = Packet & Player(Index).Pet.y & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataToMap(GetPlayerMap(Index), Packet)
    End If
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
    Packet = Packet & GetPlayerAlignment(Index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMapBut(Index, MapNum, Packet)

    If Player(Index).Pet.Alive = YES Then
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & Index & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Index).Pet.x & SEP_CHAR
        Packet = Packet & Player(Index).Pet.y & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataToMapBut(Index, MapNum, Packet)
    End If
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
    Packet = Packet & END_CHAR
    Call SendDataToAllBut(Index, Packet)
    Packet = "PETDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
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

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim x As Long
Dim y As Long
Dim I As Long
Dim o As Long
Dim p1 As String, p2 As String

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX

            With Map(MapNum).Tile(x, y)
                I = 0
                o = 0

                If .Ground <> 0 Then I = 0
                If .GroundSet <> -1 Then I = 1
                If .Mask <> 0 Then I = 2
                If .MaskSet <> -1 Then I = 3
                If .Anim <> 0 Then I = 4
                If .AnimSet <> -1 Then I = 5
                If .Fringe <> 0 Then I = 6
                If .FringeSet <> -1 Then I = 7
                If .Type <> 0 Then I = 8
                If .Data1 <> 0 Then I = 9
                If .Data2 <> 0 Then I = 10
                If .Data3 <> 0 Then I = 11
                If .String1 <> "" Then I = 12
                If .String2 <> "" Then I = 13
                If .String3 <> "" Then I = 14
                If .Mask2 <> 0 Then I = 15
                If .Mask2Set <> -1 Then I = 16
                If .M2Anim <> 0 Then I = 17
                If .M2AnimSet <> -1 Then I = 18
                If .FAnim <> 0 Then I = 19
                If .FAnimSet <> -1 Then I = 20
                If .Fringe2 <> 0 Then I = 21
                If .Fringe2Set <> -1 Then I = 22
                If .Light <> 0 Then I = 23
                If .F2Anim <> 0 Then I = 24
                If .F2AnimSet <> -1 Then I = 25
                Packet = Packet & .Ground & SEP_CHAR

                If o < I Then
                    o = o + 1
                    Packet = Packet & .GroundSet & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Mask & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .MaskSet & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Anim & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .AnimSet & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Fringe & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .FringeSet & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Type & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Data1 & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Data2 & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Data3 & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .String1 & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .String2 & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .String3 & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Mask2 & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Mask2Set & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .M2Anim & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .M2AnimSet & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .FAnim & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .FAnimSet & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Fringe2 & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Fringe2Set & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .Light & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .F2Anim & SEP_CHAR
                End If

                If o < I Then
                    o = o + 1
                    Packet = Packet & .F2AnimSet & SEP_CHAR
                End If
                Packet = Packet & NEXT_CHAR & SEP_CHAR
            End With

        Next
    Next
    For x = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(MapNum).Npc(x) & SEP_CHAR
        Packet = Packet & Map(MapNum).NpcSpawn(x).Used & SEP_CHAR & Map(MapNum).NpcSpawn(x).x & SEP_CHAR & Map(MapNum).NpcSpawn(x).y & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    x = Int(Len(Packet) / 2)
    p1 = Mid$(Packet, 1, x)
    p2 = Mid$(Packet, x + 1, Len(Packet) - x)
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim I As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_ITEMS

        If MapNum > 0 Then
            Packet = Packet & MapItem(MapNum, I).num & SEP_CHAR & MapItem(MapNum, I).Value & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & MapItem(MapNum, I).x & SEP_CHAR & MapItem(MapNum, I).y & SEP_CHAR
        End If
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim I As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, I).num & SEP_CHAR & MapItem(MapNum, I).Value & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & MapItem(MapNum, I).x & SEP_CHAR & MapItem(MapNum, I).y & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim I As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_NPCS

        If MapNum > 0 Then
            Packet = Packet & MapNpc(MapNum, I).num & SEP_CHAR & MapNpc(MapNum, I).x & SEP_CHAR & MapNpc(MapNum, I).y & SEP_CHAR & MapNpc(MapNum, I).Dir & SEP_CHAR
        End If
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
Dim I As Long
Dim Packet As String

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR
    For I = 1 To Max_Classes
        Packet = Packet & GetClassName(I) & SEP_CHAR & GetClassMaxHP(I) & SEP_CHAR & GetClassMaxMP(I) & SEP_CHAR & GetClassMaxSP(I) & SEP_CHAR & Class(I).STR & SEP_CHAR & Class(I).DEF & SEP_CHAR & Class(I).Speed & SEP_CHAR & Class(I).Magi & SEP_CHAR & Class(I).MaleSprite & SEP_CHAR & Class(I).FemaleSprite & SEP_CHAR & Class(I).Locked & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNpcs(ByVal Index As Long)
Dim I As Long

    For I = 1 To MAX_NPCS

        If Trim$(Npc(I).Name) <> "" Then
            Call SendUpdateNpcTo(Index, I)
        End If
    Next
End Sub

Sub SendOnlineList()
Dim Packet As String
Dim I As Long
Dim N As Long

    Packet = ""
    N = 0
    For I = 1 To MAX_PLAYERS

        If IsPlaying(I) Then
            Packet = Packet & SEP_CHAR & GetPlayerName(I) & SEP_CHAR
            N = N + 1
        End If
    Next
    Packet = "ONLINELIST" & SEP_CHAR & N & Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendPlayerData(ByVal Index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on the map
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
    Packet = Packet & GetPlayerAlignment(Index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)

    If Player(Index).Pet.Alive = YES Then
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & Index & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Index).Pet.x & SEP_CHAR
        Packet = Packet & Player(Index).Pet.y & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataToMap(GetPlayerMap(Index), Packet)
    End If
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "SPELLS" & SEP_CHAR
    For I = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(Index, I) & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerXY(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendShops(ByVal Index As Long)
Dim I As Long

    For I = 1 To MAX_SHOPS

        If Trim$(Shop(I).Name) <> "" Then
            Call SendUpdateShopTo(Index, I)
        End If
    Next
End Sub

Sub SendSP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendSpeech(ByVal Index As Long)
Dim I As Long

    For I = 1 To MAX_SPEECH

        If Trim$(Speech(I).Name) <> "" Then
            Call SendSpeechTo(Index, I)
        End If
    Next
End Sub

Sub SendSpeechTo(ByVal Index As Long, ByVal SpcNum As Long)
Dim Packet As String
Dim I, o As Long

    Packet = "SPEECH" & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR
    For I = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(I).Exit & SEP_CHAR & Speech(SpcNum).num(I).text & SEP_CHAR & Speech(SpcNum).num(I).SaidBy & SEP_CHAR & Speech(SpcNum).num(I).Respond & SEP_CHAR & Speech(SpcNum).num(I).Script & SEP_CHAR
        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(I).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(I).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(I).Responces(o).text & SEP_CHAR
        Next
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendSpeechToAll(ByVal SpcNum As Long)
Dim Packet As String
Dim I, o As Long

    Packet = "SPEECH" & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR
    For I = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(I).Exit & SEP_CHAR & Speech(SpcNum).num(I).text & SEP_CHAR & Speech(SpcNum).num(I).SaidBy & SEP_CHAR & Speech(SpcNum).num(I).Respond & SEP_CHAR & Speech(SpcNum).num(I).Script & SEP_CHAR
        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(I).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(I).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(I).Responces(o).text & SEP_CHAR
        Next
    Next
    Packet = Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendSpells(ByVal Index As Long)
Dim I As Long

    For I = 1 To MAX_SPELLS

        If Trim$(Spell(I).Name) <> "" Then
            Call SendUpdateSpellTo(Index, I)
        End If
    Next
End Sub

Sub SendStats(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerstr(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & SEP_CHAR & GetPlayerNextLargeBladesLevel(Index) & SEP_CHAR & GetPlayerLargeBladesExp(Index) & SEP_CHAR & GetPlayerLargeBladesLevel(Index) & SEP_CHAR & GetPlayerNextSmallBladesLevel(Index) & SEP_CHAR & GetPlayerSmallBladesExp(Index) & SEP_CHAR & GetPlayerSmallBladesLevel(Index) & SEP_CHAR & GetPlayerNextBluntWeaponsLevel(Index) & SEP_CHAR & GetPlayerBluntWeaponsExp(Index) & SEP_CHAR & GetPlayerBluntWeaponsLevel(Index) & SEP_CHAR & GetPlayerNextPolesLevel(Index) & SEP_CHAR & GetPlayerPolesExp(Index) & SEP_CHAR & GetPlayerPolesLevel(Index) & SEP_CHAR & GetPlayerNextAxesLevel(Index) & SEP_CHAR & GetPlayerAxesExp(Index) & SEP_CHAR & GetPlayerAxesLevel(Index) & SEP_CHAR
    Packet = Packet & GetPlayerNextThrownLevel(Index) & SEP_CHAR & GetPlayerThrownExp(Index) & SEP_CHAR & GetPlayerThrownLevel(Index) & SEP_CHAR & GetPlayerNextXbowsLevel(Index) & SEP_CHAR & GetPlayerXbowsExp(Index) & SEP_CHAR & GetPlayerXbowsLevel(Index) & SEP_CHAR & GetPlayerNextBowsLevel(Index) & SEP_CHAR & GetPlayerBowsExp(Index) & SEP_CHAR & GetPlayerBowsLevel(Index) & SEP_CHAR & GetPlayerNextFishLevel(Index) & SEP_CHAR & GetPlayerFishExp(Index) & SEP_CHAR & GetPlayerFishLevel(Index) & SEP_CHAR & GetPlayerNextMineLevel(Index) & SEP_CHAR & GetPlayerMineExp(Index) & SEP_CHAR & GetPlayerMineLevel(Index) & SEP_CHAR & GetPlayerNextLJackingLevel(Index) & SEP_CHAR & GetPlayerLJackingExp(Index) & SEP_CHAR & GetPlayerLJackingLevel(Index) & SEP_CHAR & GetPlayerArrowsAmount(Index) & SEP_CHAR & GetPlayerPoisoned(Index) & SEP_CHAR & GetPlayerDiseased(Index) & SEP_CHAR & END_CHAR
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
    Next
    Call SpawnAllMapNpcs
End Sub

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim I As Long, x As Long, y As Long, z As Long, XX As Long

    z = 0
    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For I = 1 To 6
        For XX = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(I).Value(XX).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(XX).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(XX).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(XX).GetValue & SEP_CHAR

            ' Item #
            x = Shop(ShopNum).TradeItem(I).Value(XX).GetItem

            If Item(x).Type = ITEM_TYPE_SPELL Then

                ' Spell class requirement
                y = Spell(Item(x).Data1).ClassReq

                If y = 0 Then
                    Call PlayerMsg(Index, Trim$(Item(x).Name) & " can be used by all classes.", Yellow)
                Else
                    Call PlayerMsg(Index, Trim$(Item(x).Name) & " can only be used by a " & GetClassName(y) & ".", Yellow)
                End If
            End If

            If x < 1 Then
                z = z + 1
            End If
        Next
    Next
    Packet = Packet & END_CHAR

    If z = (MAX_TRADES * 6) Then
        Call PlayerMsg(Index, "This shop has nothing to sell!", BrightRed)
    Else
        Call SendDataTo(Index, Packet)
    End If
End Sub

Sub SendUpdateArrowTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEArrow" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEARROW" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateEmoticonTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Emoticons(ItemNum).Type & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & Emoticons(ItemNum).sound & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Emoticons(ItemNum).Type & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & Emoticons(ItemNum).sound & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound & SEP_CHAR & Item(ItemNum).LevelReq & SEP_CHAR & Item(ItemNum).Element & SEP_CHAR & Item(ItemNum).StamRemove & SEP_CHAR & Item(ItemNum).Rarity & SEP_CHAR & Item(ItemNum).BowsReq & SEP_CHAR & Item(ItemNum).LargeBladesReq & SEP_CHAR & Item(ItemNum).SmallBladesReq & SEP_CHAR & Item(ItemNum).BluntWeaponsReq & SEP_CHAR & Item(ItemNum).PoleArmsReq & SEP_CHAR & Item(ItemNum).AxesReq & SEP_CHAR & Item(ItemNum).ThrownReq & SEP_CHAR & Item(ItemNum).XbowsReq & SEP_CHAR & Item(ItemNum).LBA & SEP_CHAR & Item(ItemNum).SBA & SEP_CHAR & Item(ItemNum).BWA
    Packet = Packet & SEP_CHAR & Item(ItemNum).PAA & SEP_CHAR & Item(ItemNum).AA & SEP_CHAR & Item(ItemNum).TWA & SEP_CHAR & Item(ItemNum).XBA & SEP_CHAR & Item(ItemNum).BA & SEP_CHAR & Item(ItemNum).Poison & SEP_CHAR & Item(ItemNum).Disease & SEP_CHAR & Item(ItemNum).AilmentDamage & SEP_CHAR & Item(ItemNum).AilmentMS & SEP_CHAR & Item(ItemNum).AilmentInterval
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound & SEP_CHAR & Item(ItemNum).LevelReq & SEP_CHAR & Item(ItemNum).Element & SEP_CHAR & Item(ItemNum).StamRemove & SEP_CHAR & Item(ItemNum).Rarity & SEP_CHAR & Item(ItemNum).BowsReq & SEP_CHAR & Item(ItemNum).LargeBladesReq & SEP_CHAR & Item(ItemNum).SmallBladesReq & SEP_CHAR & Item(ItemNum).BluntWeaponsReq & SEP_CHAR & Item(ItemNum).PoleArmsReq & SEP_CHAR & Item(ItemNum).AxesReq & SEP_CHAR & Item(ItemNum).ThrownReq & SEP_CHAR & Item(ItemNum).XbowsReq & SEP_CHAR & Item(ItemNum).LBA & SEP_CHAR & Item(ItemNum).SBA & SEP_CHAR & Item(ItemNum).BWA
    Packet = Packet & SEP_CHAR & Item(ItemNum).PAA & SEP_CHAR & Item(ItemNum).AA & SEP_CHAR & Item(ItemNum).TWA & SEP_CHAR & Item(ItemNum).XBA & SEP_CHAR & Item(ItemNum).BA & SEP_CHAR & Item(ItemNum).Poison & SEP_CHAR & Item(ItemNum).Disease & SEP_CHAR & Item(ItemNum).AilmentDamage & SEP_CHAR & Item(ItemNum).AilmentMS & SEP_CHAR & Item(ItemNum).AilmentInterval
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal npcnum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(Npc(npcnum).Name) & SEP_CHAR & Trim$(Npc(npcnum).AttackSay) & SEP_CHAR & Npc(npcnum).Sprite & SEP_CHAR & Npc(npcnum).SpawnSecs & SEP_CHAR & Npc(npcnum).Behavior & SEP_CHAR & Npc(npcnum).Range & SEP_CHAR & Npc(npcnum).STR & SEP_CHAR & Npc(npcnum).DEF & SEP_CHAR & Npc(npcnum).Speed & SEP_CHAR & Npc(npcnum).Magi & SEP_CHAR & Npc(npcnum).Big & SEP_CHAR & Npc(npcnum).MaxHp & SEP_CHAR & Npc(npcnum).Exp & SEP_CHAR & Npc(npcnum).SpawnTime & SEP_CHAR & Npc(npcnum).Speech & SEP_CHAR & Npc(npcnum).Element & SEP_CHAR & Npc(npcnum).Poison & SEP_CHAR & Npc(npcnum).AP & SEP_CHAR & Npc(npcnum).Disease & SEP_CHAR & Npc(npcnum).Quest & SEP_CHAR & Npc(npcnum).NpcDIR & SEP_CHAR & Npc(npcnum).AilmentDamage & SEP_CHAR & Npc(npcnum).AilmentInterval & SEP_CHAR & Npc(npcnum).AilmentMS & SEP_CHAR & Npc(npcnum).Spell & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal npcnum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(Npc(npcnum).Name) & SEP_CHAR & Trim$(Npc(npcnum).AttackSay) & SEP_CHAR & Npc(npcnum).Sprite & SEP_CHAR & Npc(npcnum).SpawnSecs & SEP_CHAR & Npc(npcnum).Behavior & SEP_CHAR & Npc(npcnum).Range & SEP_CHAR & Npc(npcnum).STR & SEP_CHAR & Npc(npcnum).DEF & SEP_CHAR & Npc(npcnum).Speed & SEP_CHAR & Npc(npcnum).Magi & SEP_CHAR & Npc(npcnum).Big & SEP_CHAR & Npc(npcnum).MaxHp & SEP_CHAR & Npc(npcnum).Exp & SEP_CHAR & Npc(npcnum).SpawnTime & SEP_CHAR & Npc(npcnum).Speech & SEP_CHAR & Npc(npcnum).Element & SEP_CHAR & Npc(npcnum).Poison & SEP_CHAR & Npc(npcnum).AP & SEP_CHAR & Npc(npcnum).Disease & SEP_CHAR & Npc(npcnum).Quest & SEP_CHAR & Npc(npcnum).NpcDIR & SEP_CHAR & Npc(npcnum).AilmentDamage & SEP_CHAR & Npc(npcnum).AilmentInterval & SEP_CHAR & Npc(npcnum).AilmentMS & SEP_CHAR & Npc(npcnum).Spell & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal spellnum As Long)
    Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & spellnum & SEP_CHAR & Trim$(Spell(spellnum).Name) & SEP_CHAR & Spell(spellnum).ClassReq & SEP_CHAR & Spell(spellnum).LevelReq & SEP_CHAR & Spell(spellnum).Type & SEP_CHAR & Spell(spellnum).Data1 & SEP_CHAR & Spell(spellnum).Data2 & SEP_CHAR & Spell(spellnum).Data3 & SEP_CHAR & Spell(spellnum).MPCost & SEP_CHAR & Trim$(Spell(spellnum).sound) & SEP_CHAR & Spell(spellnum).Range & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & Spell(spellnum).AE & SEP_CHAR & Spell(spellnum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateSpellToAll(ByVal spellnum As Long)
    Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & spellnum & SEP_CHAR & Trim$(Spell(spellnum).Name) & SEP_CHAR & Spell(spellnum).ClassReq & SEP_CHAR & Spell(spellnum).LevelReq & SEP_CHAR & Spell(spellnum).Type & SEP_CHAR & Spell(spellnum).Data1 & SEP_CHAR & Spell(spellnum).Data2 & SEP_CHAR & Spell(spellnum).Data3 & SEP_CHAR & Spell(spellnum).MPCost & SEP_CHAR & Trim$(Spell(spellnum).sound) & SEP_CHAR & Spell(spellnum).Range & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & Spell(spellnum).AE & SEP_CHAR & Spell(spellnum).Pic & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
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
    Next
End Sub

Sub SendWhosOnline(ByVal Index As Long)
Dim s As String
Dim N As Long, I As Long

    s = ""
    N = 0
    For I = 1 To MAX_PLAYERS

        If IsPlaying(I) And I <> Index Then
            s = s & GetPlayerName(I) & ", "
            N = N + 1
        End If
    Next

    If N = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & N & " other players online: " & s & "."
    End If
    Call PlayerMsg(Index, s, WhoColor)
End Sub

Sub SendInvSlots(ByVal Index As Long)
    Dim Packet As String

    If IsPlaying(Index) Then
        Packet = "PLAYERINVSLOTS" & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & GetPlayerLegsSlot(Index) & SEP_CHAR & GetPlayerBootsSlot(Index) & SEP_CHAR & GetPlayerGlovesSlot(Index) & SEP_CHAR & GetPlayerRing1Slot(Index) & SEP_CHAR & GetPlayerRing2Slot(Index) & SEP_CHAR & GetPlayerAmuletSlot(Index) & SEP_CHAR & END_CHAR
        Call SendDataTo(Index, Packet)
    End If

End Sub

Sub SendWornEquipment(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(Index) = GetPlayerMap(I) And I <> Index Then
                Packet = "PLAYERWORNEQ" & SEP_CHAR & I & SEP_CHAR
                If GetPlayerArmorSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerArmorSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerWeaponSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerWeaponSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerHelmetSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerHelmetSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerShieldSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerShieldSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                 If GetPlayerLegsSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerLegsSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerBootsSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerBootsSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerGlovesSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerGlovesSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerRing1Slot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerRing1Slot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerRing2Slot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerRing2Slot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerAmuletSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerAmuletSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                Packet = Packet & END_CHAR
                Call SendDataTo(Index, Packet)
                
                Packet = "PLAYERWORNEQ" & SEP_CHAR & Index & SEP_CHAR
                If GetPlayerArmorSlot(Index) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerWeaponSlot(Index) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerHelmetSlot(Index) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerShieldSlot(Index) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerLegsSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerLegsSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerBootsSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerBootsSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerGlovesSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerGlovesSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerRing1Slot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerRing1Slot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerRing2Slot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerRing2Slot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerAmuletSlot(I) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(I, GetPlayerAmuletSlot(I)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                Packet = Packet & END_CHAR
                Call SendDataTo(I, Packet)
            End If
        End If
    Next
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

Sub UpdateCaption()
    frmServer.Caption = GAME_NAME & " - Server - Powered By The Chaos Engine Source V.1.4.9 DX7 FINAL"
    frmServer.lblIP.Caption = "Ip Address: " & frmServer.Socket(0).LocalIP
    frmServer.lblPort.Caption = "Port: " & STR(frmServer.Socket(0).LocalPort)
    frmServer.TPO.Caption = "Total Players Online: " & TotalOnlinePlayers
    Exit Sub
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

Sub SendNewsTo(ByVal Index As Long)
Dim Packet As String

    Packet = "NEWS" & SEP_CHAR & ReadINI("DATA", "ServerNews", App.Path & "\News.ini") & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub QuestMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Hate As Byte)
Dim Packet As String

    Packet = "QUESTMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & Hate & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendQuest(ByVal Index As Long)
'Dim Packet As String
Dim I As Long

    For I = 1 To MAX_QUESTS
        If Trim(Quest(I).Name) <> "" Then
            Call SendUpdateQuestTo(Index, I)
        End If
    Next I
End Sub

Sub SendUpdateQuestToAll(ByVal questnum As Long)
Dim Packet As String

    Packet = "UPDATEQUEST" & SEP_CHAR & questnum & SEP_CHAR & Trim(Quest(questnum).Name) & SEP_CHAR & Trim(Quest(questnum).After) & SEP_CHAR & Trim(Quest(questnum).Before) & SEP_CHAR & Quest(questnum).ClassIsReq & SEP_CHAR & Quest(questnum).ClassReq & SEP_CHAR & Trim(Quest(questnum).During) & SEP_CHAR & Trim(Quest(questnum).End) & SEP_CHAR & Quest(questnum).ItemReq & SEP_CHAR & Quest(questnum).ItemVal & SEP_CHAR & Quest(questnum).LevelIsReq & SEP_CHAR & Quest(questnum).LevelReq & SEP_CHAR & Trim(Quest(questnum).NotHasItem) & SEP_CHAR & Quest(questnum).RewardNum & SEP_CHAR & Quest(questnum).RewardVal & SEP_CHAR & Trim(Quest(questnum).Start) & SEP_CHAR & Quest(questnum).StartItem & SEP_CHAR & Quest(questnum).StartOn & SEP_CHAR & Quest(questnum).Startval & SEP_CHAR & Quest(questnum).QuestExpReward & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateQuestTo(ByVal Index As Long, ByVal questnum As Long)
Dim Packet As String

    Packet = "UPDATEQUEST" & SEP_CHAR & questnum & SEP_CHAR & Trim(Quest(questnum).Name) & SEP_CHAR & Trim(Quest(questnum).After) & SEP_CHAR & Trim(Quest(questnum).Before) & SEP_CHAR & Quest(questnum).ClassIsReq & SEP_CHAR & Quest(questnum).ClassReq & SEP_CHAR & Trim(Quest(questnum).During) & SEP_CHAR & Trim(Quest(questnum).End) & SEP_CHAR & Quest(questnum).ItemReq & SEP_CHAR & Quest(questnum).ItemVal & SEP_CHAR & Quest(questnum).LevelIsReq & SEP_CHAR & Quest(questnum).LevelReq & SEP_CHAR & Trim(Quest(questnum).NotHasItem) & SEP_CHAR & Quest(questnum).RewardNum & SEP_CHAR & Quest(questnum).RewardVal & SEP_CHAR & Trim(Quest(questnum).Start) & SEP_CHAR & Quest(questnum).StartItem & SEP_CHAR & Quest(questnum).StartOn & SEP_CHAR & Quest(questnum).Startval & SEP_CHAR & Quest(questnum).QuestExpReward & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditQuestTo(ByVal Index As Long, ByVal questnum As Long)
Dim Packet As String

    Packet = "EDITQUEST" & SEP_CHAR & questnum & SEP_CHAR & Trim(Quest(questnum).Name) & SEP_CHAR & Trim(Quest(questnum).After) & SEP_CHAR & Trim(Quest(questnum).Before) & SEP_CHAR & Quest(questnum).ClassIsReq & SEP_CHAR & Quest(questnum).ClassReq & SEP_CHAR & Trim(Quest(questnum).During) & SEP_CHAR & Trim(Quest(questnum).End) & SEP_CHAR & Quest(questnum).ItemReq & SEP_CHAR & Quest(questnum).ItemVal & SEP_CHAR & Quest(questnum).LevelIsReq & SEP_CHAR & Quest(questnum).LevelReq & SEP_CHAR & Trim(Quest(questnum).NotHasItem) & SEP_CHAR & Quest(questnum).RewardNum & SEP_CHAR & Quest(questnum).RewardVal & SEP_CHAR & Trim(Quest(questnum).Start) & SEP_CHAR & Quest(questnum).StartItem & SEP_CHAR & Quest(questnum).StartOn & SEP_CHAR & Quest(questnum).Startval & SEP_CHAR & Quest(questnum).QuestExpReward & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub
