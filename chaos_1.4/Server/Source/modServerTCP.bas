Attribute VB_Name = "modServerTCP"

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit

'Dim ZePacket() As String ' SAFE MODE -- Uncomment for ON, comment for OFF
'Dim NumParse As Long ' SAFE MODE -- Uncomment for ON, comment for OFF
'Dim ParseIndex As Long ' SAFE MODE -- Uncomment for ON, comment for OFF

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

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, Packet)
        End If
    Next
End Sub

Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
    Call CloseSocket(index)
End Sub

Sub CloseSocket(ByVal index As Long)

    ' Make sure player was/is playing the game, and if so, save'm.
    If index > 0 Then
        Call LeftGame(index)
        Call TextAdd(frmServer.txtText(0), "Connection from " & GetPlayerIP(index) & " has been terminated.", True)
        frmServer.Socket(index).Close
        Call UpdateCaption
        Call ClearPlayer(index)
    End If

End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If
        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & " for (" & Reason & ").")
    End If
End Sub

'                --- INSTRUCTIONS ON HOW TO TURN SAFE MODE OFF ---
'
'  INTRO:
'  Safe Mode is meant to prevent your server from getting knocked down.
'  It fixes all parse subscript out of range errors.
'  It is recommended to be kept OFF, but you should turn it on when testing.
'  If you believe a person may be trying to hack, turn it on!
'  A person may knock down your server by sending invalid packet data.
'  This prevents that and the server going down because of stupid coding errors.
'  A person has other methods to knock a server down, but this is the easiest way.

'  INSTRUCTIONS:
'  Search this module for all occurences of "SAFE MODE"
'  Follow the instructions!

Sub HandleData(ByVal index As Long, ByVal Data As String)
Dim Parse() As String ' SAFE MODE -- Uncomment for OFF, comment for ON
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
Dim i As Long, N As Long, X As Long, y As Long, f As Long
Dim MapNum As Long
Dim s As String
Dim ShopNum As Long, ItemNum As Long
Dim DurNeeded As Long, GoldNeeded As Long
Dim z As Long
Dim Packet As String
Dim o As Long
Dim TempNum As Long, TempVal As Long

    'ParseIndex = index ' SAFE MODE -- Uncomment for ON, comment for OFF

    ' Handle Data
    Parse = Split(Data, SEP_CHAR) ' SAFE MODE -- Uncomment for OFF, comment for ON
    'ZePacket = Split(Data, SEP_CHAR) ' SAFE MODE -- Uncomment for ON, comment for OFF
    'NumParse = UBound(ZePacket) ' SAFE MODE -- Uncomment for ON, comment for OFF

    ' Parse's Without Being Online
    If Not IsPlaying(index) Then

        Select Case LCase$(Parse(0))

            Case "getinfo"
                Call SendInfo(index)
                Call SendNewsTo(index)
                Exit Sub

            Case "gatglasses"
                Call SendNewCharClasses(index)
                Exit Sub

            Case "newfaccountied"
            Dim Email As String
            Dim Vault As String
                If Not IsLoggedIn(index) Then
                    Name = Parse(1)
                    Password = Parse(2)
                    Email = Parse(3)
                    Vault = Parse(4)
                    
                    For i = 1 To Len(Name)
                        N = Asc(Mid$(Name, i, 1))

                        If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                        Else
                            Call PlainMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.", 1)
                            Exit Sub
                        End If
                    Next

                    If Not AccountExist(Name) Then
                        Call AddAccount(index, Name, Password, Email, Vault)
                        Call TextAdd(frmServer.txtText(0), "Account " & Name & " has been created.", True)
                        Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                        Call PlainMsg(index, "Your account has been created!", 1)
                    Else
                        Call PlainMsg(index, "Sorry, that account name is already taken!", 1)
                    End If
                End If
                Exit Sub

            Case "logination"

                If Not IsLoggedIn(index) Then
                    Name = Parse(1)
                    Password = Parse(2)
                    For i = 1 To Len(Name)
                        N = Asc(Mid$(Name, i, 1))

                        If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                        Else
                            Call PlainMsg(index, "Account duping is not allowed!", 3)
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
                    For i = 0 To UBound(temp3)
                        temp4 = temp4 & Right$("0" & Hex$(temp3(i)), 2)
                    Next
                    temp1 = SEC_CODE
                    temp2 = CLIENT_MAJOR & "." & CLIENT_MINOR & "." & CLIENT_REVISION
                    temp3 = Encryptor.EncryptData(temp1, temp2)
                    temp5 = ""
                    For i = 0 To UBound(temp3)
                        temp5 = temp5 & Right$("0" & Hex$(temp3(i)), 2)
                    Next

                    If temp4 <> temp5 Then
                        Call SendDataTo(index, "sound" & SEP_CHAR & "ANewVersionHasBeenReleased" & SEP_CHAR & END_CHAR)
                        Call PlainMsg(index, "Version outdated, please visit " & Trim$(GetVar(App.Path & "\Data.ini", "CONFIG", "WebSite")), 3)
                        Exit Sub
                    End If

                    If Not AccountExist(Name) Then
                        Call PlainMsg(index, "That account name does not exist.", 3)
                        Exit Sub
                    End If

                    If Not PasswordOK(Name, Password) Then
                        Call PlainMsg(index, "Incorrect password.", 3)
                        Exit Sub
                    End If

                    If IsMultiAccounts(Name) Then
                        Call PlainMsg(index, "Multiple account logins is not authorized.", 3)
                        Exit Sub
                    End If

                    If frmServer.Closed.Value = Checked Then
                        Call PlainMsg(index, "The server is closed at the moment!", 3)
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
                    Packs = Packs & SPRITESIZE & SEP_CHAR
                    Packs = Packs & END_CHAR
                    Call SendDataTo(index, Packs)
                    Call LoadPlayer(index, Name)
                    Call SendChars(index)
                    Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
                    Call TextAdd(frmServer.txtText(0), GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", True)
                End If
                Exit Sub

            Case "addachara"
            Dim RacePath As Long
                Name = Parse(1)
                Sex = Val(Parse(2))
                Class = Val(Parse(3))
                CharNum = Val(Parse(4))
                RacePath = Val(Parse(5))
                
                For i = 1 To Len(Name)
                    N = Asc(Mid$(Name, i, 1))

                    If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                    Else
                        Call PlainMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.", 4)
                        Exit Sub
                    End If
                Next

                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(index, "Invalid CharNum")
                    Exit Sub
                End If

                If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
                    Call HackingAttempt(index, "Invalid Sex")
                    Exit Sub
                End If

                If Class < 1 Or Class > Max_Classes Then
                    Call HackingAttempt(index, "Invalid Class")
                    Exit Sub
                End If

                If CharExist(index, CharNum) Then
                    Call PlainMsg(index, "Character already exists!", 4)
                    Exit Sub
                End If

                If FindChar(Name) Then
                    Call PlainMsg(index, "Sorry, but that name is in use!", 4)
                    Exit Sub
                End If
                Call AddChar(index, Name, Sex, Class, CharNum, RacePath)
                Call SavePlayer(index)
                Call AddLog("Character " & Name & " added to " & GetPlayerLogin(index) & "'s account.", PLAYER_LOG)
                Call SendChars(index)
                Call PlainMsg(index, "Character has been created!", 5)
                Exit Sub

            Case "delimbocharu"
                CharNum = Val(Parse(1))

                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(index, "Invalid CharNum")
                    Exit Sub
                End If
                Call DelChar(index, CharNum)
                Call AddLog("Character deleted on " & GetPlayerLogin(index) & "'s account.", PLAYER_LOG)
                Call SendChars(index)
                Call PlainMsg(index, "Character has been deleted!", 5)
                Exit Sub

            Case "usagakarim"
                CharNum = Val(Parse(1))

                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(index, "Invalid CharNum")
                    Exit Sub
                End If

                If CharExist(index, CharNum) Then
                    Player(index).CharNum = CharNum

                    If frmServer.GMOnly.Value = Checked Then
                        If GetPlayerAccess(index) <= 0 Then
                            Call PlainMsg(index, "The server is only available to GMs at the moment!", 5)

                            'Call HackingAttempt(Index, "The server is only available to GMs at the moment!")
                            Exit Sub
                        End If
                    End If
                    Call JoinGame(index)
                    CharNum = Player(index).CharNum
                    Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
                    Call TextAdd(frmServer.txtText(0), GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".", True)
                    Call UpdateCaption

                    If Not FindChar(GetPlayerName(index)) Then
                        f = FreeFile
                        Open App.Path & "\accounts\charlist.txt" For Append As #f
                        Print #f, GetPlayerName(index)
                        Close #f
                    End If
                Else
                    Call PlainMsg(index, "Character does not exist!", 5)
                End If
                Exit Sub
        End Select
    End If

    ' Parse's With Being Online And Playing
    If IsPlaying(index) = False Then Exit Sub
    If IsConnected(index) = False Then Exit Sub

    Select Case LCase$(Parse(0))

            ' :::::::::::::::::::
            ' :: Guilds Packet ::
            ' :::::::::::::::::::
            ' Access
        Case "guildchangeaccess"

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
            Exit Sub

            ' Disown
        Case "guilddisown"

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
            Exit Sub

            ' Leave Guild
        Case "guildleave"

            ' Check if they can leave
            If GetPlayerGuild(index) = "" Then
                Call PlayerMsg(index, "You are not in a guild.", Red)
                Exit Sub
            End If
            Call SetPlayerGuild(index, "")
            Call SetPlayerGuildAccess(index, 0)
            Call SendPlayerData(index)
            Exit Sub

            ' Make A New Guild
        Case "makeguild"

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
            Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 5)
            Call SendPlayerData(FindPlayer(Parse(1)))
            Exit Sub

            ' Make A Member
        Case "guildmember"

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
            Exit Sub

            ' Make A Trainie
        Case "guildtrainee"

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
            Exit Sub
            
            ' :::::::::::::::::::::::
            ' :: Bug Report packet ::
            ' :::::::::::::::::::::::
     Case "bugreport"
    Dim BugReport As String
    Dim Message As String
    If GetPlayerLevel(index) >= 1 Then
    Message = Trim(Parse(1))
    'BugReport = Time & " " & GetPlayerName(Index) & ": " & Message
            Call AddLog(GetPlayerName(index) & ": " & Message & "", BUG_LOG)
            Call PlayerMsg(index, "Thank you for reporting this bug, " & GetPlayerName(index), White)
    Exit Sub
    End If
    
             ' :::::::::::::::::::::::
            ' :: Suggestion Report packet ::
            ' :::::::::::::::::::::::
     Case "suggestionreport"
    'Dim Message As String
    If GetPlayerLevel(index) >= 1 Then
    Message = Trim(Parse(1))
            Call AddLog(GetPlayerName(index) & ": " & Message & "", SUGGESTION_LOG)
            Call PlayerMsg(index, "Thank you for reporting your Suggestions, " & GetPlayerName(index), White)
    Exit Sub
    End If

            ' ::::::::::::::::::::
            ' :: Social packets ::
            ' ::::::::::::::::::::
        Case "saymsg"
            Msg = Parse(1)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(index, "Say Text Modification")
                    Exit Sub
                End If
            Next

            If frmServer.chkM.Value = Unchecked Then
                If GetPlayerAccess(index) <= 0 Then
                    Call PlayerMsg(index, "Map messages have been disabled by the server!", BrightRed)
                    Exit Sub
                End If
            End If
            
            'Check for swearing
            If SwearCheck(Msg) = True Then
            Call PlayerMsg(index, "Please use appropriate language.", Red)
            Exit Sub
            End If
            
            Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " : " & Msg & "", PLAYER_LOG)
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " : " & Msg & "", SayColor)
            Call MapMsg2(GetPlayerMap(index), Msg, index)
            TextAdd frmServer.txtText(3), GetPlayerName(index) & " On Map " & GetPlayerMap(index) & ": " & Msg, True
            Exit Sub

        Case "emotemsg"
            Msg = Parse(1)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(index, "Emote Text Modification")
                    Exit Sub
                End If
            Next

            If frmServer.chkE.Value = Unchecked Then
                If GetPlayerAccess(index) <= 0 Then
                    Call PlayerMsg(index, "Emote messages have been disabled by the server!", BrightRed)
                    Exit Sub
                End If
            End If
            Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Msg, PLAYER_LOG)
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & Msg, EmoteColor)
            TextAdd frmServer.txtText(6), GetPlayerName(index) & " " & Msg, True
            Exit Sub

        Case "broadcastmsg"
            Msg = Parse(1)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(index, "Broadcast Text Modification")
                    Exit Sub
                End If
            Next

            If frmServer.chkBC.Value = Unchecked Then
                If GetPlayerAccess(index) <= 0 Then
                    Call PlayerMsg(index, "Broadcast messages have been disabled by the server!", BrightRed)
                    Exit Sub
                End If
            End If

            If Player(index).Mute = True Then Exit Sub
            s = GetPlayerName(index) & ": " & Msg
            Call AddLog(s, PLAYER_LOG)
            Call GlobalMsg(s, BroadcastColor)
            Call TextAdd(frmServer.txtText(0), s, True)
            TextAdd frmServer.txtText(1), GetPlayerName(index) & ": " & Msg, True
            Exit Sub

        Case "globalmsg"
            Msg = Parse(1)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(index, "Global Text Modification")
                    Exit Sub
                End If
            Next

            If frmServer.chkG.Value = Unchecked Then
                If GetPlayerAccess(index) <= 0 Then
                    Call PlayerMsg(index, "Global messages have been disabled by the server!", BrightRed)
                    Exit Sub
                End If
            End If

            If Player(index).Mute = True Then Exit Sub
            If GetPlayerAccess(index) > 0 Then
                s = "(global) " & GetPlayerName(index) & ": " & Msg
                Call AddLog(s, ADMIN_LOG)
                Call GlobalMsg(s, GlobalColor)
                Call TextAdd(frmServer.txtText(0), s, True)
            End If
            TextAdd frmServer.txtText(2), GetPlayerName(index) & ": " & Msg, True
            Exit Sub

        Case "adminmsg"
            Msg = Parse(1)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(index, "Admin Text Modification")
                    Exit Sub
                End If
            Next

            If frmServer.chkA.Value = Unchecked Then
                Call PlayerMsg(index, "Admin messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If

            If GetPlayerAccess(index) > 0 Then
                Call AddLog("(admin " & GetPlayerName(index) & ") " & Msg, ADMIN_LOG)
                Call AdminMsg("(admin " & GetPlayerName(index) & ") " & Msg, AdminColor)
            End If
            TextAdd frmServer.txtText(5), GetPlayerName(index) & ": " & Msg, True
            Exit Sub

        Case "playermsg"
            MsgTo = FindPlayer(Parse(1))
            Msg = Parse(2)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(index, "Player Msg Text Modification")
                    Exit Sub
                End If
            Next

            If frmServer.chkP.Value = Unchecked Then
                If GetPlayerAccess(index) <= 0 Then
                    Call PlayerMsg(index, "PM messages have been disabled by the server!", BrightRed)
                    Exit Sub
                End If
            End If

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
            TextAdd frmServer.txtText(4), "To " & GetPlayerName(MsgTo) & " From " & GetPlayerName(index) & ": " & Msg, True
            Exit Sub

            ' :::::::::::::::::::::::::::::
            ' :: Moving character packet ::
            ' :::::::::::::::::::::::::::::
        Case "playermove"
            If Player(index).GettingMap = YES Then
                Exit Sub
            End If
            
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

            ' :::::::::::::::::::::::::::::
            ' :: Moving character packet ::
            ' :::::::::::::::::::::::::::::
        Case "playerdir"

            If Player(index).GettingMap = YES Then
                Exit Sub
            End If
            Dir = Val(Parse(1))

            ' Prevent hacking
            If Dir < DIR_UP Or Dir > DIR_RIGHT Then
                Call HackingAttempt(index, "Invalid Direction")
                Exit Sub
            End If
            Call SetPlayerDir(index, Dir)
            Call SendDataToMapBut(index, GetPlayerMap(index), "PLAYERDIR" & SEP_CHAR & index & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & END_CHAR)
            Exit Sub
            
    Case "poison"
          Call PoisonActive(index)
          Exit Sub
          
    Case "disease"
          Call DiseaseActive(index)
          Exit Sub
            
            ' :::::::::::::::::::::::::::
    ' :: Use Guild Deed Packet ::
    ' :::::::::::::::::::::::::::
    Case "useguilddeed"
    Dim GuildName As String
        GuildName = Trim(Parse(1))
        InvNum = Val(Parse(2))
        CharNum = Player(index).CharNum
        
              'Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
                If Not FindGuild(GuildName) Then
                    f = FreeFile
                    Open App.Path & "\accounts\Guilds.txt" For Append As #f
                        Print #f, (GuildName)
                    Close #f
        
        If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_GUILDDEED Then
            If Player(index).Char(CharNum).Guild = "" Or Player(index).Char(CharNum).Guildaccess = 0 Then
               Call SetPlayerGuild(index, GuildName)
               Call SetPlayerGuildAccess(index, 4)
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                Call SendPlayerData(index)
               Call PlayerMsg(index, "Your now the leader of " & GuildName & "!", BrightBlue)
                Call GlobalMsg("" & GetPlayerName(index) & " Is now the leader of " & GuildName & "!", BrightBlue)
                'Call AddLog(GetPlayerName(index) & " created the guild " & GuildName & ". with a guild deed", GUILD_LOG)
            Else
                Call PlayerMsg(index, "You are already in a guild!", BrightRed)
            End If
        Else
            Call PlayerMsg(index, "You need an Guild Deed to make a guild!", BrightRed)
        End If
        Else
        Call PlayerMsg(index, "Theres Already a Guild Named That!", BrightRed)
        End If
        Exit Sub
        
        Case "useitem"
            InvNum = Val(Parse(1))
            CharNum = Player(index).CharNum
            Call PerformUseItem(index, InvNum, CharNum)
            Exit Sub
        

            ' ::::::::::::::::::::::::::
            ' :: Player attack packet ::
            ' ::::::::::::::::::::::::::
        Case "attack"
            If GetPlayerWeaponSlot(index) > 0 Then
                If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 > 0 Then
                    Call SendDataToMap(GetPlayerMap(index), "checkarrows" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
            End If

            ' Try to attack a player
            For i = 1 To MAX_PLAYERS

                ' Make sure we dont try to attack ourselves
                If i <> index Then

                    ' Can we attack the player?
                    If CanAttackPlayer(index, i) Then
                        If Not CanPlayerBlockHit(i) Then

                            ' Get the damage we can do
                            If Not CanPlayerCriticalHit(index) Then
                                Damage = GetPlayerDamage(index) - GetPlayerProtection(i) + (Rnd * 5) - 2
                                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & SEP_CHAR & END_CHAR)
                            Else
                                N = GetPlayerDamage(index)
                                Damage = N + Int(Rnd * Int(N / 2)) + 1 - GetPlayerProtection(i) + (Rnd * 5) - 2
                                Call BattleMsg(index, "You feel a surge of energy upon swinging!", BrightCyan, 0)
                                Call BattleMsg(i, GetPlayerName(index) & " swings with enormous might!", BrightCyan, 1)

                                'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                                'Call PlayerMsg(I, GetPlayerName(index) & " swings with enormous might!", BrightCyan)
                                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "Blow0" & SEP_CHAR & END_CHAR)
                            End If

                            If Damage > 0 Then
                                Call AttackPlayer(index, i, Damage)
                            Else
                                Call PlayerMsg(index, "Your attack does nothing.", BrightRed)

                                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                            End If
                        Else
                            Call BattleMsg(index, GetPlayerName(i) & " blocked your hit!", BrightCyan, 0)
                            Call BattleMsg(i, "You blocked " & GetPlayerName(index) & "'s hit!", BrightCyan, 1)

                            'Call PlayerMsg(index, GetPlayerName(I) & "'s " & Trim$(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked your hit!", BrightCyan)
                            'Call PlayerMsg(I, "Your " & Trim$(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                        Exit Sub
                    End If
                End If
            Next

            ' Try to attack a npc
            For i = 1 To MAX_MAP_NPCS

                ' Can we attack the npc?
                If CanAttackNpc(index, i) Then

                    ' Get the damage we can do
                    If Not CanPlayerCriticalHit(index) Then
                        Damage = GetPlayerDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).DEF / 2) + (Rnd * 5) - 2
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & SEP_CHAR & END_CHAR)
                    Else
                        N = GetPlayerDamage(index)
                        Damage = N + Int(Rnd * Int(N / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).DEF / 2) + (Rnd * 5) - 2
                        Call BattleMsg(index, "You feel a surge of energy upon swinging!", BrightCyan, 0)

                        'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "Blow0" & SEP_CHAR & END_CHAR)
                    End If

                    If Damage > 0 Then
                        Call AttackNpc(index, i, Damage)
                        Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    Else
                        Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)

                        'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                        Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)

                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                    End If
                    Exit Sub
                End If
            Next
            
            
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                 If Player(i).CorpseMap = GetPlayerMap(index) Then
                    If CanReachCorpse(index, i) = True Then
                    Call PlayerMsg(index, "You look into " & GetPlayerName(i) & "'s corpse.", Yellow)
                    Call SendUseCorpseTo(index, i)
                    Exit Sub
                    End If
                 End If
                End If
            Next i
            Exit Sub
            
            Case "takecorpseloot"
            Call PickUpCorpseLoot(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub


            ' ::::::::::::::::::::::
            ' :: Use stats packet ::
            ' ::::::::::::::::::::::
        Case "usestatpoint"
            PointType = Val(Parse(1))

            ' Prevent hacking
            If (PointType < 0) Or (PointType > 3) Then
                Call HackingAttempt(index, "Invalid Point Type")
                Exit Sub
            End If

            ' Make sure they have points
            If GetPlayerPOINTS(index) > 0 Then
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "UsingStatPoints " & index & "," & PointType
                Else

                    Select Case PointType

                        Case 0
                            Call SetPlayerstr(index, GetPlayerstr(index) + 1)
                            Call BattleMsg(index, "You have gained more strength!", 15, 0)
                            Call SendDataTo(index, "sound" & SEP_CHAR & "strengthRaised" & SEP_CHAR & END_CHAR)

                        Case 1
                            Call SetPlayerDEF(index, GetPlayerDEF(index) + 1)
                            Call BattleMsg(index, "You have gained more defense!", 15, 0)
                            Call SendDataTo(index, "sound" & SEP_CHAR & "DefenseRaised" & SEP_CHAR & END_CHAR)

                        Case 2
                            Call SetPlayerMAGI(index, GetPlayerMAGI(index) + 1)
                            Call BattleMsg(index, "You have gained more magic abilities!", 15, 0)
                            Call SendDataTo(index, "sound" & SEP_CHAR & "MagicRaised" & SEP_CHAR & END_CHAR)

                        Case 3
                            Call SetPlayerSPEED(index, GetPlayerSPEED(index) + 1)
                            Call BattleMsg(index, "You have gained more speed!", 15, 0)
                            Call SendDataTo(index, "sound" & SEP_CHAR & "SpeedRaised" & SEP_CHAR & END_CHAR)
                    End Select
                    Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)
                End If
            Else
                Call BattleMsg(index, "You have no skill points to train with!", BrightRed, 0)
            End If
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
            Call SendFP(index)
            Call SendStats(index)
            Exit Sub

            ' ::::::::::::::::::::::::::::::::
            ' :: Player info request packet ::
            ' ::::::::::::::::::::::::::::::::
        Case "playerinforequest"
            Name = Parse(1)
            i = FindPlayer(Name)

            If i > 0 Then
                Call PlayerMsg(index, "Account: " & Trim$(Player(i).Login) & ", Name: " & GetPlayerName(i), BrightGreen)

                If GetPlayerAccess(index) > ADMIN_MONITER Then
                    Call PlayerMsg(index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
                    Call PlayerMsg(index, "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), BrightGreen)
                    Call PlayerMsg(index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
                    Call PlayerMsg(index, "str: " & GetPlayerstr(i) & "  DEF: " & GetPlayerDEF(i) & "  MAGI: " & GetPlayerMAGI(i) & "  SPEED: " & GetPlayerSPEED(i), BrightGreen)
                    N = Int(GetPlayerstr(i) / 2) + Int(GetPlayerLevel(i) / 2)
                    i = Int(GetPlayerDEF(i) / 2) + Int(GetPlayerLevel(i) / 2)

                    If N > 100 Then N = 100
                    If i > 100 Then i = 100
                    Call PlayerMsg(index, "Critical Hit Chance: " & N & "%, Block Chance: " & i & "%", BrightGreen)
                End If
            Else
                Call PlayerMsg(index, "Player is not online.", White)
            End If
            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Set sprite packet ::
            ' :::::::::::::::::::::::
        Case "setsprite"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_MAPPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If

            ' The sprite
            N = Val(Parse(1))
            Call SetPlayerSprite(index, N)
            Call SendPlayerData(index)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Set player sprite packet ::
            ' ::::::::::::::::::::::::::::::
        Case "setplayersprite"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_MAPPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If

            ' The sprite
            i = FindPlayer(Parse(1))
            N = Val(Parse(2))
            Call SetPlayerSprite(i, N)
            Call SendPlayerData(i)
            Exit Sub

            ' ::::::::::::::::::::::::::
            ' :: Stats request packet ::
            ' ::::::::::::::::::::::::::
        Case "getstats"
            Call PlayerMsg(index, "-=- Stats for " & GetPlayerName(index) & " -=-", White)
            Call PlayerMsg(index, "Level: " & GetPlayerLevel(index) & "  Exp: " & GetPlayerExp(index) & "/" & GetPlayerNextLevel(index), White)
            Call PlayerMsg(index, "HP: " & GetPlayerHP(index) & "/" & GetPlayerMaxHP(index) & "  MP: " & GetPlayerMP(index) & "/" & GetPlayerMaxMP(index) & "  SP: " & GetPlayerSP(index) & "/" & GetPlayerMaxSP(index), White)
            Call PlayerMsg(index, "str: " & GetPlayerstr(index) & "  DEF: " & GetPlayerDEF(index) & "  MAGI: " & GetPlayerMAGI(index) & "  SPEED: " & GetPlayerSPEED(index), White)
            N = Int(GetPlayerstr(index) / 2) + Int(GetPlayerLevel(index) / 2)
            i = Int(GetPlayerDEF(index) / 2) + Int(GetPlayerLevel(index) / 2)

            If N > 100 Then N = 100
            If i > 100 Then i = 100
            Call PlayerMsg(index, "Critical Hit Chance: " & N & "%, Block Chance: " & i & "%", White)
            Exit Sub

            ' ::::::::::::::::::::::::::::::::::
            ' :: Player request for a new map ::
            ' ::::::::::::::::::::::::::::::::::
        Case "requestnewmap"
            Dir = Val(Parse(1))

            ' Prevent hacking
            If Dir < DIR_UP Or Dir > DIR_RIGHT Then
                Call HackingAttempt(index, "Invalid Direction")
                Exit Sub
            End If
            Call PlayerMove(index, Dir, 1)
            Exit Sub

            ' :::::::::::::::::::::
            ' :: Map data packet ::
            ' :::::::::::::::::::::
        Case "mapdata"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_MAPPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            N = 1
            MapNum = GetPlayerMap(index)
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
            i = GetPlayerMap(index)
            For y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Ground = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).GroundSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Mask = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).MaskSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Anim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).AnimSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Fringe = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).FringeSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Type = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Data1 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Data2 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Data3 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).String1 = Parse(N)
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).String2 = Parse(N)
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).String3 = Parse(N)
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Mask2 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Mask2Set = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).M2Anim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).M2AnimSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).FAnim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).FAnimSet = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Fringe2 = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Fringe2Set = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).Light = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).F2Anim = Val(Parse(N))
                        N = N + 1
                    End If

                    If Parse(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, y).F2AnimSet = Val(Parse(N))
                        N = N + 1
                    End If
                    N = N + 1
                Next
            Next
            For X = 1 To MAX_MAP_NPCS
                Map(MapNum).Npc(X) = Val(Parse(N))
                Map(MapNum).NpcSpawn(X).Used = Val(Parse(N + 1))
                Map(MapNum).NpcSpawn(X).X = Val(Parse(N + 2))
                Map(MapNum).NpcSpawn(X).y = Val(Parse(N + 3))
                N = N + 4
                Call ClearMapNpc(X, MapNum)
            Next

            ' Clear out it all
            For i = 1 To MAX_MAP_ITEMS
                Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).X, MapItem(GetPlayerMap(index), i).y)
                Call ClearMapItem(i, GetPlayerMap(index))
            Next

            ' Save the map
            Call SaveMap(MapNum)

            ' Respawn
            Call SpawnMapItems(GetPlayerMap(index))

            ' Respawn NPCS
            Call SpawnMapNpcs(GetPlayerMap(index))

            ' Reset grid
            Call ResetMapGrid(GetPlayerMap(index))

            ' Refresh map for everyone online
            For i = 1 To MAX_PLAYERS

                If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                    Call SendDataTo(i, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & SEP_CHAR & END_CHAR)

                    Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
                End If
            Next
            Exit Sub
            
            Case "sellitem"
     Dim SellItemNum As Long
     Dim SellItemSlot As Integer
     
     SellItemNum = Parse(1)
     SellItemSlot = Parse(2)
   
     If GetPlayerWeaponSlot(index) = Val(Parse(1)) Or GetPlayerArmorSlot(index) = Val(Parse(1)) Or GetPlayerShieldSlot(index) = Val(Parse(1)) Or GetPlayerHelmetSlot(index) = Val(Parse(1)) Or GetPlayerLegsSlot(index) = Val(Parse(1)) Or GetPlayerBootsSlot(index) = Val(Parse(1)) Or GetPlayerGlovesSlot(index) = Val(Parse(1)) Or GetPlayerRing1Slot(index) = Val(Parse(1)) Or GetPlayerRing2Slot(index) = Val(Parse(1)) Or GetPlayerAmuletSlot(index) = Val(Parse(1)) Then
         Call PlayerMsg(index, "You cannot sell worn items.", Red)
         Exit Sub
     End If

              Call TakeItem(index, SellItemNum, 1)
       Call GiveItem(index, 1, Item(SellItemNum).Price)
               Call SendDataTo(index, "updatesell" & SEP_CHAR & END_CHAR)
       Exit Sub


            ' ::::::::::::::::::::::::::::
            ' :: Need map yes/no packet ::
            ' ::::::::::::::::::::::::::::
        Case "needmap"

            ' Get yes/no value
            s = LCase$(Parse(1))

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

            ' :::::::::::::::::::::::::::::::::::::::::::::::
            ' :: Player trying to pick up something packet ::
            ' :::::::::::::::::::::::::::::::::::::::::::::::
        Case "mapgetitem"
            Call PlayerMapGetItem(index)
            Exit Sub

            ' ::::::::::::::::::::::::::::::::::::::::::::
            ' :: Player trying to drop something packet ::
            ' ::::::::::::::::::::::::::::::::::::::::::::
        Case "mapdropitem"
            InvNum = Val(Parse(1))
            Amount = Val(Parse(2))

            ' Prevent hacking
            If InvNum < 1 Or InvNum > MAX_INV Then
                Call HackingAttempt(index, "Invalid InvNum")
                Exit Sub
            End If

            ' Prevent hacking
            If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then

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
            If Item(GetPlayerInvItemNum(index, InvNum)).Type <> ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                If Amount > GetPlayerInvItemValue(index, InvNum) Then
                    Call HackingAttempt(index, "Item amount modification")
                    Exit Sub
                End If
            End If
            Call PlayerMapDropItem(index, InvNum, Amount)
            Call SendStats(index)
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
            Call SendFP(index)
            Exit Sub

            ' ::::::::::::::::::::::::
            ' :: Respawn map packet ::
            ' ::::::::::::::::::::::::
        Case "maprespawn"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_MAPPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If

            ' Clear out it all
            For i = 1 To MAX_MAP_ITEMS
                Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).X, MapItem(GetPlayerMap(index), i).y)
                Call ClearMapItem(i, GetPlayerMap(index))
            Next

            ' Respawn
            Call SpawnMapItems(GetPlayerMap(index))

            ' Respawn NPCS
            Call SpawnMapNpcs(GetPlayerMap(index))
            ' Reset grid
            Call ResetMapGrid(GetPlayerMap(index))
            Call PlayerMsg(index, "Map respawned.", Blue)
            Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)
            Exit Sub

            ' ::::::::::::::::::::::::
            ' :: Kick player packet ::
            ' ::::::::::::::::::::::::
        Case "kickplayer"

            ' Prevent hacking
            If GetPlayerAccess(index) <= 0 Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If

            ' The player index
            N = FindPlayer(Parse(1))

            If N <> index Then
                If N > 0 Then
                    If GetPlayerAccess(N) <= GetPlayerAccess(index) Then
                        Call GlobalMsg(GetPlayerName(N) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(index) & "!", White)
                        Call AddLog(GetPlayerName(index) & " has kicked " & GetPlayerName(N) & ".", ADMIN_LOG)
                        Call AlertMsg(N, "You have been kicked by " & GetPlayerName(index) & "!")
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

            ' :::::::::::::::::::::
            ' :: Ban list packet ::
            ' :::::::::::::::::::::
        Case "banlist"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_MAPPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            N = 1
            f = FreeFile
            Open App.Path & "\banlist.txt" For Input As #f
            Do While Not EOF(f)
                Input #f, s
                Input #f, Name
                Call PlayerMsg(index, N & ": Banned IP " & s & " by " & Name, White)
                N = N + 1
            Loop
            Close #f
            Exit Sub

            ' ::::::::::::::::::::::::
            ' :: Ban destroy packet ::
            ' ::::::::::::::::::::::::
        Case "bandestroy"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_CREATOR Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            Call Kill(App.Path & "\banlist.txt")
            Call PlayerMsg(index, "Ban list destroyed.", White)
            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Ban player packet ::
            ' :::::::::::::::::::::::
        Case "banplayer"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_MAPPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If

            ' The player index
            N = FindPlayer(Parse(1))

            If N <> index Then
                If N > 0 Then
                    If GetPlayerAccess(N) <= GetPlayerAccess(index) Then
                        Call BanIndex(N, index)
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

            ' :::::::::::::::::::::::::::::
            ' :: Request edit map packet ::
            ' :::::::::::::::::::::::::::::
        Case "requesteditmap"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_MAPPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            Call SendDataTo(index, "EDITMAP" & SEP_CHAR & END_CHAR)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Request edit item packet ::
            ' ::::::::::::::::::::::::::::::
        Case "requestedititem"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            Call SendDataTo(index, "ITEMEDITOR" & SEP_CHAR & END_CHAR)
            Exit Sub

            ' ::::::::::::::::::::::
            ' :: Edit item packet ::
            ' ::::::::::::::::::::::
        Case "edititem"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If

            ' The item #
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_ITEMS Then
                Call HackingAttempt(index, "Invalid Item Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(index) & " editing item #" & N & ".", ADMIN_LOG)
            Call SendEditItemTo(index, N)
            Exit Sub

            ' ::::::::::::::::::::::
            ' :: Save item packet ::
            ' ::::::::::::::::::::::
        Case "saveitem"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_ITEMS Then
                Call HackingAttempt(index, "Invalid Item Index")
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

            ' Save it
            Call SendUpdateItemToAll(N)
            Call SaveItem(N)
            Call AddLog(GetPlayerName(index) & " saved item #" & N & ".", ADMIN_LOG)
            Exit Sub

            ' :::::::::::::::::::::::::::::
            ' :: Request edit npc packet ::
            ' :::::::::::::::::::::::::::::
        Case "requesteditnpc"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            Call SendDataTo(index, "NPCEDITOR" & SEP_CHAR & END_CHAR)
            Exit Sub

            ' :::::::::::::::::::::
            ' :: Edit npc packet ::
            ' :::::::::::::::::::::
        Case "editnpc"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If

            ' The npc #
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_NPCS Then
                Call HackingAttempt(index, "Invalid NPC Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(index) & " editing npc #" & N & ".", ADMIN_LOG)
            Call SendEditNpcTo(index, N)
            Exit Sub

            ' :::::::::::::::::::::
            ' :: Save npc packet ::
            ' :::::::::::::::::::::
        Case "savenpc"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_NPCS Then
                Call HackingAttempt(index, "Invalid NPC Index")
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
            z = 22
            For i = 1 To MAX_NPC_DROPS
                Npc(N).ItemNPC(i).Chance = Val(Parse(z))
                Npc(N).ItemNPC(i).ItemNum = Val(Parse(z + 1))
                Npc(N).ItemNPC(i).ItemValue = Val(Parse(z + 2))
                z = z + 3
            Next

            ' Save it
            Call SendUpdateNpcToAll(N)
            Call SaveNpc(N)
            Call AddLog(GetPlayerName(index) & " saved npc #" & N & ".", ADMIN_LOG)
            Exit Sub
            
            Case "requesteditquest"
       Call callrequstedEditQuest(index)
        Exit Sub

    Case "editquest"
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        N = Val(Parse(1))
        If N < 0 Or N > MAX_QUESTS Then
            Call HackingAttempt(index, "Invalid Quest Index")
            Exit Sub
        End If
        Call AddLog(GetPlayerName(index) & " editing quest #" & N & ".", ADMIN_LOG)
        Call SendEditQuestTo(index, N)
        Exit Sub

    Case "savequest"
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        N = Val(Parse(1))
        If N < 0 Or N > MAX_QUESTS Then
            Call HackingAttempt(index, "Invalid Quests Index")
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
        Call AddLog(GetPlayerName(index) & " saved quest #" & N & ".", ADMIN_LOG)
        Exit Sub

    Case "questdone"
      Call GiveRewardItem(index, Quest(Val(Parse(1))).RewardNum, Quest(Val(Parse(1))).RewardVal, Val(Parse(3)))
            Exit Sub
            
    Case "vault"
            Dim VaultPass As String
            VaultPass = Parse(1)
               Call VaultVerify(index, VaultPass)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Request edit shop packet ::
            ' ::::::::::::::::::::::::::::::
        Case "requesteditshop"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            Call SendDataTo(index, "SHOPEDITOR" & SEP_CHAR & END_CHAR)
            Exit Sub

            ' ::::::::::::::::::::::
            ' :: Edit shop packet ::
            ' ::::::::::::::::::::::
        Case "editshop"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If

            ' The shop #
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_SHOPS Then
                Call HackingAttempt(index, "Invalid Shop Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(index) & " editing shop #" & N & ".", ADMIN_LOG)
            Call SendEditShopTo(index, N)
            Exit Sub

        Case "addfriend"
            Name = Trim$(Parse(1))

            If Not FindChar(Name) Then
                Call PlayerMsg(index, "No such player exists!", Blue)
                Exit Sub
            End If

            If Name = GetPlayerName(index) Then
                Call PlayerMsg(index, "You can't add yourself!", Blue)
                Exit Sub
            End If
            For i = 1 To MAX_FRIENDS

                If Player(index).Char(Player(index).CharNum).Friends(i) = Name Then
                    Call PlayerMsg(index, "You already have that user as a friend!", Blue)
                    Exit Sub
                End If
            Next
            For i = 1 To MAX_FRIENDS

                If Player(index).Char(Player(index).CharNum).Friends(i) = "" Then
                    Player(index).Char(Player(index).CharNum).Friends(i) = Name
                    Call PlayerMsg(index, "Friend added.", Blue)
                    Call SendFriendListTo(index)
                    Exit Sub
                End If
            Next
            Call PlayerMsg(index, "Sorry, but you have too many friends already.", Blue)
            Exit Sub

        Case "removefriend"
            Name = Trim$(Parse(1))
            For i = 1 To MAX_FRIENDS

                If Player(index).Char(Player(index).CharNum).Friends(i) = Name Then
                    Player(index).Char(Player(index).CharNum).Friends(i) = ""
                    Call PlayerMsg(index, "Friend removed.", Blue)
                    Call SendFriendListTo(index)
                    Exit Sub
                End If
            Next
            Call PlayerMsg(index, "That person isn't on your friend list!", Blue)
            Exit Sub

            ' ::::::::::::::::::::::
            ' :: Save shop packet ::
            ' ::::::::::::::::::::::
        Case "saveshop"

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
            N = 6
            For z = 1 To 6
                For i = 1 To MAX_TRADES
                    Shop(ShopNum).TradeItem(z).Value(i).GiveItem = Val(Parse(N))
                    Shop(ShopNum).TradeItem(z).Value(i).GiveValue = Val(Parse(N + 1))
                    Shop(ShopNum).TradeItem(z).Value(i).GetItem = Val(Parse(N + 2))
                    Shop(ShopNum).TradeItem(z).Value(i).GetValue = Val(Parse(N + 3))
                    N = N + 4
                Next
            Next

            ' Save it
            Call SendUpdateShopToAll(ShopNum)
            Call SaveShop(ShopNum)
            Call AddLog(GetPlayerName(index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Request edit main packet ::
            ' ::::::::::::::::::::::::::::::
        Case "requesteditmain"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_CREATOR Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            
            f = FreeFile
            Open App.Path & "\Scripts\Main.txt" For Input As #f
                Call SendDataTo(index, "MAINEDITOR" & SEP_CHAR & Input$(LOF(f), f) & SEP_CHAR & END_CHAR)
            Close #f
            Exit Sub

            ' :::::::::::::::::::::::::::::::
            ' :: Request edit spell packet ::
            ' :::::::::::::::::::::::::::::::
        Case "requesteditspell"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            Call SendDataTo(index, "SPELLEDITOR" & SEP_CHAR & END_CHAR)
            Exit Sub
            
         ' :::::::::::::::::::::
            ' :: Day/Night Stuff ::
            ' :::::::::::::::::::::

    Case "enabledaynight"
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
    If TimeDisable = False Then
        Gamespeed = 0
        frmServer.GameTimeSpeed.text = 0
        TimeDisable = True
        frmServer.Timer1.Enabled = False
        frmServer.Command69.Caption = "Enable Time"
    Else
        Gamespeed = 1
        frmServer.GameTimeSpeed.text = 1
        TimeDisable = False
        frmServer.Timer1.Enabled = True
        frmServer.Command69.Caption = "Disable Time"
    End If
            
        Exit Sub

            ' :::::::::::::::::::::::
            ' :: Edit spell packet ::
            ' :::::::::::::::::::::::
        Case "editspell"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If

            ' The spell #
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_SPELLS Then
                Call HackingAttempt(index, "Invalid Spell Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(index) & " editing spell #" & N & ".", ADMIN_LOG)
            Call SendEditSpellTo(index, N)
            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Save spell packet ::
            ' :::::::::::::::::::::::
        Case "savespell"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If

            ' Spell #
            N = Val(Parse(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_SPELLS Then
                Call HackingAttempt(index, "Invalid Spell Index")
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
            Call AddLog(GetPlayerName(index) & " saving spell #" & N & ".", ADMIN_LOG)
            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Set access packet ::
            ' :::::::::::::::::::::::
        Case "setaccess"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_CREATOR Then
                Call HackingAttempt(index, "Trying to use powers not available")
                Exit Sub
            End If

            ' The index
            N = FindPlayer(Parse(1))

            ' The access
            i = Val(Parse(2))

            ' Check for invalid access level
            If i >= 0 Or i <= 3 Then
                If GetPlayerName(index) <> GetPlayerName(N) Then
                    If GetPlayerAccess(index) > GetPlayerAccess(N) Then

                        ' Check if player is on
                        If N > 0 Then
                            If GetPlayerAccess(N) <= 0 Then
                                Call GlobalMsg(GetPlayerName(N) & " has been blessed with administrative access.", BrightBlue)
                            End If
                            Call SetPlayerAccess(N, i)
                            Call SendPlayerData(N)
                            Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(N) & "'s access.", ADMIN_LOG)
                        Else
                            Call PlayerMsg(index, "Player is not online.", White)
                        End If
                    Else
                        Call PlayerMsg(index, "Your access level is lower than " & GetPlayerName(N) & "´s.", Red)
                    End If
                Else
                    Call PlayerMsg(index, "You cant change your access.", Red)
                End If
            Else
                Call PlayerMsg(index, "Invalid access level.", Red)
            End If
            Exit Sub

        Case "whosonline"
            Call SendWhosOnline(index)
            Exit Sub

        Case "onlinelist"
            Call SendOnlineList
            Exit Sub

        Case "setmotd"

            If GetPlayerAccess(index) < ADMIN_MAPPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            Call SpecialPutVar(App.Path & "\motd.ini", "MOTD", "Msg", Parse(1))
            Call GlobalMsg("MOTD changed to: " & Parse(1), BrightCyan)
            Call AddLog(GetPlayerName(index) & " changed MOTD to: " & Parse(1), ADMIN_LOG)
            Exit Sub

        Case "traderequest"
        ' Trade num
        N = Val(Parse(1))
        z = Val(Parse(2))
        
        ' Prevent hacking
        If (N < 1) Or (N > 6) Then
            Call HackingAttempt(index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Prevent hacking
        If (z <= 0) Or (z > (MAX_TRADES * 6)) Then
            Call HackingAttempt(index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Index for shop
        i = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        'Check if its furniture
        If Item(Shop(i).TradeItem(N).Value(z).GetItem).Type = ITEM_TYPE_FURNITURE Then
        If HasItem(index, Shop(i).TradeItem(N).Value(z).GiveItem) >= Shop(i).TradeItem(N).Value(z).GiveValue And Player(index).Char(Player(index).CharNum).Hands = 0 Then
        Call TakeItem(index, Shop(i).TradeItem(N).Value(z).GiveItem, Shop(i).TradeItem(N).Value(z).GiveValue)
        Player(index).Char(Player(index).CharNum).Hands = Shop(i).TradeItem(N).Value(z).GetItem
        Call SendDataTo(index, "Sethands" & SEP_CHAR & Shop(i).TradeItem(N).Value(z).GetItem & SEP_CHAR & END_CHAR)
        Call PlayerMsg(index, "The trade was successful!", Yellow)
        Exit Sub
                Else
            Call PlayerMsg(index, "Trade unsuccessful.", BrightRed)
            Exit Sub
        End If
        End If
        
        ' Check if inv full
        If i <= 0 Then Exit Sub
        X = FindOpenInvSlot(index, Shop(i).TradeItem(N).Value(z).GetItem)
        If X = 0 Then
            Call PlayerMsg(index, "Trade unsuccessful, inventory full.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have the item
        If HasItem(index, Shop(i).TradeItem(N).Value(z).GiveItem) >= Shop(i).TradeItem(N).Value(z).GiveValue Then
            Call TakeItem(index, Shop(i).TradeItem(N).Value(z).GiveItem, Shop(i).TradeItem(N).Value(z).GiveValue)
            Call GiveItem(index, Shop(i).TradeItem(N).Value(z).GetItem, Shop(i).TradeItem(N).Value(z).GetValue)
            Call PlayerMsg(index, "The trade was successful!", Yellow)
        Else
            Call PlayerMsg(index, "Trade unsuccessful.", BrightRed)
        End If
        Exit Sub

        Case "fixitem"

            ' Inv num
            N = Val(Parse(1))

            ' Make sure its a equipable item
            If Item(GetPlayerInvItemNum(index, N)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(index, N)).Type > ITEM_TYPE_SHIELD Or Item(GetPlayerInvItemNum(index, N)).Type > ITEM_TYPE_LEGS Or Item(GetPlayerInvItemNum(index, N)).Type > ITEM_TYPE_BOOTS Or Item(GetPlayerInvItemNum(index, N)).Type > ITEM_TYPE_GLOVES Or Item(GetPlayerInvItemNum(index, N)).Type > ITEM_TYPE_RING1 Or Item(GetPlayerInvItemNum(index, N)).Type > ITEM_TYPE_RING2 Or Item(GetPlayerInvItemNum(index, N)).Type > ITEM_TYPE_AMULET Then
                Call PlayerMsg(index, "You can only fix weapons, armors, helmets, and shields, legs, boots, gloves, rings And amulets.", BrightRed)
                Exit Sub
            End If

            ' Check if they have a full inventory
            If FindOpenInvSlot(index, GetPlayerInvItemNum(index, N)) <= 0 Then
                Call PlayerMsg(index, "You have no inventory space left!", BrightRed)
                Exit Sub
            End If

            ' Check if you can actually repair the item
            If Item(ItemNum).Data1 < 0 Then
                Call PlayerMsg(index, "This item isn't repairable!", BrightRed)
                Exit Sub
            End If

            ' Now check the rate of pay
            ItemNum = GetPlayerInvItemNum(index, N)
            i = Int(Item(GetPlayerInvItemNum(index, N)).Data2 / 5)

            If i <= 0 Then i = 1
            DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(index, N)
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
                    Call SetPlayerInvItemDur(index, N, Item(ItemNum).Data1 * -1)
                    Call PlayerMsg(index, "Item has been totally restored for " & GoldNeeded & " gold!", BrightBlue)
                Else

                    ' They dont so restore as much as we can
                    DurNeeded = (HasItem(index, 1) / i)
                    GoldNeeded = Int(DurNeeded * i / 2)

                    If GoldNeeded <= 0 Then GoldNeeded = 1
                    Call TakeItem(index, 1, GoldNeeded)
                    Call SetPlayerInvItemDur(index, N, GetPlayerInvItemDur(index, N) + DurNeeded)
                    Call PlayerMsg(index, "Item has been partially fixed for " & GoldNeeded & " gold!", BrightBlue)
                End If
            Else
                Call PlayerMsg(index, "Insufficient gold to fix this item!", BrightRed)
            End If
            Exit Sub

        Case "search"
            X = Val(Parse(1))
            y = Val(Parse(2))

            ' Prevent subscript out of range
            If X < 0 Or X > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then
                Exit Sub
            End If

            ' Check for a player
            For i = 1 To MAX_PLAYERS

                If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) And GetPlayerX(i) = X And GetPlayerY(i) = y Then

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
            Next

            ' Check for an npc
            For i = 1 To MAX_MAP_NPCS

                If MapNpc(GetPlayerMap(index), i).num > 0 Then
                    If MapNpc(GetPlayerMap(index), i).X = X And MapNpc(GetPlayerMap(index), i).y = y Then

                        ' Change target
                        Player(index).Target = i
                        Player(index).TargetType = TARGET_TYPE_NPC
                        Call PlayerMsg(index, "Your target is now a " & Trim$(Npc(MapNpc(GetPlayerMap(index), i).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If
            Next

            ' Check for an item
            For i = 1 To MAX_MAP_ITEMS

                If MapItem(GetPlayerMap(index), i).num > 0 Then
                    If MapItem(GetPlayerMap(index), i).X = X And MapItem(GetPlayerMap(index), i).y = y Then
                        Call PlayerMsg(index, "You see a " & Trim$(Item(MapItem(GetPlayerMap(index), i).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If
            Next
            Exit Sub

        Case "playerchat"
            N = FindPlayer(Parse(1))

            If N < 1 Then
                Call PlayerMsg(index, "Player is not online.", White)
                Exit Sub
            End If

            If N = index Then
                Exit Sub
            End If

            If Player(index).InChat = 1 Then
                Call PlayerMsg(index, "Your already in a chat with another player!", Pink)
                Exit Sub
            End If

            If Player(N).InChat = 1 Then
                Call PlayerMsg(index, "Player is already in a chat with another player!", Pink)
                Exit Sub
            End If
            Call PlayerMsg(index, "Chat request has been sent to " & GetPlayerName(N) & ".", Pink)
            Call PlayerMsg(N, GetPlayerName(index) & " wants you to chat with them.  Type /chat to accept, or /chatdecline to decline.", Pink)
            Player(N).ChatPlayer = index
            Player(index).ChatPlayer = N
            Exit Sub

        Case "achat"
            N = Player(index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(index, "No one requested to chat with you.", Pink)
                Exit Sub
            End If

            If Player(N).ChatPlayer <> index Then
                Call PlayerMsg(index, "Chat failed.", Pink)
                Exit Sub
            End If
            Call SendDataTo(index, "PPCHATTING" & SEP_CHAR & N & SEP_CHAR & END_CHAR)
            Call SendDataTo(N, "PPCHATTING" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
            Exit Sub
            
        Case "forgetspell"
' Spell slot
N = CLng(Parse(1))

' Prevent subscript out of range
If N <= 0 Or N > MAX_PLAYER_SPELLS Then
Call HackingAttempt(index, "Invalid Spell Slot")
Exit Sub
End If

With Player(index).Char(Player(index).CharNum)
If .Spell(N) = 0 Then
Call PlayerMsg(index, "No spell here.", Red)

Else
Call PlayerMsg(index, "You have forgotten the spell """ & Trim$(Spell(.Spell(N)).Name) & """", Green)

.Spell(N) = 0
Call SendSpells(index)
End If
End With
Exit Sub

        Case "dchat"
            N = Player(index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(index, "No one requested to chat with you.", Pink)
                Exit Sub
            End If
            Call PlayerMsg(index, "Declined chat request.", Pink)
            Call PlayerMsg(N, GetPlayerName(index) & " declined your request.", Pink)
            Player(index).ChatPlayer = 0
            Player(index).InChat = 0
            Player(N).ChatPlayer = 0
            Player(N).InChat = 0
            Exit Sub

        Case "qchat"
            N = Player(index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(index, "No one requested to chat with you.", Pink)
                Exit Sub
            End If
            Call SendDataTo(index, "qchat" & SEP_CHAR & END_CHAR)
            Call SendDataTo(N, "qchat" & SEP_CHAR & END_CHAR)
            Player(index).ChatPlayer = 0
            Player(index).InChat = 0
            Player(N).ChatPlayer = 0
            Player(N).InChat = 0
            Exit Sub

        Case "sendchat"
            N = Player(index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(index, "No one requested to chat with you.", Pink)
                Exit Sub
            End If
            Call SendDataTo(N, "sendchat" & SEP_CHAR & Parse(1) & SEP_CHAR & index & SEP_CHAR & END_CHAR)
            Exit Sub

        Case "pptrade"
            N = FindPlayer(Parse(1))

            ' Check if player is online
            If N < 1 Then
                Call PlayerMsg(index, "Player is not online.", White)
                Exit Sub
            End If

            ' Prevent trading with self
            If N = index Then
                Exit Sub
            End If

            ' Check if the player is in another trade
            If Player(index).InTrade = 1 Then
                Call PlayerMsg(index, "Your already in a trade with someone else!", Pink)
                Exit Sub
            End If

            For i = 0 To 3
                If DirToX(GetPlayerX(index), i) = GetPlayerX(N) And DirToY(GetPlayerY(index), i) = GetPlayerY(N) Then
                    ' Check to see if player is already in a trade
                    If Player(N).InTrade = 1 Then
                        Call PlayerMsg(index, "Player is already in a trade!", Pink)
                        Exit Sub
                    End If
                    Call PlayerMsg(index, "Trade request has been sent to " & GetPlayerName(N) & ".", Pink)
                    Call PlayerMsg(N, GetPlayerName(index) & " wants you to trade with them.  Type /accept to accept, or /decline to decline.", Pink)
                    Player(N).TradePlayer = index
                    Player(index).TradePlayer = N
                    Exit Sub
                End If
            Next
            
            Call PlayerMsg(index, "You need to be beside the player to trade!", Pink)
            Call PlayerMsg(N, "The player needs to be beside you to trade!", Pink)
            Exit Sub

        Case "atrade"
            N = Player(index).TradePlayer

            ' Check if anyone requested a trade
            If N < 1 Then
                Call PlayerMsg(index, "No one requested a trade with you.", Pink)
                Exit Sub
            End If

            ' Check if its the right player
            If Player(N).TradePlayer <> index Then
                Call PlayerMsg(index, "Trade failed.", Pink)
                Exit Sub
            End If

            ' Check where both players are
            For i = 0 To 3
                If DirToX(GetPlayerX(index), i) = GetPlayerX(N) And DirToY(GetPlayerY(index), i) = GetPlayerY(N) Then
                    Call PlayerMsg(index, "You are trading with " & GetPlayerName(N) & "!", Pink)
                    Call PlayerMsg(N, GetPlayerName(index) & " accepted your trade request!", Pink)
                    Call SendDataTo(index, "PPTRADING" & SEP_CHAR & END_CHAR)
                    Call SendDataTo(N, "PPTRADING" & SEP_CHAR & END_CHAR)
                    For o = 1 To MAX_PLAYER_TRADES
                        Player(index).Trading(o).InvNum = 0
                        Player(index).Trading(o).InvName = ""
                        Player(N).Trading(o).InvNum = 0
                        Player(N).Trading(o).InvName = ""
                    Next
                    Player(index).InTrade = 1
                    Player(index).TradeItemMax = 0
                    Player(index).TradeItemMax2 = 0
                    Player(N).InTrade = 1
                    Player(N).TradeItemMax = 0
                    Player(N).TradeItemMax2 = 0
                    Exit Sub
                End If
            Next
            
            Call PlayerMsg(index, "The player needs to be beside you to trade!", Pink)
            Call PlayerMsg(N, "You need to be beside the player to trade!", Pink)
            Exit Sub

        Case "qtrade"
            N = Player(index).TradePlayer

            ' Check if anyone trade with player
            If N < 1 Then
                Call PlayerMsg(index, "No one requested a trade with you.", Pink)
                Exit Sub
            End If
            Call PlayerMsg(index, "Stopped trading.", Pink)
            Call PlayerMsg(N, GetPlayerName(index) & " stopped trading with you!", Pink)
            Player(index).TradeOk = 0
            Player(N).TradeOk = 0
            Player(index).TradePlayer = 0
            Player(index).InTrade = 0
            Player(N).TradePlayer = 0
            Player(N).InTrade = 0
            Call SendDataTo(index, "qtrade" & SEP_CHAR & END_CHAR)
            Call SendDataTo(N, "qtrade" & SEP_CHAR & END_CHAR)
            Exit Sub

        Case "dtrade"
            N = Player(index).TradePlayer

            ' Check if anyone trade with player
            If N < 1 Then
                Call PlayerMsg(index, "No one requested a trade with you.", Pink)
                Exit Sub
            End If
            Call PlayerMsg(index, "Declined trade request.", Pink)
            Call PlayerMsg(N, GetPlayerName(index) & " declined your request.", Pink)
            Player(index).TradePlayer = 0
            Player(index).InTrade = 0
            Player(N).TradePlayer = 0
            Player(N).InTrade = 0
            Exit Sub

        Case "updatetradeinv"
            N = Val(Parse(1))
            Player(index).Trading(N).InvNum = Val(Parse(2))
            Player(index).Trading(N).InvName = Trim$(Parse(3))
            Player(index).Trading(N).InvVal = Val(Parse(4))
            If Player(index).Trading(N).InvNum = 0 Then
                Player(index).TradeItemMax = Player(index).TradeItemMax - 1
                Player(index).TradeOk = 0
                Player(N).TradeOk = 0
                Call SendDataTo(index, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                Call SendDataTo(N, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Else
                Player(index).TradeItemMax = Player(index).TradeItemMax + 1
            End If
            Call SendDataTo(Player(index).TradePlayer, "updatetradeitem" & SEP_CHAR & N & SEP_CHAR & Player(index).Trading(N).InvNum & SEP_CHAR & Player(index).Trading(N).InvName & SEP_CHAR & Player(index).Trading(N).InvVal & SEP_CHAR & END_CHAR)
            Exit Sub


        Case "swapitems"
            Dim cur As Long
            
            N = Player(index).TradePlayer

            If Player(index).TradeOk = 0 Then
                Player(index).TradeOk = 1
                Call SendDataTo(N, "trading" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            ElseIf Player(index).TradeOk = 1 Then
                Player(index).TradeOk = 0
                Call SendDataTo(N, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            End If

            If Player(index).TradeOk = 1 And Player(N).TradeOk = 1 Then
                Player(index).TradeItemMax2 = 0
                Player(N).TradeItemMax2 = 0
                For i = 1 To MAX_INV

                    If Player(index).TradeItemMax = Player(index).TradeItemMax2 Then
                        Exit For
                    End If

                    If GetPlayerInvItemNum(N, i) < 1 Then
                        Player(index).TradeItemMax2 = Player(index).TradeItemMax2 + 1
                    End If
                Next
                For i = 1 To MAX_INV

                    If Player(N).TradeItemMax = Player(N).TradeItemMax2 Then
                        Exit For
                    End If

                    If GetPlayerInvItemNum(index, i) < 1 Then
                        Player(N).TradeItemMax2 = Player(N).TradeItemMax2 + 1
                    End If
                Next

                If Player(index).TradeItemMax2 = Player(index).TradeItemMax And Player(N).TradeItemMax2 = Player(N).TradeItemMax Then
                    For i = 1 To MAX_PLAYER_TRADES
                        For X = 1 To MAX_INV

                            If GetPlayerInvItemNum(N, X) < 1 Then
                                If Player(index).Trading(i).InvNum > 0 Then
                                    If Player(index).Trading(i).InvVal > 0 Then
                                    Call GiveItem(N, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), Player(index).Trading(i).InvVal)
                                    Call TakeItem(index, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), Player(index).Trading(i).InvVal)
                                    Else
                                    Call GiveItem(N, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), 1)
                                    Call TakeItem(index, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), 1)
                                    End If
                                    Exit For
                                End If
                            End If
                        Next
                    Next
                    For i = 1 To MAX_PLAYER_TRADES
                        For X = 1 To MAX_INV

                            If GetPlayerInvItemNum(index, X) < 1 Then
                                If Player(N).Trading(i).InvNum > 0 Then
                                    If Player(N).Trading(i).InvVal > 0 Then
                                    Call GiveItem(index, GetPlayerInvItemNum(N, Player(N).Trading(i).InvNum), Player(N).Trading(i).InvVal)
                                    Call TakeItem(N, GetPlayerInvItemNum(N, Player(N).Trading(i).InvNum), Player(N).Trading(i).InvVal)
                                    Else
                                    Call GiveItem(index, GetPlayerInvItemNum(N, Player(N).Trading(i).InvNum), 1)
                                    Call TakeItem(N, GetPlayerInvItemNum(N, Player(N).Trading(i).InvNum), 1)
                                    End If
                                    Exit For
                                End If
                            End If
                        Next
                    Next
                    Call PlayerMsg(N, "Trade Successfull!", BrightGreen)
                    Call PlayerMsg(index, "Trade Successfull!", BrightGreen)
                    Call SendInventory(N)
                    Call SendInventory(index)
                Else

                    If Player(index).TradeItemMax2 < Player(index).TradeItemMax Then
                        Call PlayerMsg(index, "Your inventory is full!", BrightRed)
                        Call PlayerMsg(N, GetPlayerName(index) & "'s inventory is full!", BrightRed)
                    End If

                    If Player(N).TradeItemMax2 < Player(N).TradeItemMax Then
                        Call PlayerMsg(N, "Your inventory is full!", BrightRed)
                        Call PlayerMsg(index, GetPlayerName(N) & "'s inventory is full!", BrightRed)
                    End If
                End If
                Player(index).TradePlayer = 0
                Player(index).InTrade = 0
                Player(index).TradeOk = 0
                Player(N).TradePlayer = 0
                Player(N).InTrade = 0
                Player(N).TradeOk = 0
                Call SendDataTo(index, "qtrade" & SEP_CHAR & END_CHAR)
                Call SendDataTo(N, "qtrade" & SEP_CHAR & END_CHAR)
            End If
            Exit Sub


        Case "party"
            N = FindPlayer(Parse(1))

            If N = index Then Exit Sub
            If N > 0 Then
                If GetPlayerAccess(index) > ADMIN_MONITER Then
                    Call PlayerMsg(index, "You can't join a party, you are an admin!", BrightBlue)
                    Exit Sub
                End If

                If GetPlayerAccess(N) > ADMIN_MONITER Then
                    Call PlayerMsg(index, "Admins cannot join parties!", BrightBlue)
                    Exit Sub
                End If

                If Player(N).InParty = NO Then
                    If Player(index).PartyID > 0 Then
                        If Party(Player(index).PartyID).Member(MAX_PARTY_MEMBERS) <> 0 Then
                            Call PlayerMsg(index, GetPlayerName(N) & " has been invited to your party.", Pink)
                            Call PlayerMsg(N, GetPlayerName(index) & " has invited you to join their party.  Type /join to join, or /leave to decline.", Pink)
                            Player(N).Invited = Player(index).PartyID
                        Else
                            Call PlayerMsg(index, "Your party is full.", Pink)
                        End If
                    Else
                        o = 0
                        i = MAX_PARTIES
                        Do While i > 0

                            If Party(i).Member(1) = 0 Then o = i
                            i = i - 1

                        Loop

                        If o = 0 Then
                            Call PlayerMsg(index, "Party overload.", Pink)
                            Exit Sub
                        End If
                        Party(o).Member(1) = index
                        Player(index).InParty = YES
                        Player(index).PartyID = o
                        Player(index).Invited = 0
                        Call PlayerMsg(index, "Party created.", Pink)
                        Call PlayerMsg(index, GetPlayerName(N) & " has been invited to your party.", Pink)
                        Call PlayerMsg(N, GetPlayerName(index) & " has invited you to join their party.  Type /join to join, or /leave to decline.", Pink)
                        Player(N).Invited = Player(index).PartyID
                    End If
                Else
                    Call PlayerMsg(index, "Player is already in a party.", Pink)
                End If
            Else
                Call PlayerMsg(index, "Player is not online.", White)
            End If
            Exit Sub

        Case "joinparty"

            If Player(index).Invited > 0 Then
                o = 0
                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(index).Invited).Member(i) = 0 Then
                        If o = 0 Then o = i
                    End If
                Next

                If o <> 0 Then
                    Player(index).PartyID = Player(index).Invited
                    Player(index).InParty = YES
                    Player(index).Invited = 0
                    Party(Player(index).PartyID).Member(o) = index
                    For i = 1 To MAX_PARTY_MEMBERS

                        If Party(Player(index).PartyID).Member(i) <> 0 And Party(Player(index).PartyID).Member(i) <> index Then
                            Call PlayerMsg(Party(Player(index).PartyID).Member(i), GetPlayerName(index) & " has joined your party!", Pink)
                        End If
                    Next
                    Call PlayerMsg(index, "You have joined the party!", Pink)
                Else
                    Call PlayerMsg(index, "The party is full!", Pink)
                End If
            Else
                Call PlayerMsg(index, "You have not been invited into a party!", Pink)
            End If
            Exit Sub

        Case "leaveparty"

            If Player(index).PartyID > 0 Then
                Call PlayerMsg(index, "You have left the party.", Pink)
                N = 0
                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(index).PartyID).Member(i) = index Then N = i
                Next
                For i = N To MAX_PARTY_MEMBERS - 1
                    Party(Player(index).PartyID).Member(i) = Party(Player(index).PartyID).Member(i + 1)
                Next
                Party(Player(index).PartyID).Member(MAX_PARTY_MEMBERS) = 0
                N = 0
                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(index).PartyID).Member(i) <> 0 And Party(Player(index).PartyID).Member(i) <> index Then
                        N = N + 1
                        Call PlayerMsg(Party(Player(index).PartyID).Member(i), GetPlayerName(index) & " has left the party.", Pink)
                    End If
                Next

                If N < 2 Then
                    Call PlayerMsg(Party(Player(index).PartyID).Member(1), "The party has disbanded.", Pink)
                    Player(Party(Player(index).PartyID).Member(1)).InParty = NO
                    Player(Party(Player(index).PartyID).Member(1)).PartyID = 0
                    Party(Player(index).PartyID).Member(1) = 0
                End If
                Player(index).InParty = NO
                Player(index).PartyID = 0
            Else

                If Player(index).Invited <> 0 Then
                    For i = 1 To MAX_PARTY_MEMBERS

                        If Party(Player(index).Invited).Member(i) <> 0 And Party(Player(index).Invited).Member(i) <> index Then Call PlayerMsg(index, GetPlayerName(index) & " has declined the invitation.", Pink)
                    Next
                    Player(index).Invited = 0
                    Call PlayerMsg(index, "You have declined the invitation.", Pink)
                Else
                    Call PlayerMsg(index, "You have not been invited into a party!", Pink)
                End If
            End If
            Exit Sub

        Case "partychat"

            If Player(index).PartyID > 0 Then
                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(index).PartyID).Member(i) <> 0 Then Call PlayerMsg(Party(Player(index).PartyID).Member(i), Parse(1), PartyColor)
                Next
            Else
                Call PlayerMsg(index, "You are not in a party!", Pink)
            End If
            Exit Sub

        Case "guildchat"

            If GetPlayerGuild(index) <> "" Then
                For i = 1 To MAX_PLAYERS

                    If GetPlayerGuild(index) = GetPlayerGuild(i) Then Call PlayerMsg(i, GetPlayerGuild(index) & "-" & GetPlayerName(index) & ": " & Parse(1), GuildColor)
                Next
            Else
                Call PlayerMsg(index, "You are not in a guild!", Pink)
            End If
            Exit Sub

        Case "newmain"

            If GetPlayerAccess(index) >= ADMIN_CREATOR Then
                Dim temp As String

                f = FreeFile
                Open App.Path & "\Scripts\Main.txt" For Input As #f
                temp = Input$(LOF(f), f)
                Close #f
                f = FreeFile
                Open App.Path & "\Scripts\Backup.txt" For Output As #f
                Print #f, temp
                Close #f
                f = FreeFile
                Open App.Path & "\Scripts\Main.txt" For Output As #f
                Print #f, Parse(1)
                Close #f

                If SCRIPTING = 1 Then
                    Set MyScript = Nothing
                    Set clsScriptCommands = Nothing
                    Set MyScript = New clsSadScript
                    Set clsScriptCommands = New clsCommands
                    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
                    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
                    Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
                    Call PlayerMsg(index, "Scripts reloaded.", White)
                End If
                Call AddLog(GetPlayerName(index) & " updated the script.", ADMIN_LOG)
            End If
            Exit Sub

        Case "requestbackupmain"

            If GetPlayerAccess(index) >= ADMIN_CREATOR Then
                Dim nothertemp As String

                f = FreeFile
                Open App.Path & "\Scripts\Backup.txt" For Input As #f
                nothertemp = Input$(LOF(f), f)
                Close #f
                f = FreeFile
                Open App.Path & "\Scripts\Main.txt" For Output As #f
                Print #f, nothertemp
                Close #f

                If SCRIPTING = 1 Then
                    Set MyScript = Nothing
                    Set clsScriptCommands = Nothing
                    Set MyScript = New clsSadScript
                    Set clsScriptCommands = New clsCommands
                    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
                    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
                    Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
                    Call PlayerMsg(index, "Scripts reloaded.", White)
                End If
                Call AddLog(GetPlayerName(index) & " used the backup script.", ADMIN_LOG)
            End If
            Exit Sub

        Case "spells"
            Call SendPlayerSpells(index)
            Exit Sub
            
            Case "setattribute"
        Call ScriptSetAttribute(Val(Parse(1)), Val(Parse(2)), Val(Parse(3)), Val(Parse(4)), Val(Parse(5)), Val(Parse(6)), Val(Parse(7)), Val(Parse(8)), Val(Parse(9)), Val(Parse(10)))
        Exit Sub

        Case "cast"
            N = Val(Parse(1))
            Call CastSpell(index, N)
            Exit Sub

        Case "requestlocation"

            If GetPlayerAccess(index) < ADMIN_MAPPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            Call PlayerMsg(index, "Map: " & GetPlayerMap(index) & ", X: " & GetPlayerX(index) & ", Y: " & GetPlayerY(index), Pink)
            Exit Sub

        Case "refresh"
            Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), False)
            Call PlayerMsg(index, "Map refreshed.", White)
            Exit Sub

        Case "killpet"
            Player(index).Pet.Alive = NO
            Player(index).Pet.Sprite = 0
            Call TakeFromGrid(Player(index).Pet.Map, Player(index).Pet.X, Player(index).Pet.y)
            Packet = "PETDATA" & SEP_CHAR
            Packet = Packet & index & SEP_CHAR
            Packet = Packet & Player(index).Pet.Alive & SEP_CHAR
            Packet = Packet & Player(index).Pet.Map & SEP_CHAR
            Packet = Packet & Player(index).Pet.X & SEP_CHAR
            Packet = Packet & Player(index).Pet.y & SEP_CHAR
            Packet = Packet & Player(index).Pet.Dir & SEP_CHAR
            Packet = Packet & Player(index).Pet.Sprite & SEP_CHAR
            Packet = Packet & Player(index).Pet.HP & SEP_CHAR
            Packet = Packet & Player(index).Pet.Level * 5 & SEP_CHAR
            Packet = Packet & END_CHAR
            Call SendDataToMap(GetPlayerMap(index), Packet)
            Exit Sub

        Case "petmoveselect"
            X = Val(Parse(1))
            y = Val(Parse(2))
            Player(index).Pet.MapToGo = GetPlayerMap(index)
            Player(index).Pet.Target = 0
            Player(index).Pet.XToGo = X
            Player(index).Pet.YToGo = y
            Player(index).Pet.AttackTimer = GetTickCount
            For i = 1 To MAX_PLAYERS

                If IsPlaying(i) Then
                    If GetPlayerMap(i) = Player(index).Pet.Map Then
                        If GetPlayerX(i) = X And GetPlayerY(i) = y Then
                            Player(index).Pet.TargetType = TARGET_TYPE_PLAYER
                            Player(index).Pet.Target = i
                            Call PlayerMsg(index, "Your pet's target is now " & Trim$(GetPlayerName(i)) & ".", Yellow)
                            Exit Sub
                        End If
                    End If
                End If
            Next
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(Player(index).Pet.Map, i).num > 0 Then
                    If MapNpc(Player(index).Pet.Map, i).X = X And MapNpc(Player(index).Pet.Map, i).y = y Then
                        Player(index).Pet.TargetType = TARGET_TYPE_NPC
                        Player(index).Pet.Target = i
                        Call PlayerMsg(index, "Your pet's target is now a " & Trim$(Npc(MapNpc(Player(index).Pet.Map, i).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If
            Next
            Call PlayerMsg(index, "Pet is moving to (" & X & "," & y & ")", Yellow)
            Exit Sub

        Case "buysprite"

            ' Check if player stepped on sprite changing tile
            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
                Call PlayerMsg(index, "You need to be on a sprite tile to buy it!", BrightRed)
                Exit Sub
            End If

            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 = 0 Then
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

                        If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerBootsSlot(index) <> i And GetPlayerGlovesSlot(index) <> i And GetPlayerRing1Slot(index) <> i And GetPlayerRing2Slot(index) <> i And GetPlayerAmuletSlot(index) <> i Then
                            Call SetPlayerInvItemNum(index, i, 0)
                            Call PlayerMsg(index, "You have bought a new sprite!", BrightGreen)
                            Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                            Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
                            Call SendInventory(index)
                        End If
                    End If

                    If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerBootsSlot(index) <> i And GetPlayerGlovesSlot(index) <> i And GetPlayerRing1Slot(index) <> i And GetPlayerRing2Slot(index) <> i And GetPlayerAmuletSlot(index) <> i Then
                        Exit Sub
                    End If
                End If
            Next
            Call PlayerMsg(index, "You dont have enough to buy this sprite!", BrightRed)
            Exit Sub
            
            Case "clearowner"
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        Call PlayerMsg(index, "Owner cleared!", BrightRed)
        Map(GetPlayerMap(index)).Owner = ""
        Map(GetPlayerMap(index)).Name = "Unowned House"
        Map(GetPlayerMap(index)).Revision = Map(GetPlayerMap(index)).Revision + 1
        Call SaveMap(GetPlayerMap(index))
        Call SendDataToAll("maphouseupdate" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Owner) & SEP_CHAR & (Map(GetPlayerMap(index)).Name) & SEP_CHAR & END_CHAR)
        Exit Sub
        
        Case "buyhouse"
        Call PlayerBuyHouse(index)
        Exit Sub
        
        Case "sellhouse"
        Call PlayerBuyHouse(index)
        Exit Sub

        Case "checkcommands"
            s = Parse(1)

            If SCRIPTING = 1 Then
                PutVar App.Path & "\Scripts\Command.ini", "TEMP", "Text" & index, Trim$(s)
                MyScript.ExecuteStatement "Scripts\Main.txt", "Commands " & index
            Else
                Call PlayerMsg(index, "Thats not a valid command!", 12)
            End If
            Exit Sub

        Case "prompt"

            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerPrompt " & index & "," & Val(Parse(1)) & "," & Val(Parse(2))
            End If
            Exit Sub

        Case "requesteditarrow"

            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            Call SendDataTo(index, "ARROWEDITOR" & SEP_CHAR & END_CHAR)
            Exit Sub

        Case "editarrow"

            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_ARROWS Then
                Call HackingAttempt(index, "Invalid arrow Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(index) & " editing arrow #" & N & ".", ADMIN_LOG)
            Call SendEditArrowTo(index, N)
            Exit Sub

        Case "savearrow"

            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_ITEMS Then
                Call HackingAttempt(index, "Invalid arrow Index")
                Exit Sub
            End If
            Arrows(N).Name = Parse(2)
            Arrows(N).Pic = Val(Parse(3))
            Arrows(N).Range = Val(Parse(4))
            Call SendUpdateArrowToAll(N)
            Call SaveArrow(N)
            Call AddLog(GetPlayerName(index) & " saved arrow #" & N & ".", ADMIN_LOG)
            Exit Sub
            
        Case "gofishing"
    Dim Tool As String
    Tool = GetPlayerWeaponSlot(index)
    If Tool = 0 Then
    Call PlayerMsg(index, "You don't have a " & Parse(4) & " equipped", 15)
    ElseIf GetPlayerInvItemNum(index, Tool) = Val(Parse(1)) Then
       Call GoFishing(index, Val(Parse(3)), RandomNo(MAX_LEVEL + RandomNo(200)), Parse(2))
   Else
      Call PlayerMsg(index, "You don't have a " & Parse(4) & " equipped", 15)
    End If
    Exit Sub
    Case "gomining"
    Tool = GetPlayerWeaponSlot(index)
    If Tool = 0 Then
    Call PlayerMsg(index, "You don't have a " & Parse(4) & " equipped", 15)
    ElseIf GetPlayerInvItemNum(index, Tool) = Val(Parse(1)) Then
       Call GoMining(index, Val(Parse(3)), RandomNo(MAX_LEVEL + RandomNo(200)), Parse(2))
   Else
      Call PlayerMsg(index, "You don't have a " & Parse(4) & " equipped", 15)
    End If
    Exit Sub
    Case "goljacking"
    Tool = GetPlayerWeaponSlot(index)
    If Tool = 0 Then
    Call PlayerMsg(index, "You don't have a " & Parse(4) & " equipped", 15)
    ElseIf GetPlayerInvItemNum(index, Tool) = Val(Parse(1)) Then
       Call GoLJacking(index, Val(Parse(3)), RandomNo(MAX_LEVEL + RandomNo(200)), Parse(2))
   Else
      Call PlayerMsg(index, "You don't have a " & Parse(4) & " equipped", 15)
    End If
    Exit Sub

        Case "checkarrows"
            N = Arrows(Val(Parse(1))).Pic
            Call SendDataToMap(GetPlayerMap(index), "checkarrows" & SEP_CHAR & index & SEP_CHAR & N & SEP_CHAR & END_CHAR)
            Exit Sub

        Case "speechscript"

            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & index & "," & Parse(1)
            End If
            Exit Sub

        Case "requesteditspeech"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            Call SendDataTo(index, "SPEECHEDITOR" & SEP_CHAR & END_CHAR)
            Exit Sub

        Case "editspeech"

            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_SPEECH Then
                Call HackingAttempt(index, "Invalid Speech Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(index) & " editing speech #" & N & ".", ADMIN_LOG)
            Call SendEditSpeechTo(index, N)
            Exit Sub

        Case "savespeech"

            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_SPEECH Then
                Call HackingAttempt(index, "Invalid Speech Index")
                Exit Sub
            End If
            Speech(N).Name = Parse(2)
            Dim p As Long

            p = 3
            For i = 0 To MAX_SPEECH_OPTIONS
                Speech(N).num(i).Exit = Val(Parse(p))
                Speech(N).num(i).text = Parse(p + 1)
                Speech(N).num(i).SaidBy = Val(Parse(p + 2))
                Speech(N).num(i).Respond = Val(Parse(p + 3))
                Speech(N).num(i).Script = Val(Parse(p + 4))
                p = p + 5
                For o = 1 To 3
                    Speech(N).num(i).Responces(o).Exit = Val(Parse(p))
                    Speech(N).num(i).Responces(o).GoTo = Val(Parse(p + 1))
                    Speech(N).num(i).Responces(o).text = Parse(p + 2)
                    p = p + 3
                Next
            Next
            Call SaveSpeech(N)
            Call SendSpeechToAll(N)
            Call AddLog(GetPlayerName(index) & " saved speech #" & N & ".", ADMIN_LOG)
            Exit Sub

        Case "needspeech"
            Call SendSpeechTo(index, Val(Parse(1)))
            Exit Sub
            
        Case "hunger"
               Call HungerActive(index)
            Exit Sub

        Case "requesteditemoticon"

            ' Prevent hacking
            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            Call SendDataTo(index, "EMOTICONEDITOR" & SEP_CHAR & END_CHAR)
            Exit Sub
            
         Case "requesteditelement"
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "ELEMENTEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub

        Case "editemoticon"

            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_EMOTICONS Then
                Call HackingAttempt(index, "Invalid Emoticon Index")
                Exit Sub
            End If
            Call AddLog(GetPlayerName(index) & " editing emoticon #" & N & ".", ADMIN_LOG)
            Call SendEditEmoticonTo(index, N)
            Exit Sub
        
        Case "editelement"
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        N = Val(Parse(1))
        
        If N < 0 Or N > MAX_ELEMENTS Then
            Call HackingAttempt(index, "Invalid Element Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing element #" & N & ".", ADMIN_LOG)
        Call SendEditElementTo(index, N)
        Exit Sub

        Case "saveemoticon"

            If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(index, "Admin Cloning")
                Exit Sub
            End If
            N = Val(Parse(1))

            If N < 0 Or N > MAX_EMOTICONS Then
                Call HackingAttempt(index, "Invalid Emoticon Index")
                Exit Sub
            End If
            Emoticons(N).Type = Val(Parse(2))
            Emoticons(N).Command = Parse(3)
            Emoticons(N).Pic = Val(Parse(4))
            Emoticons(N).sound = Parse(5)
            Call SendUpdateEmoticonToAll(N)
            Call SaveEmoticon(N)
            Call AddLog(GetPlayerName(index) & " saved emoticon #" & N & ".", ADMIN_LOG)
            Exit Sub
            
        Case "saveelement"
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        N = Val(Parse(1))
        If N < 0 Or N > MAX_ELEMENTS Then
            Call HackingAttempt(index, "Invalid Element Index")
            Exit Sub
        End If

        Element(N).Name = Parse(2)
        Element(N).Strong = Val(Parse(3))
        Element(N).Weak = Val(Parse(4))
        
        Call SendUpdateElementToAll(N)
        Call SaveElement(N)
        Call AddLog(GetPlayerName(index) & " saved element #" & N & ".", ADMIN_LOG)
        Exit Sub

        Case "checkemoticons"
            Call SendDataToMap(GetPlayerMap(index), "checkemoticons" & SEP_CHAR & index & SEP_CHAR & Emoticons(Val(Parse(1))).Type & SEP_CHAR & Emoticons(Val(Parse(1))).Pic & SEP_CHAR & Emoticons(Val(Parse(1))).sound & SEP_CHAR & END_CHAR)
            Exit Sub
            
        Case "sethands"
        Player(index).Char(Player(index).CharNum).Hands = Val(Parse(1))
        Exit Sub

        Case "mapreport"
            Packs = "mapreport" & SEP_CHAR
            For i = 1 To MAX_MAPS
                Packs = Packs & Map(i).Name & SEP_CHAR
            Next
            Packs = Packs & END_CHAR
            Call SendDataTo(index, Packs)
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
            Call PlayerWarp(index, Val(Parse(1)), GetPlayerX(index), GetPlayerY(index))
            Exit Sub

        Case "warptome"
            N = FindPlayer(Parse(1))

            If N > 0 Then
                Call PlayerWarp(N, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            Else
                Call PlayerMsg(index, "Player not online!", BrightRed)
            End If
            Exit Sub

        Case "warpplayer"

            If Val(Parse(1)) > MAX_MAPS Or Val(Parse(1)) < 1 Then
                If FindPlayer(Trim$(Parse(1))) <> 0 Then
                    Call PlayerWarp(index, GetPlayerMap(FindPlayer(Trim$(Parse(1)))), GetPlayerX(FindPlayer(Trim$(Parse(1)))), GetPlayerY(FindPlayer(Trim$(Parse(1)))))

                    If Player(index).Pet.Alive = YES Then
                        Player(index).Pet.Map = GetPlayerMap(index)
                        Player(index).Pet.X = GetPlayerX(index)
                        Player(index).Pet.y = GetPlayerY(index)
                    End If
                Else
                    Call PlayerMsg(index, "'" & Parse(1) & "' is not a valid map number or an online player's name!", BrightRed)
                    Exit Sub
                End If
            Else
                Call PlayerWarp(index, Val(Parse(1)), GetPlayerX(index), GetPlayerY(index))

                If Player(index).Pet.Alive = YES Then
                    Player(index).Pet.Map = GetPlayerMap(index)
                    Player(index).Pet.X = GetPlayerX(index)
                    Player(index).Pet.y = GetPlayerY(index)
                End If
            End If
            Exit Sub

        Case "arrowhit"
            N = Val(Parse(1))
            z = Val(Parse(2))
            X = Val(Parse(3))
            y = Val(Parse(4))

            If N = TARGET_TYPE_PLAYER Then

                ' Make sure we dont try to attack ourselves
                If z <> index Then

                    ' Can we attack the player?
                    If CanAttackPlayerWithArrow(index, z) Then
                        If Not CanPlayerBlockHit(z) Then

                            ' Get the damage we can do
                            If Not CanPlayerCriticalHit(index) Then
                                Damage = GetPlayerDamage(index) - GetPlayerProtection(z) + (Rnd * 5) - 2
                                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & SEP_CHAR & END_CHAR)
                            Else
                                N = GetPlayerDamage(index)
                                Damage = N + Int(Rnd * Int(N / 2)) + 1 - GetPlayerProtection(z) + (Rnd * 5) - 2
                                Call BattleMsg(index, "You feel a surge of energy upon shooting!", BrightCyan, 0)
                                Call BattleMsg(z, GetPlayerName(index) & " shoots with amazing accuracy!", BrightCyan, 1)

                                'Call PlayerMsg(index, "You feel a surge of energy upon shooting!", BrightCyan)
                                'Call PlayerMsg(z, GetPlayerName(index) & " shoots with amazing accuracy!", BrightCyan)
                                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "Blow0" & SEP_CHAR & END_CHAR)
                            End If

                            If Damage > 0 Then
                                Call AttackPlayer(index, z, Damage)
                            Else
                                Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)
                                Call BattleMsg(z, GetPlayerName(z) & "'s attack did nothing.", BrightRed, 1)

                                'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                            End If
                        Else
                            Call BattleMsg(index, GetPlayerName(z) & " blocked your hit!", BrightCyan, 0)
                            Call BattleMsg(z, "You blocked " & GetPlayerName(index) & "'s hit!", BrightCyan, 1)

                            'Call PlayerMsg(index, GetPlayerName(z) & "'s " & Trim$(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked your hit!", BrightCyan)
                            'Call PlayerMsg(z, "Your " & Trim$(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                        Exit Sub
                    End If
                End If
            ElseIf N = TARGET_TYPE_NPC Then

                ' Can we attack the npc?
                If CanAttackNpcWithArrow(index, z) Then

                    ' Get the damage we can do
                    If Not CanPlayerCriticalHit(index) Then
                        Damage = GetPlayerDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).DEF / 2) + (Rnd * 5) - 2
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & SEP_CHAR & END_CHAR)
                    Else
                        N = GetPlayerDamage(index)
                        Damage = N + Int(Rnd * Int(N / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).DEF / 2) + (Rnd * 5) - 2
                        Call BattleMsg(index, "You feel a surge of energy upon shooting!", BrightCyan, 0)

                        'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "Blow0" & SEP_CHAR & END_CHAR)
                    End If

                    If Damage > 0 Then
                        Call AttackNpc(index, z, Damage)
                        Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                    Else
                        Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)

                        'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                        Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)

                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                    End If
                    Exit Sub
                End If
            End If
            Exit Sub
    End Select
    
    Select Case LCase$(Parse(0))
        Case "bankdeposit"
            X = GetPlayerInvItemNum(index, Val(Parse(1)))
            i = FindOpenBankSlot(index, X)
            If i = 0 Then
                Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Bank full!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
       
            If Val(Parse(2)) > GetPlayerInvItemValue(index, Val(Parse(1))) Then
                Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You cant deposit more than you have!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
       
            If GetPlayerWeaponSlot(index) = Val(Parse(1)) Or GetPlayerArmorSlot(index) = Val(Parse(1)) Or GetPlayerShieldSlot(index) = Val(Parse(1)) Or GetPlayerHelmetSlot(index) = Val(Parse(1)) Or GetPlayerLegsSlot(index) = Val(Parse(1)) Or GetPlayerBootsSlot(index) = Val(Parse(1)) Or GetPlayerGlovesSlot(index) = Val(Parse(1)) Or GetPlayerRing1Slot(index) = Val(Parse(1)) Or GetPlayerRing2Slot(index) = Val(Parse(1)) Or GetPlayerAmuletSlot(index) = Val(Parse(1)) Then
                Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You cant deposit worn equipment!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
       
            If Item(X).Type = ITEM_TYPE_CURRENCY Or Item(X).Stackable = 1 Then
                If Val(Parse(2)) <= 0 Then
                    Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You must deposit more than 0!" & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
            End If
       
            Call TakeItem(index, X, Val(Parse(2)))
            Call GiveBankItem(index, X, Val(Parse(2)), i)
       
            Call SendBank(index)
            Exit Sub
   
        Case "bankwithdraw"
            i = GetPlayerBankItemNum(index, Val(Parse(1)))
            TempVal = Val(Parse(2))
            X = FindOpenInvSlot(index, i)
            If X = 0 Then
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

                If Trim(LCase(Item(GetPlayerInvItemNum(index, X)).Name)) <> "gold" Then
                    If GetPlayerInvItemValue(index, X) + Val(Parse(2)) > 100 Then
                        TempVal = 100 - GetPlayerInvItemValue(index, X)
                    End If
                End If
            End If
               
            Call GiveItem(index, i, TempVal)
            Call TakeBankItem(index, i, TempVal)
       
            Call SendBank(index)
            Exit Sub
    End Select
    
    Call HackingAttempt(index, "Invalid packet. (" & Parse(0) & ")")
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
            Packet = Mid$(Player(index).Buffer, 1, Start - 1)
            Player(index).Buffer = Mid$(Player(index).Buffer, Start + 1, Len(Player(index).Buffer))
            Player(index).DataPackets = Player(index).DataPackets + 1
            Start = InStr(Player(index).Buffer, END_CHAR)

            If Len(Packet) > 0 Then
                Call HandleData(index, Packet)
            End If

        Loop

        ' Not useful
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

Function IsConnected(ByVal index As Long) As Boolean

    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If

End Function

Function IsLoggedIn(ByVal index As Long) As Boolean

    If IsConnected(index) And Trim$(Player(index).Login) <> "" Then
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
    Next
End Function

Function IsPlaying(ByVal index As Long) As Boolean

    If index <= 0 Or index > MAX_PLAYERS Then
        IsPlaying = False
        Exit Function
    End If

    If IsConnected(index) And Player(index).InGame = True Then
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

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal index As Long)
Dim Packet As String

    Packet = "MAPMSG2" & SEP_CHAR & Msg & SEP_CHAR & index & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub PlainMsg(ByVal index As Long, ByVal Msg As String, ByVal num As Long)
Dim Packet As String

    Packet = "PLAINMSG" & SEP_CHAR & Msg & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendArrows(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_ARROWS
        Call SendUpdateArrowTo(index, i)
    Next
End Sub

Sub SendChars(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "ALLCHARS" & SEP_CHAR
    For i = 1 To MAX_CHARS
        Packet = Packet & Trim(Player(index).Char(i).Name) & SEP_CHAR & Trim(Class(Player(index).Char(i).Class).Name) & SEP_CHAR & Player(index).Char(i).Level & SEP_CHAR & Player(index).Char(i).Sprite & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendClasses(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 1 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendDataTo(ByVal index As Long, ByVal Data As String)

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
    Next
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> index Then
            Call SendDataTo(i, Data)
        End If
    Next
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next
End Sub

Sub SendEditArrowTo(ByVal index As Long, ByVal EmoNum As Long)
Dim Packet As String

    Packet = "EDITArrow" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Arrows(EmoNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditEmoticonTo(ByVal index As Long, ByVal EmoNum As Long)
Dim Packet As String

    Packet = "EDITEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Emoticons(EmoNum).Type & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & Emoticons(EmoNum).sound & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateElementToAll(ByVal ElementNum As Long)
Dim Packet As String

    Packet = "UPDATEELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & SEP_CHAR & END_CHAR
 Call SendDataToAll(Packet)
End Sub

Sub SendUpdateElementTo(ByVal index As Long, ByVal ElementNum As Long)
Dim Packet As String

    Packet = "UPDATEELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & SEP_CHAR & END_CHAR
  Call SendDataTo(index, Packet)
End Sub

Sub SendEditElementTo(ByVal index As Long, ByVal ElementNum As Long)
Dim Packet As String

    Packet = "EDITELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & SEP_CHAR & END_CHAR
 Call SendDataTo(index, Packet)
End Sub

Sub SendEditItemTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound & SEP_CHAR & Item(ItemNum).LevelReq & SEP_CHAR & Item(ItemNum).Element & SEP_CHAR & Item(ItemNum).StamRemove & SEP_CHAR & Item(ItemNum).Rarity & SEP_CHAR & Item(ItemNum).BowsReq & SEP_CHAR & Item(ItemNum).LargeBladesReq & SEP_CHAR & Item(ItemNum).SmallBladesReq & SEP_CHAR & Item(ItemNum).BluntWeaponsReq & SEP_CHAR & Item(ItemNum).PoleArmsReq & SEP_CHAR & Item(ItemNum).AxesReq & SEP_CHAR & Item(ItemNum).ThrownReq & SEP_CHAR & Item(ItemNum).XbowsReq & SEP_CHAR & Item(ItemNum).LBA & SEP_CHAR & Item(ItemNum).SBA & SEP_CHAR & Item(ItemNum).BWA
    Packet = Packet & SEP_CHAR & Item(ItemNum).PAA & Item(ItemNum).AA & SEP_CHAR & Item(ItemNum).TWA & SEP_CHAR & Item(ItemNum).XBA & SEP_CHAR & Item(ItemNum).BA & SEP_CHAR & Item(ItemNum).Poison & SEP_CHAR & Item(ItemNum).Disease
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditNpcTo(ByVal index As Long, ByVal npcnum As Long)
Dim Packet As String
Dim i As Long

    'Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR
    'Packet = Packet & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).str & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & END_CHAR
    Packet = "EDITNPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(Npc(npcnum).Name) & SEP_CHAR & Trim$(Npc(npcnum).AttackSay) & SEP_CHAR & Npc(npcnum).Sprite & SEP_CHAR & Npc(npcnum).SpawnSecs & SEP_CHAR & Npc(npcnum).Behavior & SEP_CHAR & Npc(npcnum).Range & SEP_CHAR & Npc(npcnum).STR & SEP_CHAR & Npc(npcnum).DEF & SEP_CHAR & Npc(npcnum).Speed & SEP_CHAR & Npc(npcnum).Magi & SEP_CHAR & Npc(npcnum).Big & SEP_CHAR & Npc(npcnum).MaxHp & SEP_CHAR & Npc(npcnum).Exp & SEP_CHAR & Npc(npcnum).SpawnTime & SEP_CHAR & Npc(npcnum).Speech & SEP_CHAR & Npc(npcnum).Element & SEP_CHAR & Npc(npcnum).Poison & SEP_CHAR & Npc(npcnum).AP & SEP_CHAR & Npc(npcnum).Disease & SEP_CHAR & Npc(npcnum).Quest & SEP_CHAR
    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(npcnum).ItemNPC(i).Chance
        Packet = Packet & SEP_CHAR & Npc(npcnum).ItemNPC(i).ItemNum
        Packet = Packet & SEP_CHAR & Npc(npcnum).ItemNPC(i).ItemValue & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditShopTo(ByVal index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long, z As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To 6
        For z = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(i).Value(z).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GetValue & SEP_CHAR
        Next
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditSpeechTo(ByVal index As Long, ByVal SpcNum As Long)
Dim Packet As String
Dim i, o As Long

    Packet = "EDITSPEECH" & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR
    For i = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(i).Exit & SEP_CHAR & Speech(SpcNum).num(i).text & SEP_CHAR & Speech(SpcNum).num(i).SaidBy & SEP_CHAR & Speech(SpcNum).num(i).Respond & SEP_CHAR & Speech(SpcNum).num(i).Script & SEP_CHAR
        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(i).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).text & SEP_CHAR
        Next
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditSpellTo(ByVal index As Long, ByVal spellnum As Long)
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & spellnum & SEP_CHAR & Trim$(Spell(spellnum).Name) & SEP_CHAR & Spell(spellnum).ClassReq & SEP_CHAR & Spell(spellnum).LevelReq & SEP_CHAR & Spell(spellnum).Type & SEP_CHAR & Spell(spellnum).Data1 & SEP_CHAR & Spell(spellnum).Data2 & SEP_CHAR & Spell(spellnum).Data3 & SEP_CHAR & Spell(spellnum).MPCost & SEP_CHAR & Spell(spellnum).sound & SEP_CHAR & Spell(spellnum).Range & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & Spell(spellnum).AE & SEP_CHAR & Spell(spellnum).Pic & SEP_CHAR & Spell(spellnum).Element & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEmoticons(ByVal index As Long)
Dim i As Long

    For i = 0 To MAX_EMOTICONS

        If Trim$(Emoticons(i).Command) <> "" Then
            Call SendUpdateEmoticonTo(index, i)
        End If
    Next
End Sub

Sub SendFriendListTo(ByVal index As Long)
Dim Packet As String
Dim N As Long

    Packet = "FRIENDLIST" & SEP_CHAR
    For N = 1 To MAX_FRIENDS

        If FindPlayer(Player(index).Char(Player(index).CharNum).Friends(N)) And Player(index).Char(Player(index).CharNum).Friends(N) <> "" Then
            Packet = Packet & Player(index).Char(Player(index).CharNum).Friends(N) & SEP_CHAR
        End If
    Next
    Packet = Packet & NEXT_CHAR & SEP_CHAR
    For N = 1 To MAX_FRIENDS
        Packet = Packet & Player(index).Char(Player(index).CharNum).Friends(N) & SEP_CHAR
    Next
    Packet = Packet & NEXT_CHAR & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendFriendListToNeeded(ByVal Name As String)
Dim i, o As Long

    For i = i To MAX_PLAYERS

        If IsPlaying(i) Then
            For o = 1 To MAX_FRIENDS

                If Trim$(Player(i).Char(Player(i).CharNum).Friends(i)) = Name Then
                    Call SendFriendListTo(i)
                End If
            Next
        End If
    Next
End Sub

Sub SendHP(ByVal index As Long)
Dim Packet As String

    Packet = "playerhp" & SEP_CHAR & index & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
    Packet = "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendFP(ByVal index As Long)
Dim Packet As String

    Packet = "playerfp" & SEP_CHAR & index & SEP_CHAR & GetPlayerMaxFP(index) & SEP_CHAR & GetPlayerFP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendInfo(ByVal index As Long)
Dim Packet As String

    Packet = "INFO" & SEP_CHAR & TotalOnlinePlayers & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendBank(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERBANK" & SEP_CHAR
    For i = 1 To MAX_BANK
        Packet = Packet & GetPlayerBankItemNum(index, i) & SEP_CHAR & GetPlayerBankItemValue(index, i) & SEP_CHAR & GetPlayerBankItemDur(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
   
    Call SendDataTo(index, Packet)
End Sub

Sub SendBankUpdate(ByVal index As Long, ByVal BankSlot As Long)
Dim Packet As String
   
    Packet = "PLAYERBANKUPDATE" & SEP_CHAR & BankSlot & SEP_CHAR & GetPlayerBankItemNum(index, BankSlot) & SEP_CHAR & GetPlayerBankItemValue(index, BankSlot) & SEP_CHAR & GetPlayerBankItemDur(index, BankSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendInventory(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR & index & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
Dim Packet As String

    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & index & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & index & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendItems(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_ITEMS

        If Trim$(Item(i).Name) <> "" Then
            Call SendUpdateItemTo(index, i)
        End If
    Next
End Sub

Sub SendElements(ByVal index As Long)
Dim Packet As String
Dim i As Long

    For i = 0 To MAX_ELEMENTS
            Call SendUpdateElementTo(index, i)
    Next i
End Sub

Sub SendJoinMap(ByVal index As Long)
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
            Packet = Packet & GetPlayerAlignment(i) & SEP_CHAR
            Packet = Packet & END_CHAR
            Call SendDataTo(index, Packet)

            If Player(i).Pet.Alive = YES Then
                Packet = "PETDATA" & SEP_CHAR
                Packet = Packet & i & SEP_CHAR
                Packet = Packet & Player(i).Pet.Alive & SEP_CHAR
                Packet = Packet & Player(i).Pet.Map & SEP_CHAR
                Packet = Packet & Player(i).Pet.X & SEP_CHAR
                Packet = Packet & Player(i).Pet.y & SEP_CHAR
                Packet = Packet & Player(i).Pet.Dir & SEP_CHAR
                Packet = Packet & Player(i).Pet.Sprite & SEP_CHAR
                Packet = Packet & Player(i).Pet.HP & SEP_CHAR
                Packet = Packet & Player(i).Pet.Level * 5 & SEP_CHAR
                Packet = Packet & END_CHAR
                Call SendDataTo(index, Packet)
            End If
        End If
    Next

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
    Packet = Packet & GetPlayerAlignment(index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)

    If Player(index).Pet.Alive = YES Then
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & index & SEP_CHAR
        Packet = Packet & Player(index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(index).Pet.X & SEP_CHAR
        Packet = Packet & Player(index).Pet.y & SEP_CHAR
        Packet = Packet & Player(index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(index).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(index).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataToMap(GetPlayerMap(index), Packet)
    End If
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
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
    Packet = Packet & GetPlayerAlignment(index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMapBut(index, MapNum, Packet)

    If Player(index).Pet.Alive = YES Then
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & index & SEP_CHAR
        Packet = Packet & Player(index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(index).Pet.X & SEP_CHAR
        Packet = Packet & Player(index).Pet.y & SEP_CHAR
        Packet = Packet & Player(index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(index).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(index).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataToMapBut(index, MapNum, Packet)
    End If
End Sub

Sub SendLeftGame(ByVal index As Long)
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
    Packet = Packet & END_CHAR
    Call SendDataToAllBut(index, Packet)
    Packet = "PETDATA" & SEP_CHAR
    Packet = Packet & index & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToAllBut(index, Packet)
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim X As Long
Dim y As Long
Dim i As Long
Dim o As Long
Dim p1 As String, p2 As String

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR
    For y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX

            With Map(MapNum).Tile(X, y)
                i = 0
                o = 0

                If .Ground <> 0 Then i = 0
                If .GroundSet <> -1 Then i = 1
                If .Mask <> 0 Then i = 2
                If .MaskSet <> -1 Then i = 3
                If .Anim <> 0 Then i = 4
                If .AnimSet <> -1 Then i = 5
                If .Fringe <> 0 Then i = 6
                If .FringeSet <> -1 Then i = 7
                If .Type <> 0 Then i = 8
                If .Data1 <> 0 Then i = 9
                If .Data2 <> 0 Then i = 10
                If .Data3 <> 0 Then i = 11
                If .String1 <> "" Then i = 12
                If .String2 <> "" Then i = 13
                If .String3 <> "" Then i = 14
                If .Mask2 <> 0 Then i = 15
                If .Mask2Set <> -1 Then i = 16
                If .M2Anim <> 0 Then i = 17
                If .M2AnimSet <> -1 Then i = 18
                If .FAnim <> 0 Then i = 19
                If .FAnimSet <> -1 Then i = 20
                If .Fringe2 <> 0 Then i = 21
                If .Fringe2Set <> -1 Then i = 22
                If .Light <> 0 Then i = 23
                If .F2Anim <> 0 Then i = 24
                If .F2AnimSet <> -1 Then i = 25
                Packet = Packet & .Ground & SEP_CHAR

                If o < i Then
                    o = o + 1
                    Packet = Packet & .GroundSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Mask & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .MaskSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Anim & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .AnimSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Fringe & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .FringeSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Type & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Data1 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Data2 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Data3 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .String1 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .String2 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .String3 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Mask2 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Mask2Set & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .M2Anim & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .M2AnimSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .FAnim & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .FAnimSet & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Fringe2 & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Fringe2Set & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .Light & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .F2Anim & SEP_CHAR
                End If

                If o < i Then
                    o = o + 1
                    Packet = Packet & .F2AnimSet & SEP_CHAR
                End If
                Packet = Packet & NEXT_CHAR & SEP_CHAR
            End With

        Next
    Next
    For X = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(MapNum).Npc(X) & SEP_CHAR
        Packet = Packet & Map(MapNum).NpcSpawn(X).Used & SEP_CHAR & Map(MapNum).NpcSpawn(X).X & SEP_CHAR & Map(MapNum).NpcSpawn(X).y & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    X = Int(Len(Packet) / 2)
    p1 = Mid$(Packet, 1, X)
    p2 = Mid$(Packet, X + 1, Len(Packet) - X)
    Call SendDataTo(index, Packet)
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS

        If MapNum > 0 Then
            Packet = Packet & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
        End If
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS

        If MapNum > 0 Then
            Packet = Packet & MapNpc(MapNum, i).num & SEP_CHAR & MapNpc(MapNum, i).X & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
        End If
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendMP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendNewCharClasses(ByVal index As Long)
Dim i As Long
Dim Packet As String

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 1 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).MaleSprite & SEP_CHAR & Class(i).FemaleSprite & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendNpcs(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_NPCS

        If Trim$(Npc(i).Name) <> "" Then
            Call SendUpdateNpcTo(index, i)
        End If
    Next
End Sub

Sub SendOnlineList()
Dim Packet As String
Dim i As Long
Dim N As Long

    Packet = ""
    N = 0
    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Packet = Packet & SEP_CHAR & GetPlayerName(i) & SEP_CHAR
            N = N + 1
        End If
    Next
    Packet = "ONLINELIST" & SEP_CHAR & N & Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendPlayerData(ByVal index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on the map
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
    Packet = Packet & GetPlayerAlignment(index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)

    If Player(index).Pet.Alive = YES Then
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & index & SEP_CHAR
        Packet = Packet & Player(index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(index).Pet.X & SEP_CHAR
        Packet = Packet & Player(index).Pet.y & SEP_CHAR
        Packet = Packet & Player(index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(index).Pet.Sprite & SEP_CHAR
        Packet = Packet & Player(index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(index).Pet.Level * 5 & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataToMap(GetPlayerMap(index), Packet)
    End If
End Sub

Sub SendPlayerSpells(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "SPELLS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(index, i) & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendPlayerXY(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendShops(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SHOPS

        If Trim$(Shop(i).Name) <> "" Then
            Call SendUpdateShopTo(index, i)
        End If
    Next
End Sub

Sub SendSP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendSpeech(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SPEECH

        If Trim$(Speech(i).Name) <> "" Then
            Call SendSpeechTo(index, i)
        End If
    Next
End Sub

Sub SendSpeechTo(ByVal index As Long, ByVal SpcNum As Long)
Dim Packet As String
Dim i, o As Long

    Packet = "SPEECH" & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR
    For i = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(i).Exit & SEP_CHAR & Speech(SpcNum).num(i).text & SEP_CHAR & Speech(SpcNum).num(i).SaidBy & SEP_CHAR & Speech(SpcNum).num(i).Respond & SEP_CHAR & Speech(SpcNum).num(i).Script & SEP_CHAR
        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(i).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).text & SEP_CHAR
        Next
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendSpeechToAll(ByVal SpcNum As Long)
Dim Packet As String
Dim i, o As Long

    Packet = "SPEECH" & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR
    For i = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(i).Exit & SEP_CHAR & Speech(SpcNum).num(i).text & SEP_CHAR & Speech(SpcNum).num(i).SaidBy & SEP_CHAR & Speech(SpcNum).num(i).Respond & SEP_CHAR & Speech(SpcNum).num(i).Script & SEP_CHAR
        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(i).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).text & SEP_CHAR
        Next
    Next
    Packet = Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendSpells(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS

        If Trim$(Spell(i).Name) <> "" Then
            Call SendUpdateSpellTo(index, i)
        End If
    Next
End Sub

Sub SendStats(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerstr(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & GetPlayerNextLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & GetPlayerNextLargeBladesLevel(index) & SEP_CHAR & GetPlayerLargeBladesExp(index) & SEP_CHAR & GetPlayerLargeBladesLevel(index) & SEP_CHAR & GetPlayerNextSmallBladesLevel(index) & SEP_CHAR & GetPlayerSmallBladesExp(index) & SEP_CHAR & GetPlayerSmallBladesLevel(index) & SEP_CHAR & GetPlayerNextBluntWeaponsLevel(index) & SEP_CHAR & GetPlayerBluntWeaponsExp(index) & SEP_CHAR & GetPlayerBluntWeaponsLevel(index) & SEP_CHAR & GetPlayerNextPolesLevel(index) & SEP_CHAR & GetPlayerPolesExp(index) & SEP_CHAR & GetPlayerPolesLevel(index) & SEP_CHAR & GetPlayerNextAxesLevel(index) & SEP_CHAR & GetPlayerAxesExp(index) & SEP_CHAR & GetPlayerAxesLevel(index) & SEP_CHAR
    Packet = Packet & GetPlayerNextThrownLevel(index) & SEP_CHAR & GetPlayerThrownExp(index) & SEP_CHAR & GetPlayerThrownLevel(index) & SEP_CHAR & GetPlayerNextXbowsLevel(index) & SEP_CHAR & GetPlayerXbowsExp(index) & SEP_CHAR & GetPlayerXbowsLevel(index) & SEP_CHAR & GetPlayerNextBowsLevel(index) & SEP_CHAR & GetPlayerBowsExp(index) & SEP_CHAR & GetPlayerBowsLevel(index) & SEP_CHAR & GetPlayerNextFishLevel(index) & SEP_CHAR & GetPlayerFishExp(index) & SEP_CHAR & GetPlayerFishLevel(index) & SEP_CHAR & GetPlayerNextMineLevel(index) & SEP_CHAR & GetPlayerMineExp(index) & SEP_CHAR & GetPlayerMineLevel(index) & SEP_CHAR & GetPlayerNextLJackingLevel(index) & SEP_CHAR & GetPlayerLJackingExp(index) & SEP_CHAR & GetPlayerLJackingLevel(index) & SEP_CHAR & GetPlayerArrowsAmount(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
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
    Next
    Call SpawnAllMapNpcs
End Sub

Sub SendTrade(ByVal index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long, X As Long, y As Long, z As Long, XX As Long

    z = 0
    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To 6
        For XX = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(i).Value(XX).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GetValue & SEP_CHAR

            ' Item #
            X = Shop(ShopNum).TradeItem(i).Value(XX).GetItem

            If Item(X).Type = ITEM_TYPE_SPELL Then

                ' Spell class requirement
                y = Spell(Item(X).Data1).ClassReq

                If y = 0 Then
                    Call PlayerMsg(index, Trim$(Item(X).Name) & " can be used by all classes.", Yellow)
                Else
                    Call PlayerMsg(index, Trim$(Item(X).Name) & " can only be used by a " & GetClassName(y) & ".", Yellow)
                End If
            End If

            If X < 1 Then
                z = z + 1
            End If
        Next
    Next
    Packet = Packet & END_CHAR

    If z = (MAX_TRADES * 6) Then
        Call PlayerMsg(index, "This shop has nothing to sell!", BrightRed)
    Else
        Call SendDataTo(index, Packet)
    End If
End Sub

Sub SendUpdateArrowTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEArrow" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEARROW" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateEmoticonTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Emoticons(ItemNum).Type & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & Emoticons(ItemNum).sound & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Emoticons(ItemNum).Type & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & Emoticons(ItemNum).sound & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound & SEP_CHAR & Item(ItemNum).LevelReq & SEP_CHAR & Item(ItemNum).Element & SEP_CHAR & Item(ItemNum).StamRemove & SEP_CHAR & Item(ItemNum).Rarity & SEP_CHAR & Item(ItemNum).BowsReq & SEP_CHAR & Item(ItemNum).LargeBladesReq & SEP_CHAR & Item(ItemNum).SmallBladesReq & SEP_CHAR & Item(ItemNum).BluntWeaponsReq & SEP_CHAR & Item(ItemNum).PoleArmsReq & SEP_CHAR & Item(ItemNum).AxesReq & SEP_CHAR & Item(ItemNum).ThrownReq & SEP_CHAR & Item(ItemNum).XbowsReq & SEP_CHAR & Item(ItemNum).LBA & SEP_CHAR & Item(ItemNum).SBA & SEP_CHAR & Item(ItemNum).BWA
    Packet = Packet & SEP_CHAR & Item(ItemNum).PAA & Item(ItemNum).AA & SEP_CHAR & Item(ItemNum).TWA & SEP_CHAR & Item(ItemNum).XBA & SEP_CHAR & Item(ItemNum).BA & SEP_CHAR & Item(ItemNum).Poison & SEP_CHAR & Item(ItemNum).Disease
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound & SEP_CHAR & Item(ItemNum).LevelReq & SEP_CHAR & Item(ItemNum).Element & SEP_CHAR & Item(ItemNum).StamRemove & SEP_CHAR & Item(ItemNum).Rarity & SEP_CHAR & Item(ItemNum).BowsReq & SEP_CHAR & Item(ItemNum).LargeBladesReq & SEP_CHAR & Item(ItemNum).SmallBladesReq & SEP_CHAR & Item(ItemNum).BluntWeaponsReq & SEP_CHAR & Item(ItemNum).PoleArmsReq & SEP_CHAR & Item(ItemNum).AxesReq & SEP_CHAR & Item(ItemNum).ThrownReq & SEP_CHAR & Item(ItemNum).XbowsReq & SEP_CHAR & Item(ItemNum).LBA & SEP_CHAR & Item(ItemNum).SBA & SEP_CHAR & Item(ItemNum).BWA
    Packet = Packet & SEP_CHAR & Item(ItemNum).PAA & Item(ItemNum).AA & SEP_CHAR & Item(ItemNum).TWA & SEP_CHAR & Item(ItemNum).XBA & SEP_CHAR & Item(ItemNum).BA & SEP_CHAR & Item(ItemNum).Poison & SEP_CHAR & Item(ItemNum).Disease
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal npcnum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(Npc(npcnum).Name) & SEP_CHAR & Trim$(Npc(npcnum).AttackSay) & SEP_CHAR & Npc(npcnum).Sprite & SEP_CHAR & Npc(npcnum).SpawnSecs & SEP_CHAR & Npc(npcnum).Behavior & SEP_CHAR & Npc(npcnum).Range & SEP_CHAR & Npc(npcnum).STR & SEP_CHAR & Npc(npcnum).DEF & SEP_CHAR & Npc(npcnum).Speed & SEP_CHAR & Npc(npcnum).Magi & SEP_CHAR & Npc(npcnum).Big & SEP_CHAR & Npc(npcnum).MaxHp & SEP_CHAR & Npc(npcnum).Exp & SEP_CHAR & Npc(npcnum).SpawnTime & SEP_CHAR & Npc(npcnum).Speech & SEP_CHAR & Npc(npcnum).Element & SEP_CHAR & Npc(npcnum).Poison & SEP_CHAR & Npc(npcnum).AP & SEP_CHAR & Npc(npcnum).Disease & SEP_CHAR & Npc(npcnum).Quest & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal npcnum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(Npc(npcnum).Name) & SEP_CHAR & Trim$(Npc(npcnum).AttackSay) & SEP_CHAR & Npc(npcnum).Sprite & SEP_CHAR & Npc(npcnum).SpawnSecs & SEP_CHAR & Npc(npcnum).Behavior & SEP_CHAR & Npc(npcnum).Range & SEP_CHAR & Npc(npcnum).STR & SEP_CHAR & Npc(npcnum).DEF & SEP_CHAR & Npc(npcnum).Speed & SEP_CHAR & Npc(npcnum).Magi & SEP_CHAR & Npc(npcnum).Big & SEP_CHAR & Npc(npcnum).MaxHp & SEP_CHAR & Npc(npcnum).Exp & SEP_CHAR & Npc(npcnum).SpawnTime & SEP_CHAR & Npc(npcnum).Speech & SEP_CHAR & Npc(npcnum).Element & SEP_CHAR & Npc(npcnum).Poison & SEP_CHAR & Npc(npcnum).AP & SEP_CHAR & Npc(npcnum).Disease & SEP_CHAR & Npc(npcnum).Quest & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal spellnum As Long)
    Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & spellnum & SEP_CHAR & Trim$(Spell(spellnum).Name) & SEP_CHAR & Spell(spellnum).ClassReq & SEP_CHAR & Spell(spellnum).LevelReq & SEP_CHAR & Spell(spellnum).Type & SEP_CHAR & Spell(spellnum).Data1 & SEP_CHAR & Spell(spellnum).Data2 & SEP_CHAR & Spell(spellnum).Data3 & SEP_CHAR & Spell(spellnum).MPCost & SEP_CHAR & Trim$(Spell(spellnum).sound) & SEP_CHAR & Spell(spellnum).Range & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & Spell(spellnum).AE & SEP_CHAR & Spell(spellnum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateSpellToAll(ByVal spellnum As Long)
    Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & spellnum & SEP_CHAR & Trim$(Spell(spellnum).Name) & SEP_CHAR & Spell(spellnum).ClassReq & SEP_CHAR & Spell(spellnum).LevelReq & SEP_CHAR & Spell(spellnum).Type & SEP_CHAR & Spell(spellnum).Data1 & SEP_CHAR & Spell(spellnum).Data2 & SEP_CHAR & Spell(spellnum).Data3 & SEP_CHAR & Spell(spellnum).MPCost & SEP_CHAR & Trim$(Spell(spellnum).sound) & SEP_CHAR & Spell(spellnum).Range & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & Spell(spellnum).AE & SEP_CHAR & Spell(spellnum).Pic & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendWeatherTo(ByVal index As Long)
Dim Packet As String

    If RainIntensity <= 0 Then RainIntensity = 1
    Packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & RainIntensity & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendWeatherToAll()
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
    Next
End Sub

Sub SendWhosOnline(ByVal index As Long)
Dim s As String
Dim N As Long, i As Long

    s = ""
    N = 0
    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> index Then
            s = s & GetPlayerName(i) & ", "
            N = N + 1
        End If
    Next

    If N = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & N & " other players online: " & s & "."
    End If
    Call PlayerMsg(index, s, WhoColor)
End Sub

Sub SendInvSlots(ByVal index As Long)
    Dim Packet As String

    If IsPlaying(index) Then
        Packet = "PLAYERINVSLOTS" & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & GetPlayerLegsSlot(index) & SEP_CHAR & GetPlayerBootsSlot(index) & SEP_CHAR & GetPlayerGlovesSlot(index) & SEP_CHAR & GetPlayerRing1Slot(index) & SEP_CHAR & GetPlayerRing2Slot(index) & SEP_CHAR & GetPlayerAmuletSlot(index) & SEP_CHAR & END_CHAR
        Call SendDataTo(index, Packet)
    End If

End Sub

Sub SendWornEquipment(ByVal index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(index) = GetPlayerMap(i) And i <> index Then
                Packet = "PLAYERWORNEQ" & SEP_CHAR & i & SEP_CHAR
                If GetPlayerArmorSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerArmorSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerWeaponSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerWeaponSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerHelmetSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerHelmetSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerShieldSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerShieldSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                 If GetPlayerLegsSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerLegsSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerBootsSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerBootsSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerGlovesSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerGlovesSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerRing1Slot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerRing1Slot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerRing2Slot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerRing2Slot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerAmuletSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerAmuletSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                Packet = Packet & END_CHAR
                Call SendDataTo(index, Packet)
                
                Packet = "PLAYERWORNEQ" & SEP_CHAR & index & SEP_CHAR
                If GetPlayerArmorSlot(index) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(index, GetPlayerArmorSlot(index)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerWeaponSlot(index) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerHelmetSlot(index) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerShieldSlot(index) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(index, GetPlayerShieldSlot(index)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerLegsSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerLegsSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerBootsSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerBootsSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerGlovesSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerGlovesSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerRing1Slot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerRing1Slot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerRing2Slot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerRing2Slot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                If GetPlayerAmuletSlot(i) > 0 Then
                    Packet = Packet & GetPlayerInvItemNum(i, GetPlayerAmuletSlot(i)) & SEP_CHAR
                Else
                    Packet = Packet & 0 & SEP_CHAR
                End If
                Packet = Packet & END_CHAR
                Call SendDataTo(i, Packet)
            End If
        End If
    Next
End Sub


Sub SocketConnected(ByVal index As Long)

    If index <> 0 Then

        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
        If Not IsBanned(GetPlayerIP(index)) Then
            Call TextAdd(frmServer.txtText(0), "Received connection from " & GetPlayerIP(index) & ".", True)
        Else
            Call AlertMsg(index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
        End If

        'Else
        ' Tried multiple connections
        '    Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
        'End If
    End If
End Sub

Sub UpdateCaption()
    frmServer.Caption = GAME_NAME & " - Server - Powered By Elysium Source"
    frmServer.lblIP.Caption = "Ip Address: " & frmServer.Socket(0).LocalIP
    frmServer.lblPort.Caption = "Port: " & STR(frmServer.Socket(0).LocalPort)
    frmServer.TPO.Caption = "Total Players Online: " & TotalOnlinePlayers
    Exit Sub
End Sub

' SAFE MODE -- Uncomment for ON, comment for OFF (whole function)
'Function Parse(ByVal index As Long) As String
'    If index > NumParse Then
'        Call HackingAttempt(ParseIndex, "Subscript out of range, " & ZePacket(0))
'        Exit Function
'    End If
'
'    Parse = ZePacket(index)
'End Function

Sub SendGameClockTo(ByVal index As Long)
Dim Packet As String

    Packet = "GAMECLOCK" & SEP_CHAR & GameClock & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendGameClockToAll()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendGameClockTo(i)
        End If
    Next i
End Sub

Sub DisabledTime()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call DisabledTimeTo(i)
        End If
    Next i
    
End Sub

Sub DisabledTimeTo(ByVal index As Long)
Dim Packet As String

    Packet = "DTIME" & SEP_CHAR & TimeDisable & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendNewsTo(ByVal index As Long)
Dim Packet As String

    Packet = "NEWS" & SEP_CHAR & ReadINI("DATA", "ServerNews", App.Path & "\News.ini") & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub QuestMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "QUESTMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendQuest(ByVal index As Long)
'Dim Packet As String
Dim i As Long

    For i = 1 To MAX_QUESTS
        If Trim(Quest(i).Name) <> "" Then
            Call SendUpdateQuestTo(index, i)
        End If
    Next i
End Sub

Sub SendUpdateQuestToAll(ByVal questnum As Long)
Dim Packet As String

    Packet = "UPDATEQUEST" & SEP_CHAR & questnum & SEP_CHAR & Trim(Quest(questnum).Name) & SEP_CHAR & Trim(Quest(questnum).After) & SEP_CHAR & Trim(Quest(questnum).Before) & SEP_CHAR & Quest(questnum).ClassIsReq & SEP_CHAR & Quest(questnum).ClassReq & SEP_CHAR & Trim(Quest(questnum).During) & SEP_CHAR & Trim(Quest(questnum).End) & SEP_CHAR & Quest(questnum).ItemReq & SEP_CHAR & Quest(questnum).ItemVal & SEP_CHAR & Quest(questnum).LevelIsReq & SEP_CHAR & Quest(questnum).LevelReq & SEP_CHAR & Trim(Quest(questnum).NotHasItem) & SEP_CHAR & Quest(questnum).RewardNum & SEP_CHAR & Quest(questnum).RewardVal & SEP_CHAR & Trim(Quest(questnum).Start) & SEP_CHAR & Quest(questnum).StartItem & SEP_CHAR & Quest(questnum).StartOn & SEP_CHAR & Quest(questnum).Startval & SEP_CHAR & Quest(questnum).QuestExpReward & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateQuestTo(ByVal index As Long, ByVal questnum As Long)
Dim Packet As String

    Packet = "UPDATEQUEST" & SEP_CHAR & questnum & SEP_CHAR & Trim(Quest(questnum).Name) & SEP_CHAR & Trim(Quest(questnum).After) & SEP_CHAR & Trim(Quest(questnum).Before) & SEP_CHAR & Quest(questnum).ClassIsReq & SEP_CHAR & Quest(questnum).ClassReq & SEP_CHAR & Trim(Quest(questnum).During) & SEP_CHAR & Trim(Quest(questnum).End) & SEP_CHAR & Quest(questnum).ItemReq & SEP_CHAR & Quest(questnum).ItemVal & SEP_CHAR & Quest(questnum).LevelIsReq & SEP_CHAR & Quest(questnum).LevelReq & SEP_CHAR & Trim(Quest(questnum).NotHasItem) & SEP_CHAR & Quest(questnum).RewardNum & SEP_CHAR & Quest(questnum).RewardVal & SEP_CHAR & Trim(Quest(questnum).Start) & SEP_CHAR & Quest(questnum).StartItem & SEP_CHAR & Quest(questnum).StartOn & SEP_CHAR & Quest(questnum).Startval & SEP_CHAR & Quest(questnum).QuestExpReward & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditQuestTo(ByVal index As Long, ByVal questnum As Long)
Dim Packet As String

    Packet = "EDITQUEST" & SEP_CHAR & questnum & SEP_CHAR & Trim(Quest(questnum).Name) & SEP_CHAR & Trim(Quest(questnum).After) & SEP_CHAR & Trim(Quest(questnum).Before) & SEP_CHAR & Quest(questnum).ClassIsReq & SEP_CHAR & Quest(questnum).ClassReq & SEP_CHAR & Trim(Quest(questnum).During) & SEP_CHAR & Trim(Quest(questnum).End) & SEP_CHAR & Quest(questnum).ItemReq & SEP_CHAR & Quest(questnum).ItemVal & SEP_CHAR & Quest(questnum).LevelIsReq & SEP_CHAR & Quest(questnum).LevelReq & SEP_CHAR & Trim(Quest(questnum).NotHasItem) & SEP_CHAR & Quest(questnum).RewardNum & SEP_CHAR & Quest(questnum).RewardVal & SEP_CHAR & Trim(Quest(questnum).Start) & SEP_CHAR & Quest(questnum).StartItem & SEP_CHAR & Quest(questnum).StartOn & SEP_CHAR & Quest(questnum).Startval & SEP_CHAR & Quest(questnum).QuestExpReward & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub
