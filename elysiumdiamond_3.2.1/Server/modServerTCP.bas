Attribute VB_Name = "modServerTCP"

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.
Option Explicit

'Dim ZePacket() As String ' SAFE MODE -- Uncomment for ON, comment for OFF
'Dim NumParse As Long ' SAFE MODE -- Uncomment for ON, comment for OFF
'Dim ParseIndex As Long ' SAFE MODE -- Uncomment for ON, comment for OFF
Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (Index = 0) Then
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

    Packet = ADMINMSG_CHAR & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, Packet)
        End If

    Next

End Sub

Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Dim Packet As String

    Packet = ALERTMSG_CHAR & SEP_CHAR & Msg & END_CHAR
    Call SendDataTo(Index, Packet)
    Call CloseSocket(Index)
End Sub

Sub CloseSocket(ByVal Index As Long)

    ' Make sure player was/is playing the game, and if so, save'm.
    If Index > 0 Then
        Call LeftGame(Index)
        Call AddLog("Connection from " & GetPlayerIP(Index) & " has been terminated.", "serverlog.txt")
        'Call TextAdd(frmServer.txtText(0), "Connection from " & GetPlayerIP(Index) & " has been terminated.", True)
        frmServer.Socket(Index).Close
        Call UpdateCaption
        Call ClearPlayer(Index)
    Else
        Call AddLog("Invalid index while running Sub CloseSocket, Index = " & Index, "errorlist.txt")
    End If

End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Packet As String

    Packet = GLOBALMSG_CHAR & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)

    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & " for (" & Reason & ").")
    Else
        Call AddLog("Avoided error in Sub HackingAttempt(" & Index & "," & Reason & ")", "errorlist.txt")
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
Sub HandleData(ByVal Index As Long, ByVal Data As String)
On Error GoTo ErrHandler

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
    Dim i As Long, N As Long, X As Long, Y As Long, f As Long
    Dim MapNum As Long
    Dim s As String
    Dim ShopNum As Long, ItemNum As Long
    Dim DurNeeded As Long, GoldNeeded As Long
    Dim z As Long
    Dim Packet As String
    Dim o As Long

    'ParseIndex = index ' SAFE MODE -- Uncomment for ON, comment for OFF
    ' Handle Data
    Parse$ = Split(Data, SEP_CHAR) ' SAFE MODE -- Uncomment for OFF, comment for ON

    'ZePacket = Split(Data, SEP_CHAR) ' SAFE MODE -- Uncomment for ON, comment for OFF
    'NumParse = UBound(ZePacket) ' SAFE MODE -- Uncomment for ON, comment for OFF
    ' Parse's Without Being Online
    If Not IsPlaying(Index) Then

        Select Case (Parse$(0))

            Case GETINFO_CHAR
                Call SendInfo(Index)
                Exit Sub

            Case GATGLASSES_CHAR
                Call SendNewCharClasses(Index)
                Exit Sub

            Case NEWFACCOUNTIED_CHAR

                If Not IsLoggedIn(Index) Then
                    Name = Parse$(1)
                    Password = Parse$(2)

                    For i = 1 To Len(Name)
                        N = Asc(Mid$(Name, i, 1))

                        If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                        Else
                            Call PlainMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.", 1)
                            Exit Sub
                        End If

                    Next

                    If Not AccountExist(Name) Then
                        Call AddAccount(Index, Name, Password)
                        'Call AddLog("Account " & Name & " has been created.", "serverlog.txt")
                        'Call TextAdd(frmServer.txtText(0), "Account " & Name & " has been created.", True)
                        Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                        Call PlainMsg(Index, "Your account has been created!", 1)
                    Else
                        Call PlainMsg(Index, "Sorry, that account name is already taken!", 1)
                    End If
                End If

                Exit Sub

            Case DELIMACCOUNTED_CHAR

                If Not IsLoggedIn(Index) Then
                    Name = Parse$(1)
                    Password = Parse$(2)

                    If Not AccountExist(Name) Then
                        Call PlainMsg(Index, "That account name does not exist.", 2)
                        Exit Sub
                    End If

                    If Not PasswordOK(Name, Password) Then
                        Call PlainMsg(Index, "Incorrect password.", 2)
                        Exit Sub
                    End If

                    Call LoadPlayer(Index, Name)

                    For i = 1 To MAX_CHARS

                        If Trim$(Player(Index).Char(i).Name) <> vbNullString Then
                            Call DeleteName(Player(Index).Char(i).Name)
                        End If

                    Next

                    Call ClearPlayer(Index)
                    Call Kill(App.Path & "\accounts\" & Trim$(Name) & ".ini")
                    Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
                    Call PlainMsg(Index, "Your account has been deleted.", 2)
                End If

                Exit Sub

            Case LOGINATION_CHAR

                If Not IsLoggedIn(Index) Then
                    Name = Parse$(1)
                    Password = Parse$(2)

                    For i = 1 To Len(Name)
                        N = Asc(Mid$(Name, i, 1))

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
                    temp1 = Parse$(6)
                    temp2 = Parse$(3) & "." & Parse$(4) & "." & Parse$(5)
                    temp3 = Encryptor.EncryptData(temp1, temp2)
                    temp4 = vbNullString

                    For i = 0 To UBound(temp3)
                        temp4 = temp4 & Right$("0" & Hex$(temp3(i)), 2)
                    Next

                    temp1 = SEC_CODE
                    temp2 = CLIENT_MAJOR & "." & CLIENT_MINOR & "." & CLIENT_REVISION
                    temp3 = Encryptor.EncryptData(temp1, temp2)
                    temp5 = vbNullString

                    For i = 0 To UBound(temp3)
                        temp5 = temp5 & Right$("0" & Hex$(temp3(i)), 2)
                    Next

                    If temp4 <> temp5 Then
                        'Call SendDataTo(Index, SOUND_CHAR & SEP_CHAR & "ANewVersionHasBeenReleased" & END_CHAR)
                        Call SendSound(Index, NEWVERSIONRELEASED_SOUND, SDT)
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

                    'If frmServer.Closed.Value = Checked Then
                    '    Call PlainMsg(Index, "The server is closed at the moment!", 3)
                    '    Exit Sub
                    'End If

                    Dim Packs As String

                    Packs = MAXINFO_CHAR & SEP_CHAR
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
                    Packs = Packs & END_CHAR
                    Call SendDataTo(Index, Packs)
                    Call LoadPlayer(Index, Name)
                    Call SendChars(Index)
                    Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
                    'Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", True)
                End If

                Exit Sub

            Case ADDACHARA_CHAR
                Name = Parse$(1)
                Sex = Val(Parse$(2))
                Class = Val(Parse$(3))
                CharNum = Val(Parse$(4))

                For i = 1 To Len(Name)
                    N = Asc(Mid$(Name, i, 1))

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

                Call AddChar(Index, Name, Sex, Class, CharNum)
                Call SavePlayer(Index)
                Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
                Call SendChars(Index)
                Call PlainMsg(Index, "Character has been created!", 5)
                Exit Sub

            Case DELIMBOCHARU_CHAR
                CharNum = Val(Parse$(1))

                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(Index, "Invalid CharNum")
                    Exit Sub
                End If

                Call DelChar(Index, CharNum)
                Call AddLog("Character deleted on " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
                Call SendChars(Index)
                Call PlainMsg(Index, "Character has been deleted!", 5)
                Exit Sub

            Case USAGAKARIM_CHAR
                CharNum = Val(Parse$(1))

                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(Index, "Invalid CharNum")
                    Exit Sub
                End If

                If CharExist(Index, CharNum) Then
                    Player(Index).CharNum = CharNum

                    'If frmServer.GMOnly.Value = Checked Then
                    '    If GetPlayerAccess(Index) <= 0 Then
                    '        Call PlainMsg(Index, "The server is only available to GMs at the moment!", 5)

                            'Call HackingAttempt(Index, "The server is only available to GMs at the moment!")
                    '        Exit Sub
                    '    End If
                    'End If

                    Call JoinGame(Index)
                    CharNum = Player(Index).CharNum
                    Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
                    'Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", True)
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
    
    ' Make sure we don't receive any RTE 9's because of an invalid index!
    If Index <= 0 Then
        Call AddLog("Avoided error in HandleData, invalid index value: " & Index, "errorlist.txt")
        Exit Sub
    End If

    Select Case (Parse$(0))

            ' :::::::::::::::::::
            ' :: Guilds Packet ::
            ' :::::::::::::::::::
            ' Access
        Case GUILDCHANGEACCESS_CHAR

            ' Check the requirements.
            If FindPlayer(Parse$(1)) = 0 Then
                Call PlayerMsg(Index, "Player is offline", White)
                Exit Sub
            End If

            If GetPlayerGuild(FindPlayer(Parse$(1))) <> GetPlayerGuild(Index) Then
                Call PlayerMsg(Index, "Player is not in your guild", Red)
                Exit Sub
            End If
            
            If Parse$(2) > 5 Or GetPlayerGuildAccess(Index) <> 5 Then
                Call HackingAttempt(Index, "Messing with the packets!")
                Exit Sub
            End If

            'Set the player's new access level
            Call SetPlayerGuildAccess(FindPlayer(Parse$(1)), Parse$(2))
            Call SendPlayerData(FindPlayer(Parse$(1)))
            Exit Sub

            ' Disown
        Case GUILDDISOWN_CHAR

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
            Exit Sub

            ' Leave Guild
        Case GUILDLEAVE_CHAR

            ' Check if they can leave
            If GetPlayerGuild(Index) = vbNullString Then
                Call PlayerMsg(Index, "You are not in a guild.", Red)
                Exit Sub
            End If

            Call SetPlayerGuild(Index, vbNullString)
            Call SetPlayerGuildAccess(Index, 0)
            Call SendPlayerData(Index)
            Exit Sub

            ' Make A New Guild
        Case MAKEGUILD_CHAR

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
            
            If GuildExist(Parse$(2)) = YES Then
                Call PlayerMsg(Index, "The guild has already been created!", Red)
                Exit Sub
            End If

            ' If everything is ok then lets make the guild
            Call MakeGuild(Parse$(2))
            Call SetPlayerGuild(FindPlayer(Parse$(1)), (Parse$(2)))
            Call SetPlayerGuildAccess(FindPlayer(Parse$(1)), 5)
            Call SendPlayerData(FindPlayer(Parse$(1)))
            Exit Sub

            ' Make A Member
        Case GUILDMEMBER_CHAR

            ' Check if its possible to admit the member
            If FindPlayer(Parse$(1)) = 0 Then
                Call PlayerMsg(Index, "Player is offline", White)
                Exit Sub
            End If

            If GetPlayerGuild(FindPlayer(Parse$(1))) <> GetPlayerGuild(Index) Then
                Call PlayerMsg(Index, "That player is not in your guild", Red)
                Exit Sub
            End If

            If GetPlayerGuildAccess(FindPlayer(Parse$(1))) > 1 Then
                Call PlayerMsg(Index, "That player has already been admitted", Red)
                Exit Sub
            End If

            'All has gone well, set the guild access to 1
            Call SetPlayerGuild(FindPlayer(Parse$(1)), GetPlayerGuild(Index))
            Call SetPlayerGuildAccess(FindPlayer(Parse$(1)), 1)
            Call SendPlayerData(FindPlayer(Parse$(1)))
            Exit Sub

            ' Make A Trainie
        Case GUILDTRAINEE_CHAR

            ' Check if its possible to induct member
            If FindPlayer(Parse$(1)) = 0 Then
                Call PlayerMsg(Index, "Player is offline", White)
                Exit Sub
            End If

            If GetPlayerGuild(FindPlayer(Parse$(1))) <> vbNullString Then
                Call PlayerMsg(Index, "Player is already in a guild", Red)
                Exit Sub
            End If

            'It is possible, so set the guild to index's guild, and the access level to 0
            Call SetPlayerGuild(FindPlayer(Parse$(1)), GetPlayerGuild(Index))
            Call SetPlayerGuildAccess(FindPlayer(Parse$(1)), 0)
            Call SendPlayerData(FindPlayer(Parse$(1)))
            Exit Sub

            ' ::::::::::::::::::::
            ' :: Social packets ::
            ' ::::::::::::::::::::
        Case SAYMSG_CHAR
            Msg = Parse$(1)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(Index, "Say Text Modification")
                    Exit Sub
                End If

            Next

            'If frmServer.chkM.Value = Unchecked Then
            '    If GetPlayerAccess(Index) <= 0 Then
            '        Call PlayerMsg(Index, "Map messages have been disabled by the server!", BrightRed)
            '        Exit Sub
            '    End If
            'End If

            Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & ": " & Msg & vbNullString, PLAYER_LOG)
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Msg & vbNullString, SayColor)
            Call MapMsg2(GetPlayerMap(Index), Msg, Index)
            'TextAdd frmServer.txtText(3), GetPlayerName(Index) & " On Map " & GetPlayerMap(Index) & ": " & Msg, True
            Exit Sub

        Case EMOTEMSG_CHAR
            Msg = Parse$(1)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(Index, "Emote Text Modification")
                    Exit Sub
                End If

            Next

            'If frmServer.chkE.Value = Unchecked Then
            '    If GetPlayerAccess(Index) <= 0 Then
            '        Call PlayerMsg(Index, "Emote messages have been disabled by the server!", BrightRed)
            '        Exit Sub
            '    End If
            'End If

            Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Msg, EmoteColor)
            'TextAdd frmServer.txtText(6), GetPlayerName(Index) & " " & Msg, True
            Exit Sub

        Case BROADCASTMSG_CHAR
            Msg = Parse$(1)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(Index, "Broadcast Text Modification")
                    Exit Sub
                End If

            Next

            'If frmServer.chkBC.Value = Unchecked Then
            '    If GetPlayerAccess(Index) <= 0 Then
            '        Call PlayerMsg(Index, "Broadcast messages have been disabled by the server!", BrightRed)
            '        Exit Sub
            '    End If
            'End If

            If Player(Index).Mute = True Then Exit Sub
            s = GetPlayerName(Index) & ": " & Msg
            Call AddLog(s, PLAYER_LOG)
            Call GlobalMsg(s, BroadcastColor)
            'Call TextAdd(frmServer.txtText(0), s, True)
            'TextAdd frmServer.txtText(1), GetPlayerName(Index) & ": " & Msg, True
            Exit Sub

        Case GLOBALMSG_CHAR
            Msg = Parse$(1)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(Index, "Global Text Modification")
                    Exit Sub
                End If

            Next

            'If frmServer.chkG.Value = Unchecked Then
            '    If GetPlayerAccess(Index) <= 0 Then
            '        Call PlayerMsg(Index, "Global messages have been disabled by the server!", BrightRed)
            '        Exit Sub
            '    End If
            'End If

            If Player(Index).Mute = True Then Exit Sub
            If GetPlayerAccess(Index) > 0 Then
                s = "(global) " & GetPlayerName(Index) & ": " & Msg
                Call AddLog(s, ADMIN_LOG)
                Call GlobalMsg(s, GlobalColor)
                'Call TextAdd(frmServer.txtText(0), s, True)
            End If

            Call AddLog(GetPlayerName(Index) & ": " & Msg, "serverlog.txt")
            Exit Sub

        Case ADMINMSG_CHAR
            Msg = Parse$(1)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(Index, "Admin Text Modification")
                    Exit Sub
                End If

            Next

            'If frmServer.chkA.Value = Unchecked Then
            '    Call PlayerMsg(Index, "Admin messages have been disabled by the server!", BrightRed)
            '    Exit Sub
            'End If

            If GetPlayerAccess(Index) > 0 Then
                Call AddLog("(admin " & GetPlayerName(Index) & ") " & Msg, ADMIN_LOG)
                Call AdminMsg("(admin " & GetPlayerName(Index) & ") " & Msg, AdminColor)
            End If

            Call AddLog(GetPlayerName(Index) & ": " & Msg, "serverlog.txt")
            Exit Sub

        Case PLAYERMSG_CHAR
            MsgTo = FindPlayer(Parse$(1))
            Msg = Parse$(2)

            ' Prevent hacking
            For i = 1 To Len(Msg)

                If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
                    Call HackingAttempt(Index, "Player Msg Text Modification")
                    Exit Sub
                End If

            Next

            'If frmServer.chkP.Value = Unchecked Then
            '    If GetPlayerAccess(Index) <= 0 Then
            '        Call PlayerMsg(Index, "PM messages have been disabled by the server!", BrightRed)
            '        Exit Sub
            '    End If
            'End If

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

            Call AddLog("To " & GetPlayerName(MsgTo) & " From " & GetPlayerName(Index) & ": " & Msg, "serverlog.txt")
            Exit Sub

            ' :::::::::::::::::::::::::::::
            ' :: Moving character packet ::
            ' :::::::::::::::::::::::::::::
        Case PLAYERMOVE_CHAR

            If Player(Index).GettingMap = YES Then
                Exit Sub
            End If

            Dir = Val(Parse$(1))
            Movement = Val(Parse$(2))

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
        Case PLAYERDIR_CHAR

            If Player(Index).GettingMap = YES Then
                Exit Sub
            End If

            Dir = Val(Parse$(1))

            ' Prevent hacking
            If Dir < DIR_UP Or Dir > DIR_RIGHT Then
                Call HackingAttempt(Index, "Invalid Direction")
                Exit Sub
            End If

            Call SetPlayerDir(Index, Dir)
            Call SendDataToMapBut(Index, GetPlayerMap(Index), PLAYERDIR_CHAR & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
            Exit Sub

            ' :::::::::::::::::::::
            ' :: Use item packet ::
            ' :::::::::::::::::::::
        Case USEITEM_CHAR
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
                N = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
                Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long, n5 As Long, n6 As Long

                n1 = Item(GetPlayerInvItemNum(Index, InvNum)).StrReq
                n2 = Item(GetPlayerInvItemNum(Index, InvNum)).DefReq
                n3 = Item(GetPlayerInvItemNum(Index, InvNum)).SpeedReq
                n6 = Item(GetPlayerInvItemNum(Index, InvNum)).MagicReq
                n4 = Item(GetPlayerInvItemNum(Index, InvNum)).ClassReq
                n5 = Item(GetPlayerInvItemNum(Index, InvNum)).AccessReq

                ' Find out what kind of item it is
                Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type

                    Case ITEM_TYPE_ARMOR

                        If InvNum <> GetPlayerArmorSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If

                            Call SetPlayerArmorSlot(Index, InvNum)
                        Else
                            Call SetPlayerArmorSlot(Index, 0)
                        End If

                        Call SendWornEquipment(Index)

                    Case ITEM_TYPE_WEAPON

                        If InvNum <> GetPlayerWeaponSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If

                            Call SetPlayerWeaponSlot(Index, InvNum)
                        Else
                            Call SetPlayerWeaponSlot(Index, 0)
                        End If

                        Call SendWornEquipment(Index)

                    Case ITEM_TYPE_HELMET

                        If InvNum <> GetPlayerHelmetSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
                                Exit Sub
                            End If

                            Call SetPlayerHelmetSlot(Index, InvNum)
                        Else
                            Call SetPlayerHelmetSlot(Index, 0)
                        End If

                        Call SendWornEquipment(Index)

                    Case ITEM_TYPE_SHIELD

                        If InvNum <> GetPlayerShieldSlot(Index) Then
                            If n4 > 0 Then
                                If GetPlayerClass(Index) <> n4 Then
                                    Call PlayerMsg(Index, "You need to be class " & GetClassName(n4) & " to use this item!", BrightRed)
                                    Exit Sub
                                End If
                            End If

                            If GetPlayerAccess(Index) < n5 Then
                                Call PlayerMsg(Index, "Your access needs to be higher then " & n5 & "!", BrightRed)
                                Exit Sub
                            End If

                            If Int(GetPlayerstr(Index)) < n1 Then
                                Call PlayerMsg(Index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerDEF(Index)) < n2 Then
                                Call PlayerMsg(Index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerSPEED(Index)) < n3 Then
                                Call PlayerMsg(Index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
                                Exit Sub
                            ElseIf Int(GetPlayerMAGI(Index)) < n6 Then
                                Call PlayerMsg(Index, "Your magic is too low to equip this item!  Required Magic (" & n6 & ")", BrightRed)
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
                        X = DirToX(GetPlayerX(Index), GetPlayerDir(Index))
                        Y = DirToY(GetPlayerY(Index), GetPlayerDir(Index))

                        If Not IsValid(X, Y) Then Exit Sub

                        ' Check if a key exists
                        If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY Then

                            ' Check if the key they are using matches the map key
                            If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(X, Y).Data1 Then
                                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                                Call SendDataToMap(GetPlayerMap(Index), MAPKEY_CHAR & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)

                                If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                                    Call MapMsg(GetPlayerMap(Index), "A door has been unlocked!", White)
                                Else
                                    Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), White)
                                End If

                                'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Key" & END_CHAR)
                                Call SendSound(Index, KEY_SOUND, SDTM)

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
                            Call SendDataToMap(GetPlayerMap(Index), MAPKEY_CHAR & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                            'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Key" & END_CHAR)
                            Call SendSound(Index, KEY_SOUND, SDTM)
                        End If

                    Case ITEM_TYPE_SPELL

                        ' Get the spell num
                        N = Item(GetPlayerInvItemNum(Index, InvNum)).Data1

                        If N > 0 Then

                            ' Make sure they are the right class
                            If Spell(N).ClassReq = GetPlayerClass(Index) Or Spell(N).ClassReq = 0 Then
                                If Spell(N).LevelReq = 0 And Player(Index).Char(Player(Index).CharNum).Access < 1 Then
                                    Call PlayerMsg(Index, "This spell can only be used by admins!", BrightRed)
                                    Exit Sub
                                End If

                                ' Make sure they are the right level
                                i = GetSpellReqLevel(N)

                                If n6 > i Then i = n6
                                If i <= GetPlayerLevel(Index) Then
                                    i = FindOpenSpellSlot(Index)

                                    ' Make sure they have an open spell slot
                                    If i > 0 Then

                                        ' Make sure they dont already have the spell
                                        If Not HasSpell(Index, N) Then
                                            Call SetPlayerSpell(Index, i, N)
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
                                Call PlayerMsg(Index, "This spell can only be learned by a " & GetClassName(Spell(N).ClassReq) & ".", White)
                            End If

                        Else
                            Call PlayerMsg(Index, "This scroll is not connected to a spell, please inform an admin!", White)
                        End If

                    Case ITEM_TYPE_PET
                        Player(Index).Pet.Alive = YES
                        Player(Index).Pet.Sprite = Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                        Player(Index).Pet.Dir = DIR_UP
                        Player(Index).Pet.Map = GetPlayerMap(Index)
                        Player(Index).Pet.MapToGo = 0
                        Player(Index).Pet.X = GetPlayerX(Index) + Int(Rnd * 3 - 1)

                        If Player(Index).Pet.X < 0 Or Player(Index).Pet.X > MAX_MAPX Then Player(Index).Pet.X = GetPlayerX(Index)
                        Player(Index).Pet.XToGo = -1
                        Player(Index).Pet.Y = GetPlayerY(Index) + Int(Rnd * 3 - 1)

                        If Player(Index).Pet.Y < 0 Or Player(Index).Pet.Y > MAX_MAPY Then Player(Index).Pet.Y = GetPlayerY(Index)
                        Player(Index).Pet.YToGo = -1
                        Player(Index).Pet.Level = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
                        Player(Index).Pet.HP = Player(Index).Pet.Level * 5
                        Call AddToGrid(Player(Index).Pet.Map, Player(Index).Pet.X, Player(Index).Pet.Y)
                        Packet = PETDATA_CHAR & SEP_CHAR
                        Packet = Packet & Index & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.X & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.Y & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.Sprite & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
                        Packet = Packet & Player(Index).Pet.Level * 5 & SEP_CHAR
                        Packet = Packet & END_CHAR
                        Call SendDataToMap(GetPlayerMap(Index), Packet)

                        ' Excuse the ugly code, I'm rushing
                        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                        Call PlayerMsg(Index, "You got a pet!", White)
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
        Case ATTACK_CHAR

            If GetPlayerWeaponSlot(Index) > 0 Then
                If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 > 0 Then
                    Call SendDataToMap(GetPlayerMap(Index), CHECKARROWS_CHAR & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
                    Exit Sub
                End If
            End If

            ' Try to attack a player
            For i = 1 To MAX_PLAYERS

                ' Make sure we dont try to attack ourselves
                If i <> Index Then

                    ' Can we attack the player?
                    If CanAttackPlayer(Index, i) Then
                        If Not CanPlayerBlockHit(i) Then

                            ' Get the damage we can do
                            If Not CanPlayerCriticalHit(Index) Then
                                Damage = GetPlayerDamage(Index) - GetPlayerProtection(i) + (Rnd * 5) - 2
                                'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & END_CHAR)
                                Call SendSound(Index, ATTACK_SOUND, SDTM)
                            Else
                                N = GetPlayerDamage(Index)
                                Damage = N + Int(Rnd * Int(N / 2)) + 1 - GetPlayerProtection(i) + (Rnd * 5) - 2
                                Call BattleMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan, 0)
                                Call BattleMsg(i, GetPlayerName(Index) & " swings with enormous might!", BrightCyan, 1)

                                'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                                'Call PlayerMsg(I, GetPlayerName(index) & " swings with enormous might!", BrightCyan)
                                'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Blow0" & END_CHAR)
                                Call SendSound(Index, CRITICALHIT_SOUND, SDTM)
                            End If

                            If Damage > 0 Then
                                Call AttackPlayer(Index, i, Damage)
                            Else
                                Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                                'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "miss" & END_CHAR)
                                Call SendSound(Index, MISS_SOUND, SDTM)
                            End If

                        Else
                            Call BattleMsg(Index, GetPlayerName(i) & " blocked your hit!", BrightCyan, 0)
                            Call BattleMsg(i, "You blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan, 1)

                            'Call PlayerMsg(index, GetPlayerName(I) & "'s " & Trim$(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked your hit!", BrightCyan)
                            'Call PlayerMsg(I, "Your " & Trim$(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                            'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "miss" & END_CHAR)
                            Call SendSound(Index, MISS_SOUND, SDTM)
                        End If

                        Exit Sub
                    End If
                End If

            Next

            ' Try to attack a npc
            For i = 1 To MAX_MAP_NPCS

                ' Can we attack the npc?
                If CanAttackNpc(Index, i) Then

                    ' Get the damage we can do
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), i).num).DEF / 2) + (Rnd * 5) - 2
                        'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & END_CHAR)
                        Call SendSound(Index, ATTACK_SOUND, SDTM)
                    Else
                        N = GetPlayerDamage(Index)
                        Damage = N + Int(Rnd * Int(N / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), i).num).DEF / 2) + (Rnd * 5) - 2
                        Call BattleMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan, 0)

                        'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                        'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Blow0" & END_CHAR)
                        Call SendSound(Index, CRITICALHIT_SOUND, SDTM)
                    End If

                    If Damage > 0 Then
                        Call AttackNpc(Index, i, Damage)
                        Call SendDataTo(Index, BLITPLAYERDMG_CHAR & SEP_CHAR & Damage & SEP_CHAR & i & END_CHAR)
                    Else
                        Call BattleMsg(Index, "Your attack does nothing.", BrightRed, 0)

                        'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                        Call SendDataTo(Index, BLITPLAYERDMG_CHAR & SEP_CHAR & Damage & SEP_CHAR & i & END_CHAR)
                        'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "miss" & END_CHAR)
                        Call SendSound(Index, MISS_SOUND, SDTM)
                    End If

                    Exit Sub
                End If

            Next

            Exit Sub

            ' ::::::::::::::::::::::
            ' :: Use stats packet ::
            ' ::::::::::::::::::::::
        Case USESTATPOINT_CHAR
            PointType = Val(Parse$(1))

            ' Prevent hacking
            If (PointType < 0) Or (PointType > 3) Then
                Call HackingAttempt(Index, "Invalid Point Type")
                Exit Sub
            End If

            ' Make sure they have points
            If GetPlayerPOINTS(Index) > 0 Then
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "UsingStatPoints " & Index & "," & PointType
                Else

                    Select Case PointType

                        Case 0
                            Call SetPlayerstr(Index, GetPlayerstr(Index) + 1)
                            Call BattleMsg(Index, "You have gained more strength!", 15, 0)
                            'Call SendDataTo(Index, SOUND_CHAR & SEP_CHAR & "strengthRaised" & END_CHAR)
                            Call SendSound(Index, STRENGTHRAISED_SOUND, SDT)

                        Case 1
                            Call SetPlayerDEF(Index, GetPlayerDEF(Index) + 1)
                            Call BattleMsg(Index, "You have gained more defense!", 15, 0)
                            'Call SendDataTo(Index, SOUND_CHAR & SEP_CHAR & "DefenseRaised" & END_CHAR)
                            Call SendSound(Index, DEFENSERAISED_SOUND, SDT)

                        Case 2
                            Call SetPlayerMAGI(Index, GetPlayerMAGI(Index) + 1)
                            Call BattleMsg(Index, "You have gained more magic abilities!", 15, 0)
                            'Call SendDataTo(Index, SOUND_CHAR & SEP_CHAR & "MagicRaised" & END_CHAR)
                            Call SendSound(Index, MAGICRAISED_SOUND, SDT)

                        Case 3
                            Call SetPlayerSPEED(Index, GetPlayerSPEED(Index) + 1)
                            Call BattleMsg(Index, "You have gained more speed!", 15, 0)
                            'Call SendDataTo(Index, SOUND_CHAR & SEP_CHAR & "SpeedRaised" & END_CHAR)
                            Call SendSound(Index, SPEEDRAISED_SOUND, SDT)
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
            Exit Sub

            ' ::::::::::::::::::::::::::::::::
            ' :: Player info request packet ::
            ' ::::::::::::::::::::::::::::::::
        Case PLAYERINFOREQUEST_CHAR
            Name = Parse$(1)
            i = FindPlayer(Name)

            If i > 0 Then
                Call PlayerMsg(Index, "Account: " & Trim$(Player(i).Login) & ", Name: " & GetPlayerName(i), BrightGreen)

                If GetPlayerAccess(Index) > ADMIN_MONITER Then
                    Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
                    Call PlayerMsg(Index, "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), BrightGreen)
                    Call PlayerMsg(Index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
                    Call PlayerMsg(Index, "str: " & GetPlayerstr(i) & "  DEF: " & GetPlayerDEF(i) & "  MAGI: " & GetPlayerMAGI(i) & "  SPEED: " & GetPlayerSPEED(i), BrightGreen)
                    N = Int(GetPlayerstr(i) / 2) + Int(GetPlayerLevel(i) / 2)
                    i = Int(GetPlayerDEF(i) / 2) + Int(GetPlayerLevel(i) / 2)

                    If N > 100 Then N = 100
                    If i > 100 Then i = 100
                    Call PlayerMsg(Index, "Critical Hit Chance: " & N & "%, Block Chance: " & i & "%", BrightGreen)
                End If

            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If

            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Set sprite packet ::
            ' :::::::::::::::::::::::
        Case SETSPRITE_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' The sprite
            N = Val(Parse$(1))
            Call SetPlayerSprite(Index, N)
            Call SendPlayerData(Index)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Set player sprite packet ::
            ' ::::::::::::::::::::::::::::::
        Case SETPLAYERSPRITE_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' The sprite
            i = FindPlayer(Parse$(1))
            N = Val(Parse$(2))
            Call SetPlayerSprite(i, N)
            Call SendPlayerData(i)
            Exit Sub

            ' ::::::::::::::::::::::::::
            ' :: Stats request packet ::
            ' ::::::::::::::::::::::::::
        Case GETSTATS_CHAR
            Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(Index) & " -=-", White)
            Call PlayerMsg(Index, "Level: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), White)
            Call PlayerMsg(Index, "HP: " & GetPlayerHP(Index) & "/" & GetPlayerMaxHP(Index) & "  MP: " & GetPlayerMP(Index) & "/" & GetPlayerMaxMP(Index) & "  SP: " & GetPlayerSP(Index) & "/" & GetPlayerMaxSP(Index), White)
            Call PlayerMsg(Index, "str: " & GetPlayerstr(Index) & "  DEF: " & GetPlayerDEF(Index) & "  MAGI: " & GetPlayerMAGI(Index) & "  SPEED: " & GetPlayerSPEED(Index), White)
            N = Int(GetPlayerstr(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
            i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)

            If N > 100 Then N = 100
            If i > 100 Then i = 100
            Call PlayerMsg(Index, "Critical Hit Chance: " & N & "%, Block Chance: " & i & "%", White)
            Exit Sub

            ' ::::::::::::::::::::::::::::::::::
            ' :: Player request for a new map ::
            ' ::::::::::::::::::::::::::::::::::
        Case REQUESTNEWMAP_CHAR
            Dir = Val(Parse$(1))

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
        Case MAPDATA_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = 1
            MapNum = GetPlayerMap(Index)
            Call ClearMap(MapNum)
            Map(MapNum).Name = Parse$(N + 1)
            Map(MapNum).Revision = Val(Parse$(N + 2)) + 1
            Map(MapNum).Moral = Val(Parse$(N + 3))
            Map(MapNum).Up = Val(Parse$(N + 4))
            Map(MapNum).Down = Val(Parse$(N + 5))
            Map(MapNum).Left = Val(Parse$(N + 6))
            Map(MapNum).Right = Val(Parse$(N + 7))
            Map(MapNum).Music = Parse$(N + 8)
            Map(MapNum).BootMap = Val(Parse$(N + 9))
            Map(MapNum).BootX = Val(Parse$(N + 10))
            Map(MapNum).BootY = Val(Parse$(N + 11))
            Map(MapNum).Indoors = Val(Parse$(N + 12))
            N = N + 13
            i = GetPlayerMap(Index)

            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Ground = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).GroundSet = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Mask = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).MaskSet = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Anim = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).AnimSet = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Fringe = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).FringeSet = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Type = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Data1 = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Data2 = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Data3 = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).String1 = Parse$(N)
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).String2 = Parse$(N)
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).String3 = Parse$(N)
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Mask2 = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Mask2Set = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).M2Anim = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).M2AnimSet = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).FAnim = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).FAnimSet = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Fringe2 = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Fringe2Set = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).Light = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).F2Anim = Val(Parse$(N))
                        N = N + 1
                    End If

                    If Parse$(N) <> NEXT_CHAR Then
                        Map(i).Tile(X, Y).F2AnimSet = Val(Parse$(N))
                        N = N + 1
                    End If

                    N = N + 1
                Next
            Next

            For X = 1 To MAX_MAP_NPCS
                Map(MapNum).Npc(X) = Val(Parse$(N))
                Map(MapNum).NpcSpawn(X).Used = Val(Parse$(N + 1))
                Map(MapNum).NpcSpawn(X).X = Val(Parse$(N + 2))
                Map(MapNum).NpcSpawn(X).Y = Val(Parse$(N + 3))
                N = N + 4
                Call ClearMapNpc(X, MapNum)
            Next

            ' Clear out it all
            For i = 1 To MAX_MAP_ITEMS
                Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
                Call ClearMapItem(i, GetPlayerMap(Index))
            Next

            ' Save the map
            Call SaveMap(MapNum)

            ' Respawn
            Call SpawnMapItems(GetPlayerMap(Index))

            ' Respawn NPCS
            Call SpawnMapNpcs(GetPlayerMap(Index))

            ' Reset grid
            Call ResetMapGrid(GetPlayerMap(Index))

            ' Refresh map for everyone online
            For i = 1 To MAX_PLAYERS

                If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                    Call SendDataTo(i, CHECKFORMAP_CHAR & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & END_CHAR)

                    'Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
                End If

            Next

            Exit Sub

            ' ::::::::::::::::::::::::::::
            ' :: Need map yes/no packet ::
            ' ::::::::::::::::::::::::::::
        Case NEEDMAP_CHAR

            ' Get yes/no value
            s = LCase$(Parse$(1))

            If s = "yes" Then
                Call SendMap(Index, GetPlayerMap(Index))
                Call SendMapItemsTo(Index, GetPlayerMap(Index))
                Call SendMapNpcsTo(Index, GetPlayerMap(Index))
                Call SendJoinMap(Index)
                Player(Index).GettingMap = NO
                Call SendDataTo(Index, MAPDONE_CHAR & END_CHAR)
            Else
                Call SendMapItemsTo(Index, GetPlayerMap(Index))
                Call SendMapNpcsTo(Index, GetPlayerMap(Index))
                Call SendJoinMap(Index)
                Player(Index).GettingMap = NO
                Call SendDataTo(Index, MAPDONE_CHAR & END_CHAR)
            End If

            Exit Sub

            ' :::::::::::::::::::::::::::::::::::::::::::::::
            ' :: Player trying to pick up something packet ::
            ' :::::::::::::::::::::::::::::::::::::::::::::::
        Case MAPGETITEM_CHAR
            Call PlayerMapGetItem(Index)
            Exit Sub

            ' ::::::::::::::::::::::::::::::::::::::::::::
            ' :: Player trying to drop something packet ::
            ' ::::::::::::::::::::::::::::::::::::::::::::
        Case MAPDROPITEM_CHAR
            InvNum = Val(Parse$(1))
            Amount = Val(Parse$(2))

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
        Case MAPRESPAWN_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' Clear out it all
            For i = 1 To MAX_MAP_ITEMS
                Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
                Call ClearMapItem(i, GetPlayerMap(Index))
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
        Case KICKPLAYER_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) <= 0 Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' The player index
            N = FindPlayer(Parse$(1))

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
        Case BANLIST_CHAR

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
        Case BANDESTROY_CHAR

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
        Case BANPLAYER_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' The player index
            N = FindPlayer(Parse$(1))

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
        Case REQUESTEDITMAP_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, EDITMAP_CHAR & END_CHAR)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Request edit item packet ::
            ' ::::::::::::::::::::::::::::::
        Case REQUESTEDITITEM_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, ITEMEDITOR_CHAR & END_CHAR)
            Exit Sub

            ' ::::::::::::::::::::::
            ' :: Edit item packet ::
            ' ::::::::::::::::::::::
        Case EDITITEM_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' The item #
            N = Val(Parse$(1))

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
        Case SAVEITEM_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse$(1))

            If N < 0 Or N > MAX_ITEMS Then
                Call HackingAttempt(Index, "Invalid Item Index")
                Exit Sub
            End If

            ' Update the item
            Item(N).Name = Parse$(2)
            Item(N).Pic = Val(Parse$(3))
            Item(N).Type = Val(Parse$(4))
            Item(N).Data1 = Val(Parse$(5))
            Item(N).Data2 = Val(Parse$(6))
            Item(N).Data3 = Val(Parse$(7))
            Item(N).StrReq = Val(Parse$(8))
            Item(N).DefReq = Val(Parse$(9))
            Item(N).SpeedReq = Val(Parse$(10))
            Item(N).MagicReq = Val(Parse$(11))
            Item(N).ClassReq = Val(Parse$(12))
            Item(N).AccessReq = Val(Parse$(13))
            Item(N).AddHP = Val(Parse$(14))
            Item(N).AddMP = Val(Parse$(15))
            Item(N).AddSP = Val(Parse$(16))
            Item(N).AddStr = Val(Parse$(17))
            Item(N).AddDef = Val(Parse$(18))
            Item(N).AddMagi = Val(Parse$(19))
            Item(N).AddSpeed = Val(Parse$(20))
            Item(N).AddEXP = Val(Parse$(21))
            Item(N).Desc = Parse$(22)
            Item(N).AttackSpeed = Val(Parse$(23))

            ' Save it
            Call SendUpdateItemToAll(N)
            Call SaveItem(N)
            Call AddLog(GetPlayerName(Index) & " saved item #" & N & ".", ADMIN_LOG)
            Exit Sub

            ' :::::::::::::::::::::::::::::
            ' :: Request edit npc packet ::
            ' :::::::::::::::::::::::::::::
        Case REQUESTEDITNPC_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, NPCEDITOR_CHAR & END_CHAR)
            Exit Sub

            ' :::::::::::::::::::::
            ' :: Edit npc packet ::
            ' :::::::::::::::::::::
        Case EDITNPC_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' The npc #
            N = Val(Parse$(1))

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
        Case SAVENPC_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse$(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_NPCS Then
                Call HackingAttempt(Index, "Invalid NPC Index")
                Exit Sub
            End If

            ' Update the npc
            Npc(N).Name = Parse$(2)
            Npc(N).AttackSay = Parse$(3)
            Npc(N).Sprite = Val(Parse$(4))
            Npc(N).SpawnSecs = Val(Parse$(5))
            Npc(N).Behavior = Val(Parse$(6))
            Npc(N).Range = Val(Parse$(7))
            Npc(N).STR = Val(Parse$(8))
            Npc(N).DEF = Val(Parse$(9))
            Npc(N).Speed = Val(Parse$(10))
            Npc(N).Magi = Val(Parse$(11))
            Npc(N).Big = Val(Parse$(12))
            Npc(N).MaxHp = Val(Parse$(13))
            Npc(N).Exp = Val(Parse$(14))
            Npc(N).SpawnTime = Val(Parse$(15))
            Npc(N).Speech = Val(Parse$(16))
            z = 17

            For i = 1 To MAX_NPC_DROPS
                Npc(N).ItemNPC(i).Chance = Val(Parse$(z))
                Npc(N).ItemNPC(i).ItemNum = Val(Parse$(z + 1))
                Npc(N).ItemNPC(i).ItemValue = Val(Parse$(z + 2))
                z = z + 3
            Next

            ' Save it
            Call SendUpdateNpcToAll(N)
            Call SaveNpc(N)
            Call AddLog(GetPlayerName(Index) & " saved npc #" & N & ".", ADMIN_LOG)
            Exit Sub

            ' ::::::::::::::::::::::::::::::
            ' :: Request edit shop packet ::
            ' ::::::::::::::::::::::::::::::
        Case REQUESTEDITSHOP_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, SHOPEDITOR_CHAR & END_CHAR)
            Exit Sub

            ' ::::::::::::::::::::::
            ' :: Edit shop packet ::
            ' ::::::::::::::::::::::
        Case EDITSHOP_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' The shop #
            N = Val(Parse$(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_SHOPS Then
                Call HackingAttempt(Index, "Invalid Shop Index")
                Exit Sub
            End If

            Call AddLog(GetPlayerName(Index) & " editing shop #" & N & ".", ADMIN_LOG)
            Call SendEditShopTo(Index, N)
            Exit Sub

        Case ADDFRIEND_CHAR
            Name = Trim$(Parse$(1))

            If Not FindChar(Name) Then
                Call PlayerMsg(Index, "No such player exists!", Blue)
                Exit Sub
            End If

            If Name = GetPlayerName(Index) Then
                Call PlayerMsg(Index, "You can't add yourself!", Blue)
                Exit Sub
            End If

            For i = 1 To MAX_FRIENDS

                If Player(Index).Char(Player(Index).CharNum).Friends(i) = Name Then
                    Call PlayerMsg(Index, "You already have that user as a friend!", Blue)
                    Exit Sub
                End If

            Next

            For i = 1 To MAX_FRIENDS

                If Player(Index).Char(Player(Index).CharNum).Friends(i) = vbNullString Then
                    Player(Index).Char(Player(Index).CharNum).Friends(i) = Name
                    Call PlayerMsg(Index, "Friend added.", Blue)
                    Call SendFriendListTo(Index)
                    Exit Sub
                End If

            Next

            Call PlayerMsg(Index, "Sorry, but you have too many friends already.", Blue)
            Exit Sub

        Case REMOVEFRIEND_CHAR
            Name = Trim$(Parse$(1))

            For i = 1 To MAX_FRIENDS

                If Player(Index).Char(Player(Index).CharNum).Friends(i) = Name Then
                    Player(Index).Char(Player(Index).CharNum).Friends(i) = vbNullString
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
        Case SAVESHOP_CHAR

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
            N = 6

            For z = 1 To 6
                For i = 1 To MAX_TRADES
                    Shop(ShopNum).TradeItem(z).Value(i).GiveItem = Val(Parse$(N))
                    Shop(ShopNum).TradeItem(z).Value(i).GiveValue = Val(Parse$(N + 1))
                    Shop(ShopNum).TradeItem(z).Value(i).GetItem = Val(Parse$(N + 2))
                    Shop(ShopNum).TradeItem(z).Value(i).GetValue = Val(Parse$(N + 3))
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
        Case REQUESTEDITMAIN_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_CREATOR Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            f = FreeFile
            Open App.Path & "\Scripts\Main.txt" For Input As #f
            Call SendDataTo(Index, MAINEDITOR_CHAR & SEP_CHAR & Input$(LOF(f), f) & END_CHAR)
            Close #f
            Exit Sub

            ' :::::::::::::::::::::::::::::::
            ' :: Request edit spell packet ::
            ' :::::::::::::::::::::::::::::::
        Case REQUESTEDITSPELL_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, SPELLEDITOR_CHAR & END_CHAR)
            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Edit spell packet ::
            ' :::::::::::::::::::::::
        Case EDITSPELL_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' The spell #
            N = Val(Parse$(1))

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
        Case SAVESPELL_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            ' Spell #
            N = Val(Parse$(1))

            ' Prevent hacking
            If N < 0 Or N > MAX_SPELLS Then
                Call HackingAttempt(Index, "Invalid Spell Index")
                Exit Sub
            End If

            ' Update the spell
            Spell(N).Name = Parse$(2)
            Spell(N).ClassReq = Val(Parse$(3))
            Spell(N).LevelReq = Val(Parse$(4))
            Spell(N).Type = Val(Parse$(5))
            Spell(N).Data1 = Val(Parse$(6))
            Spell(N).Data2 = Val(Parse$(7))
            Spell(N).Data3 = Val(Parse$(8))
            Spell(N).MPCost = Val(Parse$(9))
            Spell(N).sound = Val(Parse$(10))
            Spell(N).Range = Val(Parse$(11))
            Spell(N).SpellAnim = Val(Parse$(12))
            Spell(N).SpellTime = Val(Parse$(13))
            Spell(N).SpellDone = Val(Parse$(14))
            Spell(N).AE = Val(Parse$(15))

            ' Save it
            Call SendUpdateSpellToAll(N)
            Call SaveSpell(N)
            Call AddLog(GetPlayerName(Index) & " saving spell #" & N & ".", ADMIN_LOG)
            Exit Sub

            ' :::::::::::::::::::::::
            ' :: Set access packet ::
            ' :::::::::::::::::::::::
        Case SETACCESS_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_CREATOR Then
                Call HackingAttempt(Index, "Trying to use powers not available")
                Exit Sub
            End If

            ' The index
            N = FindPlayer(Parse$(1))

            ' The access
            i = Val(Parse$(2))

            ' Check for invalid access level
            If i >= 0 Or i <= 3 Then
                If GetPlayerName(Index) <> GetPlayerName(N) Then
                    If GetPlayerAccess(Index) > GetPlayerAccess(N) Then

                        ' Check if player is on
                        If N > 0 Then
                            If GetPlayerAccess(N) <= 0 Then
                                Call GlobalMsg(GetPlayerName(N) & " has been blessed with administrative access.", BrightBlue)
                            End If

                            Call SetPlayerAccess(N, i)
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

        Case WHOSONLINE_CHAR
            Call SendWhosOnline(Index)
            Exit Sub

        Case ONLINELIST_CHAR
            Call SendOnlineList
            Exit Sub

        Case SETMOTD_CHAR

            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SpecialPutVar(App.Path & "\motd.ini", "MOTD", "Msg", Parse$(1))
            Call GlobalMsg("MOTD changed to: " & Parse$(1), BrightCyan)
            Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Parse$(1), ADMIN_LOG)
            Exit Sub

        Case TRADEREQUEST_CHAR

            ' Trade num
            N = Val(Parse$(1))
            z = Val(Parse$(2))

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
            ' I = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
            i = Val(Parse$(3))

            ' Check if inv full
            If i <= 0 Then Exit Sub
            X = FindOpenInvSlot(Index, Shop(i).TradeItem(N).Value(z).GetItem)

            If X = 0 Then
                Call PlayerMsg(Index, "Trade unsuccessful, inventory full.", BrightRed)
                Exit Sub
            End If

            ' Check if they have the item
            If HasItem(Index, Shop(i).TradeItem(N).Value(z).GiveItem) >= Shop(i).TradeItem(N).Value(z).GiveValue Then
                Call TakeItem(Index, Shop(i).TradeItem(N).Value(z).GiveItem, Shop(i).TradeItem(N).Value(z).GiveValue)
                Call GiveItem(Index, Shop(i).TradeItem(N).Value(z).GetItem, Shop(i).TradeItem(N).Value(z).GetValue)
                Call PlayerMsg(Index, "The trade was successful!", Yellow)
            Else
                Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
            End If

            Exit Sub

        Case FIXITEM_CHAR

            ' Inv num
            N = Val(Parse$(1))

            ' Make sure its a equipable item
            If Item(GetPlayerInvItemNum(Index, N)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(Index, N)).Type > ITEM_TYPE_SHIELD Then
                Call PlayerMsg(Index, "You can only fix weapons, armors, helmets, and shields.", BrightRed)
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
            i = Int(Item(GetPlayerInvItemNum(Index, N)).Data2 / 5)

            If i <= 0 Then i = 1
            DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(Index, N)
            GoldNeeded = Int(DurNeeded * i / 2)

            If GoldNeeded <= 0 Then GoldNeeded = 1

            ' Check if they even need it repaired
            If DurNeeded <= 0 Then
                Call PlayerMsg(Index, "This item is in perfect condition!", White)
                Exit Sub
            End If

            ' Check if they have enough for at least one point
            If HasItem(Index, 1) >= i Then

                ' Check if they have enough for a total restoration
                If HasItem(Index, 1) >= GoldNeeded Then
                    Call TakeItem(Index, 1, GoldNeeded)
                    Call SetPlayerInvItemDur(Index, N, Item(ItemNum).Data1 * -1)
                    Call PlayerMsg(Index, "Item has been totally restored for " & GoldNeeded & " gold!", BrightBlue)
                Else

                    ' They dont so restore as much as we can
                    DurNeeded = (HasItem(Index, 1) / i)
                    GoldNeeded = Int(DurNeeded * i / 2)

                    If GoldNeeded <= 0 Then GoldNeeded = 1
                    Call TakeItem(Index, 1, GoldNeeded)
                    Call SetPlayerInvItemDur(Index, N, GetPlayerInvItemDur(Index, N) + DurNeeded)
                    Call PlayerMsg(Index, "Item has been partially fixed for " & GoldNeeded & " gold!", BrightBlue)
                End If

            Else
                Call PlayerMsg(Index, "Insufficient gold to fix this item!", BrightRed)
            End If

            Exit Sub

        Case SEARCH_CHAR
            X = Val(Parse$(1))
            Y = Val(Parse$(2))

            ' Prevent subscript out of range
            If X < 0 Or X > MAX_MAPX Or Y < 0 Or Y > MAX_MAPY Then
                Exit Sub
            End If

            ' Check for a player
            For i = 1 To MAX_PLAYERS

                If IsPlaying(i) And GetPlayerMap(Index) = GetPlayerMap(i) And GetPlayerX(i) = X And GetPlayerY(i) = Y Then

                    ' Consider the player
                    If GetPlayerLevel(i) >= GetPlayerLevel(Index) + 5 Then
                        Call PlayerMsg(Index, "You wouldn't stand a chance.", BrightRed)
                    Else

                        If GetPlayerLevel(i) > GetPlayerLevel(Index) Then
                            Call PlayerMsg(Index, "This one seems to have an advantage over you.", Yellow)
                        Else

                            If GetPlayerLevel(i) = GetPlayerLevel(Index) Then
                                Call PlayerMsg(Index, "This would be an even fight.", White)
                            Else

                                If GetPlayerLevel(Index) >= GetPlayerLevel(i) + 5 Then
                                    Call PlayerMsg(Index, "You could slaughter that player.", BrightBlue)
                                Else

                                    If GetPlayerLevel(Index) > GetPlayerLevel(i) Then
                                        Call PlayerMsg(Index, "You would have an advantage over that player.", Yellow)
                                    End If
                                End If
                            End If
                        End If
                    End If

                    ' Change target
                    Player(Index).Target = i
                    Player(Index).TargetType = TARGET_TYPE_PLAYER
                    Call PlayerMsg(Index, "Your target is now " & GetPlayerName(i) & ".", Yellow)
                    Exit Sub
                End If

            Next

            ' Check for an npc
            For i = 1 To MAX_MAP_NPCS

                If MapNpc(GetPlayerMap(Index), i).num > 0 Then
                    If MapNpc(GetPlayerMap(Index), i).X = X And MapNpc(GetPlayerMap(Index), i).Y = Y Then

                        ' Change target
                        Player(Index).Target = i
                        Player(Index).TargetType = TARGET_TYPE_NPC
                        Call PlayerMsg(Index, "Your target is now a " & Trim$(Npc(MapNpc(GetPlayerMap(Index), i).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If

            Next

            ' Check for an item
            For i = 1 To MAX_MAP_ITEMS

                If MapItem(GetPlayerMap(Index), i).num > 0 Then
                    If MapItem(GetPlayerMap(Index), i).X = X And MapItem(GetPlayerMap(Index), i).Y = Y Then
                        Call PlayerMsg(Index, "You see a " & Trim$(Item(MapItem(GetPlayerMap(Index), i).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If

            Next

            Exit Sub

        Case PLAYERCHAT_CHAR
            N = FindPlayer(Parse$(1))

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

        Case ACHAT_CHAR
            N = Player(Index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
                Exit Sub
            End If

            If Player(N).ChatPlayer <> Index Then
                Call PlayerMsg(Index, "Chat failed.", Pink)
                Exit Sub
            End If

            Call SendDataTo(Index, PPCHATTING_CHAR & SEP_CHAR & N & END_CHAR)
            Call SendDataTo(N, PPCHATTING_CHAR & SEP_CHAR & Index & END_CHAR)
            Exit Sub

        Case DCHAT_CHAR
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

        Case QCHAT_CHAR
            N = Player(Index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
                Exit Sub
            End If

            Call SendDataTo(Index, QCHAT_CHAR & END_CHAR)
            Call SendDataTo(N, QCHAT_CHAR & END_CHAR)
            Player(Index).ChatPlayer = 0
            Player(Index).InChat = 0
            Player(N).ChatPlayer = 0
            Player(N).InChat = 0
            Exit Sub

        Case SENDCHAT_CHAR
            N = Player(Index).ChatPlayer

            If N < 1 Then
                Call PlayerMsg(Index, "No one requested to chat with you.", Pink)
                Exit Sub
            End If

            Call SendDataTo(N, SENDCHAT_CHAR & SEP_CHAR & Parse$(1) & SEP_CHAR & Index & END_CHAR)
            Exit Sub

        Case PPTRADE_CHAR
            N = FindPlayer(Parse$(1))

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

            For i = 0 To 3

                If DirToX(GetPlayerX(Index), i) = GetPlayerX(N) And DirToY(GetPlayerY(Index), i) = GetPlayerY(N) Then

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

        Case ATRADE_CHAR
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
            For i = 0 To 3

                If DirToX(GetPlayerX(Index), i) = GetPlayerX(N) And DirToY(GetPlayerY(Index), i) = GetPlayerY(N) Then
                    Call PlayerMsg(Index, "You are trading with " & GetPlayerName(N) & "!", Pink)
                    Call PlayerMsg(N, GetPlayerName(Index) & " accepted your trade request!", Pink)
                    Call SendDataTo(Index, PPTRADING_CHAR & END_CHAR)
                    Call SendDataTo(N, PPTRADING_CHAR & END_CHAR)

                    For o = 1 To MAX_PLAYER_TRADES
                        Player(Index).Trading(o).InvNum = 0
                        Player(Index).Trading(o).InvName = vbNullString
                        Player(N).Trading(o).InvNum = 0
                        Player(N).Trading(o).InvName = vbNullString
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

        Case QTRADE_CHAR
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
            Call SendDataTo(Index, QTRADE_CHAR & END_CHAR)
            Call SendDataTo(N, QTRADE_CHAR & END_CHAR)
            Exit Sub

        Case DTRADE_CHAR
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

        Case UPDATETRADEINV_CHAR
            N = Val(Parse$(1))
            Player(Index).Trading(N).InvNum = Val(Parse$(2))
            Player(Index).Trading(N).InvName = Trim$(Parse$(3))

            If Player(Index).Trading(N).InvNum = 0 Then
                Player(Index).TradeItemMax = Player(Index).TradeItemMax - 1
                Player(Index).TradeOk = 0
                Player(N).TradeOk = 0
                Call SendDataTo(Index, TRADING_CHAR & SEP_CHAR & 0 & END_CHAR)
                Call SendDataTo(N, TRADING_CHAR & SEP_CHAR & 0 & END_CHAR)
            Else
                Player(Index).TradeItemMax = Player(Index).TradeItemMax + 1
            End If

            Call SendDataTo(Player(Index).TradePlayer, UPDATETRADEITEM_CHAR & SEP_CHAR & N & SEP_CHAR & Player(Index).Trading(N).InvNum & SEP_CHAR & Player(Index).Trading(N).InvName & END_CHAR)
            Exit Sub

        Case SWAPITEMS_CHAR
            N = Player(Index).TradePlayer

            If Player(Index).TradeOk = 0 Then
                Player(Index).TradeOk = 1
                Call SendDataTo(N, TRADING_CHAR & SEP_CHAR & 1 & END_CHAR)
            ElseIf Player(Index).TradeOk = 1 Then
                Player(Index).TradeOk = 0
                Call SendDataTo(N, TRADING_CHAR & SEP_CHAR & 0 & END_CHAR)
            End If

            If Player(Index).TradeOk = 1 And Player(N).TradeOk = 1 Then
                Player(Index).TradeItemMax2 = 0
                Player(N).TradeItemMax2 = 0

                For i = 1 To MAX_INV

                    If Player(Index).TradeItemMax = Player(Index).TradeItemMax2 Then
                        Exit For
                    End If

                    If GetPlayerInvItemNum(N, i) < 1 Then
                        Player(Index).TradeItemMax2 = Player(Index).TradeItemMax2 + 1
                    End If

                Next

                For i = 1 To MAX_INV

                    If Player(N).TradeItemMax = Player(N).TradeItemMax2 Then
                        Exit For
                    End If

                    If GetPlayerInvItemNum(Index, i) < 1 Then
                        Player(N).TradeItemMax2 = Player(N).TradeItemMax2 + 1
                    End If

                Next

                If Player(Index).TradeItemMax2 = Player(Index).TradeItemMax And Player(N).TradeItemMax2 = Player(N).TradeItemMax Then

                    For i = 1 To MAX_PLAYER_TRADES
                        For X = 1 To MAX_INV

                            If GetPlayerInvItemNum(N, X) < 1 Then
                                If Player(Index).Trading(i).InvNum > 0 Then
                                    Call GiveItem(N, GetPlayerInvItemNum(Index, Player(Index).Trading(i).InvNum), 1)
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, Player(Index).Trading(i).InvNum), 1)
                                    Exit For
                                End If
                            End If

                        Next
                    Next

                    For i = 1 To MAX_PLAYER_TRADES
                        For X = 1 To MAX_INV

                            If GetPlayerInvItemNum(Index, X) < 1 Then
                                If Player(N).Trading(i).InvNum > 0 Then
                                    Call GiveItem(Index, GetPlayerInvItemNum(N, Player(N).Trading(i).InvNum), 1)
                                    Call TakeItem(N, GetPlayerInvItemNum(N, Player(N).Trading(i).InvNum), 1)
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
                Call SendDataTo(Index, QTRADE_CHAR & END_CHAR)
                Call SendDataTo(N, QTRADE_CHAR & END_CHAR)
            End If

            Exit Sub

        Case PARTY_CHAR
            N = FindPlayer(Parse$(1))

            If N = Index Then Exit Sub
            If N > 0 Then
                If GetPlayerAccess(Index) > ADMIN_MONITER Then
                    Call PlayerMsg(Index, "You can't join a party, you are an admin!", BrightBlue)
                    Exit Sub
                End If

                If GetPlayerAccess(N) > ADMIN_MONITER Then
                    Call PlayerMsg(Index, "Admins cannot join parties!", BrightBlue)
                    Exit Sub
                End If

                If Player(N).InParty = NO Then
                    If Player(Index).PartyID > 0 Then
                        If Party(Player(Index).PartyID).Member(MAX_PARTY_MEMBERS) <> 0 Then
                            Call PlayerMsg(Index, GetPlayerName(N) & " has been invited to your party.", Pink)
                            Call PlayerMsg(N, GetPlayerName(Index) & " has invited you to join their party.  Type /join to join, or /leave to decline.", Pink)
                            Player(N).Invited = Player(Index).PartyID
                        Else
                            Call PlayerMsg(Index, "Your party is full.", Pink)
                        End If

                    Else
                        o = 0
                        i = MAX_PARTIES

                        Do While i > 0

                            If Party(i).Member(1) = 0 Then o = i
                            i = i - 1
                        Loop

                        If o = 0 Then
                            Call PlayerMsg(Index, "Party overload.", Pink)
                            Exit Sub
                        End If

                        Party(o).Member(1) = Index
                        Player(Index).InParty = YES
                        Player(Index).PartyID = o
                        Player(Index).Invited = 0
                        Call PlayerMsg(Index, "Party created.", Pink)
                        Call PlayerMsg(Index, GetPlayerName(N) & " has been invited to your party.", Pink)
                        Call PlayerMsg(N, GetPlayerName(Index) & " has invited you to join their party.  Type /join to join, or /leave to decline.", Pink)
                        Player(N).Invited = Player(Index).PartyID
                    End If

                Else
                    Call PlayerMsg(Index, "Player is already in a party.", Pink)
                End If

            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If

            Exit Sub

        Case JOINPARTY_CHAR

            If Player(Index).Invited > 0 Then
                o = 0

                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Index).Invited).Member(i) = 0 Then
                        If o = 0 Then o = i
                    End If

                Next

                If o <> 0 Then
                    Player(Index).PartyID = Player(Index).Invited
                    Player(Index).InParty = YES
                    Player(Index).Invited = 0
                    Party(Player(Index).PartyID).Member(o) = Index

                    For i = 1 To MAX_PARTY_MEMBERS

                        If Party(Player(Index).PartyID).Member(i) <> 0 And Party(Player(Index).PartyID).Member(i) <> Index Then
                            Call PlayerMsg(Party(Player(Index).PartyID).Member(i), GetPlayerName(Index) & " has joined your party!", Pink)
                        End If

                    Next

                    Call PlayerMsg(Index, "You have joined the party!", Pink)
                Else
                    Call PlayerMsg(Index, "The party is full!", Pink)
                End If

            Else
                Call PlayerMsg(Index, "You have not been invited into a party!", Pink)
            End If

            Exit Sub

        Case LEAVEPARTY_CHAR

            If Player(Index).PartyID > 0 Then
                Call PlayerMsg(Index, "You have left the party.", Pink)
                N = 0

                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Index).PartyID).Member(i) = Index Then N = i
                Next

                For i = N To MAX_PARTY_MEMBERS - 1
                    Party(Player(Index).PartyID).Member(i) = Party(Player(Index).PartyID).Member(i + 1)
                Next

                Party(Player(Index).PartyID).Member(MAX_PARTY_MEMBERS) = 0
                N = 0

                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Index).PartyID).Member(i) <> 0 And Party(Player(Index).PartyID).Member(i) <> Index Then
                        N = N + 1
                        Call PlayerMsg(Party(Player(Index).PartyID).Member(i), GetPlayerName(Index) & " has left the party.", Pink)
                    End If

                Next

                If N < 2 Then
                    Call PlayerMsg(Party(Player(Index).PartyID).Member(1), "The party has disbanded.", Pink)
                    Player(Party(Player(Index).PartyID).Member(1)).InParty = NO
                    Player(Party(Player(Index).PartyID).Member(1)).PartyID = 0
                    Party(Player(Index).PartyID).Member(1) = 0
                End If

                Player(Index).InParty = NO
                Player(Index).PartyID = 0
            Else

                If Player(Index).Invited <> 0 Then

                    For i = 1 To MAX_PARTY_MEMBERS

                        If Party(Player(Index).Invited).Member(i) <> 0 And Party(Player(Index).Invited).Member(i) <> Index Then Call PlayerMsg(Index, GetPlayerName(Index) & " has declined the invitation.", Pink)
                    Next

                    Player(Index).Invited = 0
                    Call PlayerMsg(Index, "You have declined the invitation.", Pink)
                Else
                    Call PlayerMsg(Index, "You have not been invited into a party!", Pink)
                End If
            End If

            Exit Sub

        Case PARTYCHAT_CHAR

            If Player(Index).PartyID > 0 Then

                For i = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Index).PartyID).Member(i) <> 0 Then Call PlayerMsg(Party(Player(Index).PartyID).Member(i), Parse$(1), PartyColor)
                Next

            Else
                Call PlayerMsg(Index, "You are not in a party!", Pink)
            End If

            Exit Sub

        Case GUILDCHAT_CHAR

            If GetPlayerGuild(Index) <> vbNullString Then

                For i = 1 To MAX_PLAYERS

                    If GetPlayerGuild(Index) = GetPlayerGuild(i) Then Call PlayerMsg(i, Parse$(1), GuildColor)
                Next

            Else
                Call PlayerMsg(Index, "You are not in a guild!", Pink)
            End If

            Exit Sub

        Case NEWMAIN_CHAR

            If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
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
                Print #f, Parse$(1)
                Close #f

                If SCRIPTING = 1 Then
                    Set MyScript = Nothing
                    Set clsScriptCommands = Nothing
                    Set MyScript = New clsSadScript
                    Set clsScriptCommands = New clsCommands
                    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
                    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
                    Call AddLog("Scripts reloaded.", "serverlog.txt")
                    'Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
                    Call PlayerMsg(Index, "Scripts reloaded.", White)
                End If

                Call AddLog(GetPlayerName(Index) & " updated the script.", ADMIN_LOG)
            End If

            Exit Sub

        Case REQUESTBACKUPMAIN_CHAR

            If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
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
                    Call AddLog("Scripts reloaded.", "serverlog.txt")
                    'Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
                    Call PlayerMsg(Index, "Scripts reloaded.", White)
                End If

                Call AddLog(GetPlayerName(Index) & " used the backup script.", ADMIN_LOG)
            End If

            Exit Sub

        Case SPELLS_CHAR
            Call SendPlayerSpells(Index)
            Exit Sub

        Case CAST_CHAR
            N = Val(Parse$(1))
            Call CastSpell(Index, N)
            Exit Sub

        Case REQUESTLOCATION_CHAR

            If GetPlayerAccess(Index) < ADMIN_MAPPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call PlayerMsg(Index, "Map: " & GetPlayerMap(Index) & ", X: " & GetPlayerX(Index) & ", Y: " & GetPlayerY(Index), Pink)
            Exit Sub

        Case REFRESH_CHAR
            Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), False)
            Call PlayerMsg(Index, "Map refreshed.", White)
            Exit Sub

        Case KILLPET_CHAR
            If Player(Index).Pet.Alive = YES Then
                Player(Index).Pet.Alive = NO
                Player(Index).Pet.Sprite = 0
                Call TakeFromGrid(Player(Index).Pet.Map, Player(Index).Pet.X, Player(Index).Pet.Y)
                Packet = PETDATA_CHAR & SEP_CHAR
                Packet = Packet & Index & SEP_CHAR
                Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
                Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
                Packet = Packet & Player(Index).Pet.X & SEP_CHAR
                Packet = Packet & Player(Index).Pet.Y & SEP_CHAR
                Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
                Packet = Packet & Player(Index).Pet.Sprite & SEP_CHAR
                Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
                Packet = Packet & Player(Index).Pet.Level * 5 & SEP_CHAR
                Packet = Packet & END_CHAR
                Call SendDataToMap(GetPlayerMap(Index), Packet)
            ElseIf Player(Index).Pet.Alive = NO Then
                Call PlayerMsg(Index, "You do not have a pet.", Red)
            End If
            Exit Sub

        Case PETMOVESELECT_CHAR
            X = Val(Parse$(1))
            Y = Val(Parse$(2))
            Player(Index).Pet.MapToGo = GetPlayerMap(Index)
            Player(Index).Pet.Target = 0
            Player(Index).Pet.XToGo = X
            Player(Index).Pet.YToGo = Y
            Player(Index).Pet.AttackTimer = GetTickCount

            For i = 1 To MAX_PLAYERS

                If IsPlaying(i) Then
                    If GetPlayerMap(i) = Player(Index).Pet.Map Then
                        If GetPlayerX(i) = X And GetPlayerY(i) = Y Then
                            Player(Index).Pet.TargetType = TARGET_TYPE_PLAYER
                            Player(Index).Pet.Target = i
                            Call PlayerMsg(Index, "Your pet's target is now " & Trim$(GetPlayerName(i)) & ".", Yellow)
                            Exit Sub
                        End If
                    End If
                End If

            Next

            For i = 1 To MAX_MAP_NPCS

                If MapNpc(Player(Index).Pet.Map, i).num > 0 Then
                    If MapNpc(Player(Index).Pet.Map, i).X = X And MapNpc(Player(Index).Pet.Map, i).Y = Y Then
                        Player(Index).Pet.TargetType = TARGET_TYPE_NPC
                        Player(Index).Pet.Target = i
                        Call PlayerMsg(Index, "Your pet's target is now a " & Trim$(Npc(MapNpc(Player(Index).Pet.Map, i).num).Name) & ".", Yellow)
                        Exit Sub
                    End If
                End If

            Next

            Exit Sub

        Case BUYSPRITE_CHAR

            ' Check if player stepped on sprite changing tile
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
                Call PlayerMsg(Index, "You need to be on a sprite tile to buy it!", BrightRed)
                Exit Sub
            End If

            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 = 0 Then
                Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                Call SendDataToMap(GetPlayerMap(Index), CHECKSPRITE_CHAR & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
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
                            Call SendDataToMap(GetPlayerMap(Index), CHECKSPRITE_CHAR & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
                            Call SendInventory(Index)
                        End If

                    Else

                        If GetPlayerWeaponSlot(Index) <> i And GetPlayerArmorSlot(Index) <> i And GetPlayerShieldSlot(Index) <> i And GetPlayerHelmetSlot(Index) <> i Then
                            Call SetPlayerInvItemNum(Index, i, 0)
                            Call PlayerMsg(Index, "You have bought a new sprite!", BrightGreen)
                            Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                            Call SendDataToMap(GetPlayerMap(Index), CHECKSPRITE_CHAR & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
                            Call SendInventory(Index)
                        End If
                    End If

                    If GetPlayerWeaponSlot(Index) <> i And GetPlayerArmorSlot(Index) <> i And GetPlayerShieldSlot(Index) <> i And GetPlayerHelmetSlot(Index) <> i Then
                        Exit Sub
                    End If
                End If

            Next

            Call PlayerMsg(Index, "You dont have enough to buy this sprite!", BrightRed)
            Exit Sub

        Case CHECKCOMMANDS_CHAR
            s = Parse$(1)

            If SCRIPTING = 1 Then
                PutVar App.Path & "\Scripts\Command.ini", "TEMP", "Text" & Index, Trim$(s)
                MyScript.ExecuteStatement "Scripts\Main.txt", "Commands " & Index
            Else
                Call PlayerMsg(Index, "Thats not a valid command!", 12)
            End If

            Exit Sub

        Case PROMPT_CHAR

            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerPrompt " & Index & "," & Val(Parse$(1)) & "," & Val(Parse$(2))
            End If

            Exit Sub

        Case REQUESTEDITARROW_CHAR

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, ARROWEDITOR_CHAR & END_CHAR)
            Exit Sub

        Case EDITARROW_CHAR

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse$(1))

            If N < 0 Or N > MAX_ARROWS Then
                Call HackingAttempt(Index, "Invalid arrow Index")
                Exit Sub
            End If

            Call AddLog(GetPlayerName(Index) & " editing arrow #" & N & ".", ADMIN_LOG)
            Call SendEditArrowTo(Index, N)
            Exit Sub

        Case SAVEARROW_CHAR

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse$(1))

            If N < 0 Or N > MAX_ITEMS Then
                Call HackingAttempt(Index, "Invalid arrow Index")
                Exit Sub
            End If

            Arrows(N).Name = Parse$(2)
            Arrows(N).Pic = Val(Parse$(3))
            Arrows(N).Range = Val(Parse$(4))
            Call SendUpdateArrowToAll(N)
            Call SaveArrow(N)
            Call AddLog(GetPlayerName(Index) & " saved arrow #" & N & ".", ADMIN_LOG)
            Exit Sub

        Case CHECKARROWS_CHAR
            N = Arrows(Val(Parse$(1))).Pic
            Call SendDataToMap(GetPlayerMap(Index), CHECKARROWS_CHAR & SEP_CHAR & Index & SEP_CHAR & N & END_CHAR)
            Exit Sub

        Case SPEECHSCRIPT_CHAR

            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Index & "," & Parse$(1)
            End If

            Exit Sub

        Case REQUESTEDITSPEECH_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, SPEECHEDITOR_CHAR & END_CHAR)
            Exit Sub

        Case EDITSPEECH_CHAR

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse$(1))

            If N < 0 Or N > MAX_SPEECH Then
                Call HackingAttempt(Index, "Invalid Speech Index")
                Exit Sub
            End If

            Call AddLog(GetPlayerName(Index) & " editing speech #" & N & ".", ADMIN_LOG)
            Call SendEditSpeechTo(Index, N)
            Exit Sub

        Case SAVESPEECH_CHAR

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse$(1))

            If N < 0 Or N > MAX_SPEECH Then
                Call HackingAttempt(Index, "Invalid Speech Index")
                Exit Sub
            End If

            Speech(N).Name = Parse$(2)
            Dim p As Long

            p = 3

            For i = 0 To MAX_SPEECH_OPTIONS
                Speech(N).num(i).Exit = Val(Parse$(p))
                Speech(N).num(i).text = Parse$(p + 1)
                Speech(N).num(i).SaidBy = Val(Parse$(p + 2))
                Speech(N).num(i).Respond = Val(Parse$(p + 3))
                Speech(N).num(i).Script = Val(Parse$(p + 4))
                p = p + 5

                For o = 1 To 3
                    Speech(N).num(i).Responces(o).Exit = Val(Parse$(p))
                    Speech(N).num(i).Responces(o).GoTo = Val(Parse$(p + 1))
                    Speech(N).num(i).Responces(o).text = Parse$(p + 2)
                    p = p + 3
                Next
            Next

            Call SaveSpeech(N)
            Call SendSpeechToAll(N)
            Call AddLog(GetPlayerName(Index) & " saved speech #" & N & ".", ADMIN_LOG)
            Exit Sub

        Case NEEDSPEECH_CHAR
            Call SendSpeechTo(Index, Val(Parse$(1)))
            Exit Sub

        Case REQUESTEDITEMOTICON_CHAR

            ' Prevent hacking
            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            Call SendDataTo(Index, EMOTICONEDITOR_CHAR & END_CHAR)
            Exit Sub

        Case EDITEMOTICON_CHAR

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse$(1))

            If N < 0 Or N > MAX_EMOTICONS Then
                Call HackingAttempt(Index, "Invalid Emoticon Index")
                Exit Sub
            End If

            Call AddLog(GetPlayerName(Index) & " editing emoticon #" & N & ".", ADMIN_LOG)
            Call SendEditEmoticonTo(Index, N)
            Exit Sub

        Case SAVEEMOTICON_CHAR

            If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
                Call HackingAttempt(Index, "Admin Cloning")
                Exit Sub
            End If

            N = Val(Parse$(1))

            If N < 0 Or N > MAX_EMOTICONS Then
                Call HackingAttempt(Index, "Invalid Emoticon Index")
                Exit Sub
            End If

            Emoticons(N).Type = Val(Parse$(2))
            Emoticons(N).Command = Parse$(3)
            Emoticons(N).Pic = Val(Parse$(4))
            Emoticons(N).sound = Parse$(5)
            Call SendUpdateEmoticonToAll(N)
            Call SaveEmoticon(N)
            Call AddLog(GetPlayerName(Index) & " saved emoticon #" & N & ".", ADMIN_LOG)
            Exit Sub

        Case CHECKEMOTICONS_CHAR
            Call SendDataToMap(GetPlayerMap(Index), CHECKEMOTICONS_CHAR & SEP_CHAR & Index & SEP_CHAR & Emoticons(Val(Parse$(1))).Type & SEP_CHAR & Emoticons(Val(Parse$(1))).Pic & SEP_CHAR & Emoticons(Val(Parse$(1))).sound & END_CHAR)
            Exit Sub

        Case MAPREPORT_CHAR
            Packs = MAPREPORT_CHAR & SEP_CHAR

            For i = 1 To MAX_MAPS
                Packs = Packs & Map(i).Name & SEP_CHAR
            Next

            Packs = Packs & END_CHAR
            Call SendDataTo(Index, Packs)
            Exit Sub

        Case GMTIME_CHAR
            GameTime = Val(Parse$(1))
            Call SendTimeToAll
            Exit Sub

        Case WEATHER_CHAR
            GameWeather = Val(Parse$(1))
            Call SendWeatherToAll
            Exit Sub

        Case WARPTO_CHAR
            Call PlayerWarp(Index, Val(Parse$(1)), GetPlayerX(Index), GetPlayerY(Index))
            Exit Sub

        Case WARPTOME_CHAR
            N = FindPlayer(Parse$(1))

            If N > 0 Then
                Call PlayerWarp(N, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Else
                Call PlayerMsg(Index, "Player not online!", BrightRed)
            End If

            Exit Sub

        Case WARPPLAYER_CHAR

            If Val(Parse$(1)) > MAX_MAPS Or Val(Parse$(1)) < 1 Then
                If FindPlayer(Trim$(Parse$(1))) <> 0 Then
                    Call PlayerWarp(Index, GetPlayerMap(FindPlayer(Trim$(Parse$(1)))), GetPlayerX(FindPlayer(Trim$(Parse$(1)))), GetPlayerY(FindPlayer(Trim$(Parse$(1)))))

                    If Player(Index).Pet.Alive = YES Then
                        Player(Index).Pet.Map = GetPlayerMap(Index)
                        Player(Index).Pet.X = GetPlayerX(Index)
                        Player(Index).Pet.Y = GetPlayerY(Index)
                        Player(Index).Pet.MapToGo = -1
                        Player(Index).Pet.XToGo = -1
                        Player(Index).Pet.YToGo = -1
                    End If

                Else
                    Call PlayerMsg(Index, "'" & Parse$(1) & "' is not a valid map number or an online player's name!", BrightRed)
                    Exit Sub
                End If

            Else
                Call PlayerWarp(Index, Val(Parse$(1)), GetPlayerX(Index), GetPlayerY(Index))

                If Player(Index).Pet.Alive = YES Then
                    Player(Index).Pet.Map = GetPlayerMap(Index)
                    Player(Index).Pet.X = GetPlayerX(Index)
                    Player(Index).Pet.Y = GetPlayerY(Index)
                    Player(Index).Pet.MapToGo = -1
                    Player(Index).Pet.XToGo = -1
                    Player(Index).Pet.YToGo = -1
                End If
            End If

            Exit Sub

        Case ARROWHIT_CHAR
            N = Val(Parse$(1))
            z = Val(Parse$(2))
            X = Val(Parse$(3))
            Y = Val(Parse$(4))

            If N = TARGET_TYPE_PLAYER Then

                ' Make sure we dont try to attack ourselves
                If z <> Index Then

                    ' Can we attack the player?
                    If CanAttackPlayerWithArrow(Index, z) Then
                        If Not CanPlayerBlockHit(z) Then

                            ' Get the damage we can do
                            If Not CanPlayerCriticalHit(Index) Then
                                Damage = GetPlayerDamage(Index) - GetPlayerProtection(z) + (Rnd * 5) - 2
                                'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & END_CHAR)
                                Call SendSound(Index, ATTACK_SOUND, SDTM)
                            Else
                                N = GetPlayerDamage(Index)
                                Damage = N + Int(Rnd * Int(N / 2)) + 1 - GetPlayerProtection(z) + (Rnd * 5) - 2
                                Call BattleMsg(Index, "You feel a surge of energy upon shooting!", BrightCyan, 0)
                                Call BattleMsg(z, GetPlayerName(Index) & " shoots with amazing accuracy!", BrightCyan, 1)

                                'Call PlayerMsg(index, "You feel a surge of energy upon shooting!", BrightCyan)
                                'Call PlayerMsg(z, GetPlayerName(index) & " shoots with amazing accuracy!", BrightCyan)
                                'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Blow0" & END_CHAR)
                                Call SendSound(Index, CRITICALHIT_SOUND, SDTM)
                            End If

                            If Damage > 0 Then
                                Call AttackPlayer(Index, z, Damage)
                            Else
                                Call BattleMsg(Index, "Your attack does nothing.", BrightRed, 0)
                                Call BattleMsg(z, GetPlayerName(z) & "'s attack did nothing.", BrightRed, 1)

                                'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                                'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "miss" & END_CHAR)
                                Call SendSound(Index, MISS_SOUND, SDTM)
                            End If

                        Else
                            Call BattleMsg(Index, GetPlayerName(z) & " blocked your hit!", BrightCyan, 0)
                            Call BattleMsg(z, "You blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan, 1)

                            'Call PlayerMsg(index, GetPlayerName(z) & "'s " & Trim$(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked your hit!", BrightCyan)
                            'Call PlayerMsg(z, "Your " & Trim$(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                            'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "miss" & END_CHAR)
                            Call SendSound(Index, MISS_SOUND, SDTM)
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
                        'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Blow" & Int(Rnd * 7) + 1 & END_CHAR)
                        Call SendSound(Index, ATTACK_SOUND, SDTM)
                    Else
                        N = GetPlayerDamage(Index)
                        Damage = N + Int(Rnd * Int(N / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), z).num).DEF / 2) + (Rnd * 5) - 2
                        Call BattleMsg(Index, "You feel a surge of energy upon shooting!", BrightCyan, 0)

                        'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                        'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Blow0" & END_CHAR)
                        Call SendSound(Index, CRITICALHIT_SOUND, SDTM)
                    End If

                    If Damage > 0 Then
                        Call AttackNpc(Index, z, Damage)
                        Call SendDataTo(Index, BLITPLAYERDMG_CHAR & SEP_CHAR & Damage & SEP_CHAR & z & END_CHAR)
                    Else
                        Call BattleMsg(Index, "Your attack does nothing.", BrightRed, 0)

                        'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                        Call SendDataTo(Index, BLITPLAYERDMG_CHAR & SEP_CHAR & Damage & SEP_CHAR & z & END_CHAR)
                        'Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "miss" & END_CHAR)
                        Call SendSound(Index, MISS_SOUND, SDTM)
                    End If

                    Exit Sub
                End If
            End If

            Exit Sub
            
        Case FORGETSPELL_CHAR
                ' Spell slot
                N = Val(Parse$(1))

                ' Prevent subscript out of range
                If N <= 0 Or N > MAX_PLAYER_SPELLS Then
                    Call HackingAttempt(Index, "Invalid Spell Slot")
                    Exit Sub
                End If

                If IsPlaying(Index) = True Then
                    With Player(Index).Char(Player(Index).CharNum)
                        If .Spell(N) = 0 Then
                            Call PlayerMsg(Index, "No spell here.", Red)
                        Else
                            Call PlayerMsg(Index, "You have forgotten the spell, " & Trim$(Spell(.Spell(N)).Name) & ".", Green)
                            .Spell(N) = 0
                            Call SendSpells(Index)
                        End If
                    End With
                End If
                Exit Sub
                
        Case RETURNSCRIPT_CHAR
                
                If Index <= 0 Then Exit Sub
                
                If IsPlaying(Index) = False Then Exit Sub
                
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnReturn " & Index
                Exit Sub
                
        'I think this is causing errors, so I did something else about it
        'Case CLOSINGDOWN_CHAR
        '        If Index > 0 Then
        '            Call CloseSocket(Index)
        '        Else
        '            Call AddLog("Avoided error in CLOSINGDOWN_CHAR packet, Index: " & Index, "errorlist.txt")
        '        End If
        '        Exit Sub

    End Select

    Call HackingAttempt(Index, "Invalid packet. (" & Parse$(0) & ")")
    Exit Sub
    
ErrHandler:
    Call AddLog("Error while running Sub HandleData(" & Index & ", " & Data & ") !", "errorlist.txt")
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
    Dim Buffer As String
    Dim Packet As String
    Dim top As String * 3
    Dim Start As Long

    If Index > 0 Then
        
        frmServer.Socket(Index).GetData Buffer, vbString, DataLength
        
        DataKBIn = Len(Buffer)
        Call UpdateCaption

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
    Else
        Call AddLog("Avoided error in IncomingData(" & Index & "," & DataLength & ")", "errorlist.txt")
    End If

End Sub

Function IsBanned(ByVal ip As String) As Boolean
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
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(ip, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If

    Loop

    Close #f
End Function

Public Sub MakeGuild(ByRef GuildName As String)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\Accounts\Guilds.txt"
    f = FreeFile

    Open FileName For Append As #f
        Print #f, GuildName
    Close #f

End Sub

Function GuildExist(ByRef GuildName As String) As Byte
Dim FileName As String, G As String
Dim f As Long

    GuildExist = NO
    
    FileName = App.Path & "\Accounts\Guilds.txt"
    
    f = FreeFile
    Open FileName For Input As #f
    
    Do While Not EOF(f)
        Input #f, G
        
        ' Is made?
        If Trim$(LCase$(G)) = Trim$(LCase$(GuildName)) Then
            GuildExist = YES
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

    Next

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

    If Index <= 0 Or Index > MAX_PLAYERS Then
        IsPlaying = False
        Call AddLog("Avoided error in Function IsPlaying(" & Index & ")", "errorlist.txt")
        Exit Function
    End If

    If IsConnected(Index) And Player(Index).InGame = "YES" Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If

End Function

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Packet As String

    Packet = MAPMSG_CHAR & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal Index As Long)
    Dim Packet As String

    Packet = MAPMSG2_CHAR & SEP_CHAR & Msg & SEP_CHAR & Index & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub PlainMsg(ByVal Index As Long, ByVal Msg As String, ByVal num As Long)
    Dim Packet As String

    Packet = PLAINMSG_CHAR & SEP_CHAR & Msg & SEP_CHAR & num & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Packet As String

    Packet = PLAYERMSG_CHAR & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendArrows(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ARROWS
        Call SendUpdateArrowTo(Index, i)
    Next

End Sub

Sub SendChars(ByVal Index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = ALLCHARS_CHAR & SEP_CHAR

    For i = 1 To MAX_CHARS
        Packet = Packet & Trim$(Player(Index).Char(i).Name) & SEP_CHAR & Trim$(Class(Player(Index).Char(i).Class).Name) & SEP_CHAR & Player(Index).Char(i).Level & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendClasses(ByVal Index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = CLASSESDATA_CHAR & SEP_CHAR & Max_Classes & SEP_CHAR

    For i = 1 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)

    DataKBOut = Len(Data)

    If IsConnected(Index) Then
        frmServer.Socket(Index).SendData Data

        DoEvents
    End If
    
    Call UpdateCaption

End Sub

Sub SendDataToAll(ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> Index Then
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

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendEditArrowTo(ByVal Index As Long, ByVal EmoNum As Long)
    Dim Packet As String

    Packet = EDITARROW_CHAR & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Arrows(EmoNum).Name) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditEmoticonTo(ByVal Index As Long, ByVal EmoNum As Long)
    Dim Packet As String

    Packet = EDITEMOTICON_CHAR & SEP_CHAR & EmoNum & SEP_CHAR & Emoticons(EmoNum).Type & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & Emoticons(EmoNum).sound & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim Packet As String

    Packet = EDITITEM_CHAR & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
    Dim Packet As String
    Dim i As Long
    
    'Packet = EDITNPC_CHAR & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR
    'Packet = Packet & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).str & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & END_CHAR
    Packet = EDITNPC_CHAR & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).Speed & SEP_CHAR & Npc(NpcNum).Magi & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR & Npc(NpcNum).Speech & SEP_CHAR

    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(NpcNum).ItemNPC(i).Chance
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
    Dim Packet As String
    Dim i As Long, z As Long

    Packet = EDITSHOP_CHAR & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR

    For i = 1 To 6
        For z = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(i).Value(z).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GetValue & SEP_CHAR
        Next
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSpeechTo(ByVal Index As Long, ByVal SpcNum As Long)
    Dim Packet As String
    Dim i, o As Long

    Packet = EDITSPEECH_CHAR & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR

    For i = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(i).Exit & SEP_CHAR & Speech(SpcNum).num(i).text & SEP_CHAR & Speech(SpcNum).num(i).SaidBy & SEP_CHAR & Speech(SpcNum).num(i).Respond & SEP_CHAR & Speech(SpcNum).num(i).Script & SEP_CHAR

        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(i).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).text & SEP_CHAR
        Next
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim Packet As String

    Packet = EDITSPELL_CHAR & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).sound & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEmoticons(ByVal Index As Long)
    Dim i As Long

    For i = 0 To MAX_EMOTICONS

        If Trim$(Emoticons(i).Command) <> vbNullString Then
            Call SendUpdateEmoticonTo(Index, i)
        End If

    Next

End Sub

Sub SendFriendListTo(ByVal Index As Long)
    Dim Packet As String
    Dim N As Long

    Packet = FRIENDLIST_CHAR & SEP_CHAR

    For N = 1 To MAX_FRIENDS

        If FindPlayer(Player(Index).Char(Player(Index).CharNum).Friends(N)) And Player(Index).Char(Player(Index).CharNum).Friends(N) <> vbNullString Then
            Packet = Packet & Player(Index).Char(Player(Index).CharNum).Friends(N) & SEP_CHAR
        End If

    Next

    Packet = Packet & NEXT_CHAR & SEP_CHAR

    For N = 1 To MAX_FRIENDS
        Packet = Packet & Player(Index).Char(Player(Index).CharNum).Friends(N) & SEP_CHAR
    Next

    Packet = Packet & NEXT_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
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

Sub SendHP(ByVal Index As Long)
    Dim Packet As String

    Packet = PLAYERHP_CHAR & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
    Packet = PLAYERPOINTS_CHAR & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendInfo(ByVal Index As Long)
    Dim Packet As String

    Packet = INFO_CHAR & SEP_CHAR & TotalOnlinePlayers & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendInventory(ByVal Index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = PLAYERINV_CHAR & SEP_CHAR & Index & SEP_CHAR

    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, i) & SEP_CHAR & GetPlayerInvItemValue(Index, i) & SEP_CHAR & GetPlayerInvItemDur(Index, i) & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Dim Packet As String

    Packet = PLAYERINVUPDATE_CHAR & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & Index & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendItems(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Trim$(Item(i).Name) <> vbNullString Then
            Call SendUpdateItemTo(Index, i)
        End If

    Next

End Sub

Sub SendJoinMap(ByVal Index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = vbNullString

    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> Index And GetPlayerMap(i) = GetPlayerMap(Index) Then
            Packet = PLAYERDATA_CHAR & SEP_CHAR
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
            Packet = Packet & END_CHAR
            Call SendDataTo(Index, Packet)

            If Player(i).Pet.Alive = YES Then
                Packet = PETDATA_CHAR & SEP_CHAR
                Packet = Packet & i & SEP_CHAR
                Packet = Packet & Player(i).Pet.Alive & SEP_CHAR
                Packet = Packet & Player(i).Pet.Map & SEP_CHAR
                Packet = Packet & Player(i).Pet.X & SEP_CHAR
                Packet = Packet & Player(i).Pet.Y & SEP_CHAR
                Packet = Packet & Player(i).Pet.Dir & SEP_CHAR
                Packet = Packet & Player(i).Pet.Sprite & SEP_CHAR
                Packet = Packet & Player(i).Pet.HP & SEP_CHAR
                Packet = Packet & Player(i).Pet.Level * 5 & SEP_CHAR
                Packet = Packet & END_CHAR
                Call SendDataTo(Index, Packet)
            End If
        End If

    Next

    ' Send index's player data to everyone on the map including himself
    Packet = PLAYERDATA_CHAR & SEP_CHAR
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
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)

    If Player(Index).Pet.Alive = YES Then
        Packet = PETDATA_CHAR & SEP_CHAR
        Packet = Packet & Index & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Index).Pet.X & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Y & SEP_CHAR
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

    Packet = PLAYERDATA_CHAR & SEP_CHAR
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
    Packet = Packet & END_CHAR
    Call SendDataToMapBut(Index, MapNum, Packet)

    If Player(Index).Pet.Alive = YES Then
        Packet = PETDATA_CHAR & SEP_CHAR
        Packet = Packet & Index & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Index).Pet.X & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Y & SEP_CHAR
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

    Packet = PLAYERDATA_CHAR & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & vbNullString & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & vbNullString & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToAllBut(Index, Packet)
    Packet = PETDATA_CHAR & SEP_CHAR
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
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim o As Long
    Dim p1 As String, p2 As String

    Packet = MAPDATA_CHAR & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX

            With Map(MapNum).Tile(X, Y)
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
                If .String1 <> vbNullString Then i = 12
                If .String2 <> vbNullString Then i = 13
                If .String3 <> vbNullString Then i = 14
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
        Packet = Packet & Map(MapNum).NpcSpawn(X).Used & SEP_CHAR & Map(MapNum).NpcSpawn(X).X & SEP_CHAR & Map(MapNum).NpcSpawn(X).Y & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    X = Int(Len(Packet) / 2)
    p1 = Mid$(Packet, 1, X)
    p2 = Mid$(Packet, X + 1, Len(Packet) - X)
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim Packet As String
    Dim i As Long

    Packet = MAPITEMDATA_CHAR & SEP_CHAR

    For i = 1 To MAX_MAP_ITEMS

        If MapNum > 0 Then
            Packet = Packet & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).Y & SEP_CHAR
        End If

    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
    Dim Packet As String
    Dim i As Long

    Packet = MAPITEMDATA_CHAR & SEP_CHAR

    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).Y & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
    Dim Packet As String
    Dim i As Long

    Packet = MAPNPCDATA_CHAR & SEP_CHAR

    For i = 1 To MAX_MAP_NPCS

        If MapNum > 0 Then
            Packet = Packet & MapNpc(MapNum, i).num & SEP_CHAR & MapNpc(MapNum, i).X & SEP_CHAR & MapNpc(MapNum, i).Y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
        End If

    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMP(ByVal Index As Long)
    Dim Packet As String

    Packet = PLAYERMP_CHAR & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
    Dim i As Long
    Dim Packet As String

    Packet = NEWCHARCLASSES_CHAR & SEP_CHAR & Max_Classes & SEP_CHAR

    For i = 1 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).MaleSprite & SEP_CHAR & Class(i).FemaleSprite & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Trim$(Npc(i).Name) <> vbNullString Then
            Call SendUpdateNpcTo(Index, i)
        End If

    Next

End Sub

Sub SendOnlineList()
    Dim Packet As String
    Dim i As Long
    Dim N As Long

    Packet = vbNullString
    N = 0

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Packet = Packet & SEP_CHAR & GetPlayerName(i) & SEP_CHAR
            N = N + 1
        End If

    Next

    Packet = ONLINELIST_CHAR & SEP_CHAR & N & Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim Packet As String

    ' Send index's player data to everyone including himself on the map
    Packet = PLAYERDATA_CHAR & SEP_CHAR
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
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)

    If Player(Index).Pet.Alive = YES Then
        Packet = PETDATA_CHAR & SEP_CHAR
        Packet = Packet & Index & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Index).Pet.X & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Y & SEP_CHAR
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
    Dim i As Long

    Packet = SPELLS_CHAR & SEP_CHAR

    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(Index, i) & SEP_CHAR
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerXY(ByVal Index As Long)
    Dim Packet As String

    Packet = PLAYERXY_CHAR & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendShops(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Trim$(Shop(i).Name) <> vbNullString Then
            Call SendUpdateShopTo(Index, i)
        End If

    Next

End Sub

Sub SendSP(ByVal Index As Long)
    Dim Packet As String

    Packet = PLAYERSP_CHAR & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendSpeech(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SPEECH

        If Trim$(Speech(i).Name) <> vbNullString Then
            Call SendSpeechTo(Index, i)
        End If

    Next

End Sub

Sub SendSpeechTo(ByVal Index As Long, ByVal SpcNum As Long)
    Dim Packet As String
    Dim i, o As Long

    Packet = SPEECH_CHAR & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR

    For i = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(i).Exit & SEP_CHAR & Speech(SpcNum).num(i).text & SEP_CHAR & Speech(SpcNum).num(i).SaidBy & SEP_CHAR & Speech(SpcNum).num(i).Respond & SEP_CHAR & Speech(SpcNum).num(i).Script & SEP_CHAR

        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(i).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).text & SEP_CHAR
        Next
    Next

    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendSpeechToAll(ByVal SpcNum As Long)
    Dim Packet As String
    Dim i, o As Long

    Packet = SPEECH_CHAR & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR

    For i = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).num(i).Exit & SEP_CHAR & Speech(SpcNum).num(i).text & SEP_CHAR & Speech(SpcNum).num(i).SaidBy & SEP_CHAR & Speech(SpcNum).num(i).Respond & SEP_CHAR & Speech(SpcNum).num(i).Script & SEP_CHAR

        For o = 1 To 3
            Packet = Packet & Speech(SpcNum).num(i).Responces(o).Exit & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).GoTo & SEP_CHAR & Speech(SpcNum).num(i).Responces(o).text & SEP_CHAR
        Next
    Next

    Packet = Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Trim$(Spell(i).Name) <> vbNullString Then
            Call SendUpdateSpellTo(Index, i)
        End If

    Next

End Sub

Sub SendStats(ByVal Index As Long)
    Dim Packet As String

    Packet = PLAYERSTATSPACKET_CHAR & SEP_CHAR & GetPlayerstr(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTimeTo(ByVal Index As Long)
    Dim Packet As String

    Packet = TIME_CHAR & SEP_CHAR & GameTime & END_CHAR
    Call SendDataTo(Index, Packet)
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

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
    Dim Packet As String
    Dim i As Long, X As Long, Y As Long, z As Long, XX As Long

    z = 0
    Packet = TRADE_CHAR & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR

    For i = 1 To 6
        For XX = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(i).Value(XX).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GetValue & SEP_CHAR

            ' Item #
            X = Shop(ShopNum).TradeItem(i).Value(XX).GetItem

            If Item(X).Type = ITEM_TYPE_SPELL Then

                ' Spell class requirement
                Y = Spell(Item(X).Data1).ClassReq

                If Y = 0 Then
                    Call PlayerMsg(Index, Trim$(Item(X).Name) & " can be used by all classes.", Yellow)
                Else
                    Call PlayerMsg(Index, Trim$(Item(X).Name) & " can only be used by a " & GetClassName(Y) & ".", Yellow)
                End If
            End If

            If X < 1 Then
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

    Packet = UPDATEARROW_CHAR & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
    Dim Packet As String

    Packet = UPDATEARROW_CHAR & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateEmoticonTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim Packet As String

    Packet = UPDATEEMOTICON_CHAR & SEP_CHAR & ItemNum & SEP_CHAR & Emoticons(ItemNum).Type & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & Emoticons(ItemNum).sound & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
    Dim Packet As String

    Packet = UPDATEEMOTICON_CHAR & SEP_CHAR & ItemNum & SEP_CHAR & Emoticons(ItemNum).Type & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & Emoticons(ItemNum).sound & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
    Dim Packet As String

    'Packet = UPDATEITEM_CHAR & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & END_CHAR
    Packet = UPDATEITEM_CHAR & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim Packet As String

    'Packet = UPDATEITEM_CHAR & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & END_CHAR
    Packet = UPDATEITEM_CHAR & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
    Dim Packet As String

    Packet = UPDATENPC_CHAR & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Speech & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
    Dim Packet As String

    Packet = UPDATENPC_CHAR & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Speech & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum)
    Dim Packet As String

    Packet = UPDATESHOP_CHAR & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
    Dim Packet As String

    Packet = UPDATESHOP_CHAR & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim Packet As String

    Packet = UPDATESPELL_CHAR & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim Packet As String

    Packet = UPDATESPELL_CHAR & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendWeatherTo(ByVal Index As Long)
    Dim Packet As String

    If RainIntensity <= 0 Then RainIntensity = 1
    Packet = WEATHER_CHAR & SEP_CHAR & GameWeather & SEP_CHAR & RainIntensity & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWeatherToAll()
    Dim i As Long
    Dim Weather As String

    'Select Case GameWeather

    '    Case 0
    '        Weather = "None"

    '    Case 1
    '        Weather = "Rain"

    '    Case 2
    '        Weather = "Snow"

    '    Case 3
    '        Weather = "Thunder"
    'End Select

    'frmServer.Label5.Caption = "Current Weather: " & Weather

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SendWeatherTo(i)
        End If

    Next

End Sub

Sub SendWhosOnline(ByVal Index As Long)
    Dim s As String
    Dim N As Long, i As Long

    s = vbNullString
    N = 0

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> Index Then
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

    Call PlayerMsg(Index, s, WhoColor)
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim Packet As String

    If IsPlaying(Index) Then
        Packet = PLAYERWORNEQ_CHAR & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & END_CHAR
        Call SendDataToMap(GetPlayerMap(Index), Packet)
    End If

End Sub

Sub SocketConnected(ByVal Index As Long)

    If Index > 0 Then

        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
        If Not IsBanned(GetPlayerIP(Index)) Then
            Call AddLog("Received connection from " & GetPlayerIP(Index) & ".", "serverlog.txt")
            'Call TextAdd(frmServer.txtText(0), "Received connection from " & GetPlayerIP(Index) & ".", True)
        Else
            Call AlertMsg(Index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
        End If

        'Else
        ' Tried multiple connections
        '    Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
        'End If
    Else
        Call AddLog("Avoided error in Sub SocketConnected(" & Index & ")", "errorlist.txt")
    End If

End Sub

Sub UpdateCaption()
    'frmServer.Caption = GAME_NAME & " - Server - Powered By Elysium Source"
    'frmServer.lblIP.Caption = "IP Address: " & GetIP
    'frmServer.lblPort.Caption = "Port: " & STR(frmServer.Socket(0).LocalPort)
    'frmServer.TPO.Caption = "Total Players Online: " & TotalOnlinePlayers
    
    TrayModify ToolTip, Server_BuildToolTipString
    
    'Exit Sub
End Sub

' SAFE MODE -- Uncomment for ON, comment for OFF (whole function)
'Function Parse$(ByVal index As Long) As String
'    If index > NumParse Then
'        Call HackingAttempt(ParseIndex, "Subscript out of range, " & ZePacket(0))
'        Exit Function
'    End If
'
'    Parse = ZePacket(index)

'End Function
