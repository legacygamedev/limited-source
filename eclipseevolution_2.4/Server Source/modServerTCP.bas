Attribute VB_Name = "modServerTCP"
Option Explicit


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

  Dim packet As String
  Dim i As Long

    packet = PacketID.AdminMsg & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, packet)
        End If

    Next i

End Sub

Sub AlertMsg(ByVal index As Long, ByVal Msg As String)

  Dim packet As String

    packet = PacketID.AlertMsg & SEP_CHAR & Msg & SEP_CHAR & END_CHAR

    Call SendDataTo(index, packet)
    Call CloseSocket(index)

End Sub


' Used for checking if a port is open, and opens if not
Sub CheckOpenPort(ByVal Port As Integer)

    On Error GoTo Err:

    'Check if ports.ini exists

    If VarExists(App.Path & "\Ports.ini", "PORTS", STR(Port)) Then
        If Val(GetVar(App.Path & "\Ports.ini", "PORTS", STR(Port))) = 1 Then
            Exit Sub
        End If

    End If

    'Prompy user

    If MsgBox("Would you like Eclipse to open a port in your firewall to run the server? People may not be able to connect if the port is closed.", vbYesNo) = vbYes Then
        'Use the shell to open the port
        Shell ("netsh firewall add portopening TCP " & Port & " EclipseEvolution-ServerPort")
        Call MsgBox("Port " & Port & " opened.", vbOKOnly, "Success!")
     Else
        'Keep going
        Call MsgBox("No action taken. You can open the port manually at a later date. This will not prompt you again.", vbOKOnly)
    End If

    'Write value to ports.ini
    Call PutVar(App.Path & "\Ports.ini", "PORTS", STR(Port), 1)
    Exit Sub

Err:
    Call MsgBox("Error occured - " & Err.Description & ". Running normally.", vbCritical)
    Err.Clear
    Call PutVar(App.Path & "\Ports.ini", "PORTS", STR(Port), 1)

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

Sub DisabledTime()

  Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call DisabledTimeTo(i)
        End If

    Next i

End Sub

Sub DisabledTimeTo(ByVal index As Long)

  Dim packet As String

    packet = PacketID.DTime & SEP_CHAR & TimeDisable & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)

  Dim packet As String

    packet = PacketID.GlobalMsg & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR

    Call SendDataToAll(packet)

End Sub

Sub GrapleHook(ByVal index As Long)

  Dim X As Long
  Dim Y As Long
  Dim MapNum As Long

    MapNum = GetPlayerMap(index)
    If Player(index).HookShotX <> 0 Or Player(index).HookShotY <> 0 Then
        If Player(index).locked = True Then
            Call PlayerMsg(index, "You can only fire one grappleshot at the time", 1)
            Exit Sub
        End If

    End If

    Player(index).locked = True
    Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)

    If GetPlayerDir(index) = DIR_DOWN Then
        X = GetPlayerX(index)
        Y = GetPlayerY(index) + 1

        Do While Y <= MAX_MAPY

            If Map(MapNum).tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Player(index).HookShotX = X
                Player(index).HookShotY = Y
                Exit Sub
             Else

                If Map(MapNum).tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                    Player(index).HookShotX = X
                    Player(index).HookShotY = Y
                    Exit Sub
                End If

            End If
            Y = Y + 1
        Loop

        Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Player(index).HookShotX = X
        Player(index).HookShotY = Y
        Exit Sub
    End If

    If GetPlayerDir(index) = DIR_UP Then
        X = GetPlayerX(index)
        Y = GetPlayerY(index) - 1

        Do While Y >= 0

            If Map(MapNum).tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Player(index).HookShotX = X
                Player(index).HookShotY = Y
                Exit Sub
             Else

                If Map(MapNum).tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                    Player(index).HookShotX = X
                    Player(index).HookShotY = Y
                    Exit Sub
                End If

            End If
            Y = Y - 1
        Loop

        Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Player(index).HookShotX = X
        Player(index).HookShotY = Y
        Exit Sub
    End If

    If GetPlayerDir(index) = DIR_RIGHT Then
        X = GetPlayerX(index) + 1
        Y = GetPlayerY(index)

        Do While X <= MAX_MAPX

            If Map(MapNum).tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Player(index).HookShotX = X
                Player(index).HookShotY = Y
                Exit Sub
             Else

                If Map(MapNum).tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                    Player(index).HookShotX = X
                    Player(index).HookShotY = Y
                    Exit Sub
                End If

            End If
            X = X + 1
        Loop

        Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Player(index).HookShotX = X
        Player(index).HookShotY = Y
        Exit Sub
    End If

    If GetPlayerDir(index) = DIR_LEFT Then
        X = GetPlayerX(index) - 1
        Y = GetPlayerY(index)

        Do While X >= 0

            If Map(MapNum).tile(X, Y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                Player(index).HookShotX = X
                Player(index).HookShotY = Y
                Exit Sub
             Else

                If Map(MapNum).tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                    Player(index).HookShotX = X
                    Player(index).HookShotY = Y
                    Exit Sub
                End If

            End If
            X = X - 1
        Loop

        Call SendDataToMap(GetPlayerMap(index), PacketID.HookShot & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Player(index).HookShotX = X
        Player(index).HookShotY = Y
        Exit Sub
    End If

End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)

    Exit Sub

    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & ".")
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
  Dim MsgTo As Long
  Dim Dir As Long
  Dim InvNum As Long
  Dim Amount As Long
  Dim Damage As Long
  Dim PointType As Long
  Dim Movement As Long
  Dim i As Long
  Dim n As Long
  Dim X As Long
  Dim Y As Long
  Dim f As Long
  Dim MapNum As Long
  Dim s As String
  Dim ShopNum As Long
  Dim ItemNum As Long
  Dim DurNeeded As Long
  Dim GoldNeeded As Long
  Dim z As Long
  Dim BX As Long
  Dim BY As Long
  Dim TempVal As Long
  Dim hFile
  Dim m As Long
  Dim j As Long
  Dim Imail As String
  Dim surge As String

    'Handle the data
    Parse = Split(Data, SEP_CHAR)
    
    '//!! Temp error checking
    If Len(Parse(0)) > 1 Then
        MsgBox "Error on packet!" & vbNewLine & "Header: " & Parse(0) & vbNewLine & "Packet: " & Data, vbOKOnly
    End If

    ' Parse's Without Being Online

    If Not IsPlaying(index) Then

        Select Case Parse(0)
         Case PacketID.GatGlasses
            Call SendNewCharClasses(index)
            Exit Sub

         Case PacketID.NewFAccountIED

            If Not IsLoggedIn(index) Then
                Name = Parse(1)
                Password = Parse(2)
                Imail = Parse(3)

                If Imail = "" Then
                    Call PlainMsg(index, "You need to fill in your e-mail", 1)
                End If

                For i = 1 To Len(Name)
                    n = Asc(Mid(Name, i, 1))

                    If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                     Else
                        Call PlainMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.", 1)
                        Exit Sub
                    End If

                Next i

                If Not AccountExist(Name) Then
                    Call AddAccount(index, Name, Password, Imail)
                    Call TextAdd(frmServer.txtText(0), "Account " & Name & " has been created.", True)
                    Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                    Call PlainMsg(index, "Your account has been created!", 1)
                 Else
                    Call PlainMsg(index, "Sorry, that account name is already taken!", 1)
                End If

            End If
            Exit Sub

         Case PacketID.DelimAccounted

            If Not IsLoggedIn(index) Then
                Name = Parse(1)
                Password = Parse(2)

                If Not AccountExist(Name) Then
                    Call PlainMsg(index, "That account name does not exist.", 2)
                    Exit Sub
                End If

                If Not PasswordOK(Name, Password) Then
                    Call PlainMsg(index, "Incorrect password.", 2)
                    Exit Sub
                End If

                Call LoadPlayer(index, Name)

                For i = 1 To MAX_CHARS

                    If Trim(Player(index).Char(i).Name) <> "" Then
                        Call DeleteName(Player(index).Char(i).Name)
                    End If

                Next i
                Call ClearPlayer(index)

                Kill App.Path & "\accounts\" & Name & "_info.ini"

                For i = 1 To MAX_CHARS
                    Kill App.Path & "\accounts\" & Name & "\" & "char" & i & ".dat"
                Next i

                RmDir App.Path & "\accounts\" & Name & "\"

                Call AddLog("Account " & Trim(Name) & " has been deleted.", PLAYER_LOG)
                Call PlainMsg(index, "Your account has been deleted.", 2)
            End If

            Exit Sub

         Case PacketID.Logination

            If Not IsLoggedIn(index) Then
                Name = Parse(1)
                Password = Parse(2)

                If ReadINI("CONFIG", "verified", App.Path & "\Data.ini") = 1 Then
                    If Val(ReadINI("GENERAL", "verified", App.Path & "\accounts\" & Trim(Player(index).Login) & ".ini")) = 0 Then
                        Call MsgBox("Your account hasn't been verified yet!", vbCritical)
                        Exit Sub
                    End If

                End If

                For i = 1 To Len(Name)
                    n = Asc(Mid(Name, i, 1))

                    If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                     Else
                        Call PlainMsg(index, "Account duping is not allowed!", 3)
                        Exit Sub
                    End If

                Next i

                If Val(Parse(3)) < CLIENT_MAJOR Or Val(Parse(4)) < CLIENT_MINOR Or Val(Parse(5)) < CLIENT_REVISION Then
                    Call PlainMsg(index, "Version outdated, please visit " & Trim(GetVar(App.Path & "\Data.ini", "CONFIG", "WebSite")), 3)
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

                If Parse(6) = SEC_CODE1 And Parse(7) = SEC_CODE2 And Parse(8) = SEC_CODE3 And Parse(9) = SEC_CODE4 Then
                 Else
                    Call AlertMsg(index, "Script Kiddy Alert!")
                    Exit Sub
                End If

                'Dim Packs As String
                'Packs = "MAXINFO" & SEP_CHAR
                'Packs = Packs & GAME_NAME & SEP_CHAR
                'Packs = Packs & MAX_PLAYERS & SEP_CHAR
                'Packs = Packs & MAX_ITEMS & SEP_CHAR
                'Packs = Packs & MAX_NPCS & SEP_CHAR
                'Packs = Packs & MAX_SHOPS & SEP_CHAR
                'Packs = Packs & MAX_SPELLS & SEP_CHAR
                'Packs = Packs & MAX_MAPS & SEP_CHAR
                'Packs = Packs & MAX_MAP_ITEMS & SEP_CHAR
                'Packs = Packs & MAX_MAPX & SEP_CHAR
                'Packs = Packs & MAX_MAPY & SEP_CHAR
                'Packs = Packs & MAX_EMOTICONS & SEP_CHAR
                'Packs = Packs & MAX_ELEMENTS & SEP_CHAR
                'Packs = Packs & PAPERDOLL & SEP_CHAR
                'Packs = Packs & SPRITESIZE & SEP_CHAR
                'Packs = Packs & MAX_SCRIPTSPELLS & SEP_CHAR
                'Packs = Packs & ENCRYPT_PASS & SEP_CHAR
                'Packs = Packs & ENCRYPT_TYPE & SEP_CHAR
                'Packs = Packs & END_CHAR
                'Call SendDataTo(index, Packs)

                Call LoadPlayer(index, Name)
                Call SendChars(index)

                Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
                Call TextAdd(frmServer.txtText(0), GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", True)
            End If

            Exit Sub

         Case PacketID.GiveMeTheMax
            Dim packs As String
            packs = PacketID.MaxInfo & SEP_CHAR
            packs = packs & GAME_NAME & SEP_CHAR
            packs = packs & MAX_PLAYERS & SEP_CHAR
            packs = packs & MAX_ITEMS & SEP_CHAR
            packs = packs & MAX_NPCS & SEP_CHAR
            packs = packs & MAX_SHOPS & SEP_CHAR
            packs = packs & MAX_SPELLS & SEP_CHAR
            packs = packs & MAX_MAPS & SEP_CHAR
            packs = packs & MAX_MAP_ITEMS & SEP_CHAR
            packs = packs & MAX_MAPX & SEP_CHAR
            packs = packs & MAX_MAPY & SEP_CHAR
            packs = packs & MAX_EMOTICONS & SEP_CHAR
            packs = packs & MAX_ELEMENTS & SEP_CHAR
            packs = packs & Paperdoll & SEP_CHAR
            packs = packs & Spritesize & SEP_CHAR
            packs = packs & MAX_SCRIPTSPELLS & SEP_CHAR
            packs = packs & ENCRYPT_PASS & SEP_CHAR
            packs = packs & ENCRYPT_TYPE & SEP_CHAR
            packs = packs & MAX_SKILLS & SEP_CHAR
            packs = packs & MAX_QUESTS & SEP_CHAR
            packs = packs & CUSTOM_SPRITE & SEP_CHAR
            packs = packs & GetVar(App.Path & "\Data.ini", "CONFIG", "Level") & SEP_CHAR
            packs = packs & MAX_PARTY_MEMBERS & SEP_CHAR
            packs = packs & STAT1 & SEP_CHAR
            packs = packs & STAT2 & SEP_CHAR
            packs = packs & STAT3 & SEP_CHAR
            packs = packs & STAT4 & SEP_CHAR
            packs = packs & END_CHAR
            Call SendDataTo(index, packs)
            'Send them the news too
            Call SendNewsTo(index)
            Exit Sub

         Case PacketID.AddAChara
            Dim headc As Long
            Dim bodyc As Long
            Dim legc As Long

            Name = Parse(1)
            Sex = Val(Parse(2))
            Class = Val(Parse(3))
            CharNum = Val(Parse(4))
            headc = Val(Parse(5))
            bodyc = Val(Parse(6))
            legc = Val(Parse(7))

            For i = 1 To Len(Name)
                n = Asc(Mid(Name, i, 1))

                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                 Else
                    Call PlainMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.", 4)
                    Exit Sub
                End If

            Next i

            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(index, "Invalid CharNum")
                Exit Sub
            End If

            If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
                Call HackingAttempt(index, "Invalid Sex")
                Exit Sub
            End If

            If Class < 0 Or Class > MAX_CLASSES Then
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

            Call AddChar(index, Name, Sex, Class, CharNum, headc, bodyc, legc)
            Call SavePlayer(index)
            Call AddLog("Character " & Name & " added to " & GetPlayerLogin(index) & "'s account.", PLAYER_LOG)
            Call SendChars(index)
            Call PlainMsg(index, "Character has been created!", 5)

            ' Dunno how useful this would be, but it's there if a future dev wants to work with it. -Pickle

            If Scripting = 1 Then
                Call MyScript.ExecuteStatement("Scripts\Main.txt", "OnNewChar " & index & "," & CharNum)
            End If

            Exit Sub

         Case PacketID.DelimboCharu
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

         Case PacketID.Usagakarim
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

         Case PacketID.Mail
            Call SendDataTo(index, PacketID.Mail & SEP_CHAR & Val(GetVar(App.Path & "\Data.ini", "CONFIG", "Email")) & SEP_CHAR & END_CHAR)
            Exit Sub
        End Select

    End If

    ' Parse's With Being Online And Playing
    If IsPlaying(index) = False Then Exit Sub
    If IsConnected(index) = False Then Exit Sub

    Select Case Parse(0)

        ' :::::::::::::::::::
        ' :: Guilds Packet ::
        ' :::::::::::::::::::
        ' Access
     Case PacketID.GuildChangeAccess

        If Parse(2) Like "*[!0-9]*" Then Exit Sub
        If Parse(2) = "" Then Exit Sub
        If Parse(2) = " " Then Exit Sub

        If Parse(2) > 4 Then
            Call PlayerMsg(index, "Sorry Invalid Access level", Red)
            Exit Sub
        End If

        If Parse(1) = "" Then
            Call PlayerMsg(index, "You must enter a player Name To proceed.", White)
            Exit Sub
        End If

        ' Check the requirements.

        If Parse(1) = "" Then
            Call PlayerMsg(index, "You must enter a player Name To proceed.", White)
            Exit Sub
        End If

        If FindPlayer(Parse(1)) = 0 Then
            Call PlayerMsg(index, "Player Is offline", White)
            Exit Sub
        End If

        If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(index) Then
            Call PlayerMsg(index, "Player Is Not In your guild", Red)
            Exit Sub
        End If

        If GetPlayerGuildAccess(index) < 4 Then
            Call PlayerMsg(index, "You are not the owner of this guild!", Red)
            Exit Sub
        End If

        'Set the player's New access level
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), Parse(2))
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub

        ' Disown
     Case PacketID.GuildDisown
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
        Call setplayerguild(FindPlayer(Parse(1)), "")
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub

        ' Leave Guild
     Case PacketID.GuildLeave
        ' Check if they can leave

        If GetPlayerGuild(index) = "" Then
            Call PlayerMsg(index, "You are not in a guild.", Red)
            Exit Sub
        End If

        Call setplayerguild(index, "")
        Call SetPlayerGuildAccess(index, 0)
        Call SendPlayerData(index)
        Exit Sub

        ' Make A New Guild
     Case PacketID.MakeGuild
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
        Call setplayerguild(FindPlayer(Parse(1)), (Parse(2)))
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 4)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub

        ' Make A Member
     Case PacketID.GuildMember
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
        Call setplayerguild(FindPlayer(Parse(1)), GetPlayerGuild(index))
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 1)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub

        ' Make A Trainie
     Case PacketID.GuildTrainee
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
        Call setplayerguild(FindPlayer(Parse(1)), GetPlayerGuild(index))
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub

        ' ::::::::::::::::::::
        ' :: Social packets ::
        ' ::::::::::::::::::::
     Case PacketID.SayMsg
        Msg = Parse(1)

        ' Prevent hacking

        For i = 1 To Len(Msg)

            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 255 Then
                Call HackingAttempt(index, "Say Text Modification")
                Exit Sub
            End If

        Next i

        If frmServer.chkM.Value = Unchecked Then
            If GetPlayerAccess(index) <= 0 Then
                Call PlayerMsg(index, "Map messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If

        End If

        'ASGARD
        'Check for swearing
        Msg = SwearCheck(Msg)

        Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " : " & Msg & "", PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " : " & Msg & "", SayColor)
        Call MapMsg2(GetPlayerMap(index), Msg, index)
        TextAdd frmServer.txtText(3), GetPlayerName(index) & " On Map " & GetPlayerMap(index) & ": " & Msg, True
        Exit Sub

     Case PacketID.EmoteMsg
        Msg = Parse(1)

        ' Prevent hacking

        For i = 1 To Len(Msg)

            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 255 Then
                Call HackingAttempt(index, "Emote Text Modification")
                Exit Sub
            End If

        Next i

        If frmServer.chkE.Value = Unchecked Then
            If GetPlayerAccess(index) <= 0 Then
                Call PlayerMsg(index, "Emote messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If

        End If

        Msg = SwearCheck(Msg)

        Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Msg, PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & Msg, EmoteColor)
        TextAdd frmServer.txtText(6), GetPlayerName(index) & " " & Msg, True
        Exit Sub

     Case PacketID.BroadcastMsg
        Msg = Parse(1)

        ' Prevent hacking

        For i = 1 To Len(Msg)

            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 255 Then
                Call HackingAttempt(index, "Broadcast Text Modification")
                Exit Sub
            End If

        Next i

        If frmServer.chkBC.Value = Unchecked Then
            If GetPlayerAccess(index) <= 0 Then
                Call PlayerMsg(index, "Broadcast messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If

        End If

        If Player(index).Mute = True Then Exit Sub

        Msg = SwearCheck(Msg)

        s = GetPlayerName(index) & ": " & Msg
        Call AddLog(s, PLAYER_LOG)
        Call GlobalMsg(s, BroadcastColor)
        Call TextAdd(frmServer.txtText(0), s, True)
        TextAdd frmServer.txtText(1), GetPlayerName(index) & ": " & Msg, True
        Exit Sub

     Case PacketID.GlobalMsg
        Msg = Parse(1)

        ' Prevent hacking

        For i = 1 To Len(Msg)

            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 255 Then
                Call HackingAttempt(index, "Global Text Modification")
                Exit Sub
            End If

        Next i

        If frmServer.chkG.Value = Unchecked Then
            If GetPlayerAccess(index) <= 0 Then
                Call PlayerMsg(index, "Global messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If

        End If

        If Player(index).Mute = True Then Exit Sub

        Msg = SwearCheck(Msg)

        If GetPlayerAccess(index) > 0 Then
            s = "(global) " & GetPlayerName(index) & ": " & Msg
            Call AddLog(s, ADMIN_LOG)
            Call GlobalMsg(s, GlobalColor)
            Call TextAdd(frmServer.txtText(0), s, True)
        End If

        TextAdd frmServer.txtText(2), GetPlayerName(index) & ": " & Msg, True
        Exit Sub

     Case PacketID.AdminMsg
        Msg = Parse(1)

        ' Prevent hacking

        For i = 1 To Len(Msg)

            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 255 Then
                Call HackingAttempt(index, "Admin Text Modification")
                Exit Sub
            End If

        Next i

        If frmServer.chkA.Value = Unchecked Then
            Call PlayerMsg(index, "Admin messages have been disabled by the server!", BrightRed)
            Exit Sub
        End If

        Msg = SwearCheck(Msg)

        If GetPlayerAccess(index) > 0 Then
            Call AddLog("(admin " & GetPlayerName(index) & ") " & Msg, ADMIN_LOG)
            Call AdminMsg("(admin " & GetPlayerName(index) & ") " & Msg, AdminColor)
        End If

        TextAdd frmServer.txtText(5), GetPlayerName(index) & ": " & Msg, True
        Exit Sub

     Case PacketID.PlayerMsg
        MsgTo = FindPlayer(Parse(1))
        Msg = Parse(2)

        If Msg = "" Then
            Call PlayerMsg(index, "You must send a message to private message another player.", BrightRed)
            Exit Sub
        End If

        If MsgTo = 0 Then
            Call PlayerMsg(index, "That player is not online!", BrightRed)
            Exit Sub
        End If

        ' Prevent hacking

        For i = 1 To Len(Msg)

            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 255 Then
                Call HackingAttempt(index, "Player Msg Text Modification")
                Exit Sub
            End If

        Next i

        If frmServer.chkP.Value = Unchecked Then
            If GetPlayerAccess(index) <= 0 Then
                Call PlayerMsg(index, "PM messages have been disabled by the server!", BrightRed)
                Exit Sub
            End If

        End If

        If FindPlayer(Parse(1)) = 0 Then
            Call PlayerMsg(index, "Player is offline", White)
            Exit Sub
        End If

        Msg = SwearCheck(Msg)

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

        ' :::::::::::::::::::::::
        ' :: edit main  packet ::
        ' :::::::::::::::::::::::

     Case PacketID.EditMain
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_MAPPER And 0 + Val(ReadINI(GetPlayerName(index), "editmain", App.Path & "\Acces.ini")) <> 1 Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        hFile = FreeFile
        Open App.Path & "\Scripts\Main.txt" For Input As #hFile
        frmEditor.RT.text = Input$(LOF(hFile), hFile)
        Close #hFile
        Call SendDataTo(index, PacketID.Main & SEP_CHAR & frmEditor.RT.text & SEP_CHAR & END_CHAR)
        Exit Sub

        ' :::::::::::::::::::::::
        ' :: save main  packet ::
        ' :::::::::::::::::::::::

     Case PacketID.SaveMain
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_MAPPER And 0 + Val(ReadINI(GetPlayerName(index), "editmain", App.Path & "\Acces.ini")) <> 1 Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        AFileName = "Scripts\Main.txt"
        Open App.Path & "\" & AFileName For Output As #1
        Print #1, Trim$(Parse$(1))
        Close #1
        Exit Sub

        ' :::::::::::::::::::::::::::::
        ' :: Moving character packet ::
        ' :::::::::::::::::::::::::::::
     Case PacketID.PlayerMove

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

        ' Prevent player from moving if they have been script locked

        If Player(index).locked = True Then
            Call SendPlayerXY(index)
            Exit Sub
        End If

        Call PlayerMove(index, Dir, Movement)
        Exit Sub

        ' :::::::::::::::::::::::::::::
        ' :: Moving character packet ::
        ' :::::::::::::::::::::::::::::
     Case PacketID.PlayerDir

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
        Call SendDataToMapBut(index, GetPlayerMap(index), PacketID.PlayerDir & SEP_CHAR & index & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & END_CHAR)
        Exit Sub

        ' :::::::::::::::::::::
        ' :: Use item packet ::
        ' :::::::::::::::::::::
     Case PacketID.UseItem
        InvNum = Val(Parse(1))
        CharNum = Player(index).CharNum
        ' Prevent player from using an item when they are locked

        If Player(index).lockeditems = True Then
            Exit Sub
        End If

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

            Dim n1 As Long
              Dim n2 As Long
              Dim n3 As Long
              Dim n4 As Long
              Dim n5 As Long
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

             Case ITEM_TYPE_LEGS

                If InvNum <> GetPlayerLegsSlot(index) Then
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

                    Call SetPlayerLegsSlot(index, InvNum)
                 Else
                    Call SetPlayerLegsSlot(index, 0)
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

             Case ITEM_TYPE_NECKLACE

                If InvNum <> GetPlayerNecklaceSlot(index) Then
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

                    Call SetPlayerNecklaceSlot(index, InvNum)
                 Else
                    Call SetPlayerNecklaceSlot(index, 0)
                End If

                Call SendWornEquipment(index)

             Case ITEM_TYPE_POTIONADDHP
                Call SetPlayerHP(index, GetPlayerHP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)

                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
                 Else
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                End If

                Call SendHP(index)

             Case ITEM_TYPE_POTIONADDMP
                Call SetPlayerMP(index, GetPlayerMP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)

                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
                 Else
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                End If

                Call SendMP(index)

             Case ITEM_TYPE_POTIONADDSP
                Call SetPlayerSP(index, GetPlayerSP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)

                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
                 Else
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                End If

                Call SendSP(index)

             Case ITEM_TYPE_POTIONSUBHP
                Call SetPlayerHP(index, GetPlayerHP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)

                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
                 Else
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                End If

                Call SendHP(index)

             Case ITEM_TYPE_POTIONSUBMP
                Call SetPlayerMP(index, GetPlayerMP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)

                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
                 Else
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                End If

                Call SendMP(index)

             Case ITEM_TYPE_POTIONSUBSP
                Call SetPlayerSP(index, GetPlayerSP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)

                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
                 Else
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                End If

                Call SendSP(index)

             Case ITEM_TYPE_KEY

                Select Case GetPlayerDir(index)
                 Case DIR_UP

                    If GetPlayerY(index) > 0 Then
                        X = GetPlayerX(index)
                        Y = GetPlayerY(index) - 1
                     Else
                        Exit Sub
                    End If

                 Case DIR_DOWN

                    If GetPlayerY(index) < MAX_MAPY Then
                        X = GetPlayerX(index)
                        Y = GetPlayerY(index) + 1
                     Else
                        Exit Sub
                    End If

                 Case DIR_LEFT

                    If GetPlayerX(index) > 0 Then
                        X = GetPlayerX(index) - 1
                        Y = GetPlayerY(index)
                     Else
                        Exit Sub
                    End If

                 Case DIR_RIGHT

                    If GetPlayerX(index) < MAX_MAPY Then
                        X = GetPlayerX(index) + 1
                        Y = GetPlayerY(index)
                     Else
                        Exit Sub
                    End If

                End Select

                ' Check if a key exists

                If Map(GetPlayerMap(index)).tile(X, Y).Type = TILE_TYPE_KEY Then
                    ' Check if the key they are using matches the map key

                    If GetPlayerInvItemNum(index, InvNum) = Map(GetPlayerMap(index)).tile(X, Y).Data1 Then
                        TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                        TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

                        Call SendDataToMap(GetPlayerMap(index), PacketID.MapKey & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)

                        If Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String1) = "" Then
                            Call MapMsg(GetPlayerMap(index), "A door has been unlocked!", White)
                         Else
                            Call MapMsg(GetPlayerMap(index), Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String1), White)
                        End If

                        Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)

                        ' Check if we are supposed to take away the item

                        If Map(GetPlayerMap(index)).tile(X, Y).Data2 = 1 Then
                            Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                            Call PlayerMsg(index, "The key disolves.", Yellow)
                        End If

                    End If
                End If

                If Map(GetPlayerMap(index)).tile(X, Y).Type = TILE_TYPE_DOOR Then
                    TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                    TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

                    Call SendDataToMap(GetPlayerMap(index), PacketID.MapKey & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
                End If

             Case ITEM_TYPE_SPELL
                ' Get the spell num
                n = Item(GetPlayerInvItemNum(index, InvNum)).Data1

                If n > 0 Then
                    ' Make sure they are the right class

                    If Spell(n).ClassReq - 1 = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                        If Spell(n).LevelReq = 0 And Player(index).Char(Player(index).CharNum).access < 1 Then
                            Call PlayerMsg(index, "This spell can only be used by admins!", BrightRed)
                            Exit Sub
                        End If

                        ' Make sure they are the right level
                        i = GetSpellReqLevel(n)

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

             Case ITEM_TYPE_SCRIPTED
                MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedItem " & index & "," & Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1

            End Select

            Call SendStats(index)
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)

            ' Send everyone player's equipment
            Call SendIndexWornEquipment(index)
        End If

        '  End If
        Exit Sub

     Case PacketID.PlayerMoveMouse

        If 0 + Val(ReadINI("CONFIG", PacketID.Mouse, App.Path & "\Data.ini")) = 1 Then
            Call SendDataTo(index, PacketID.Mouse & SEP_CHAR & END_CHAR)
        End If

        If Player(index).GettingMap = YES Then
            Exit Sub
        End If

        Dir = Val(Parse(1))
        Movement = 1
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

        ' Prevent player from moving if they have been script locked

        If Player(index).locked = True Then
            Call SendPlayerXY(index)
            Exit Sub
        End If

        Exit Sub

     Case PacketID.Warp
        Dim direction As Long

        direction = Val(Parse(1))

        If direction = 0 Then

            If Map(GetPlayerMap(index)).Up > 0 Then
                Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), MAX_MAPY)
                Exit Sub
            End If

        End If

        If direction = 1 Then

            If Map(GetPlayerMap(index)).Down > 0 Then
                Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                Exit Sub
            End If

        End If

        If direction = 2 Then

            If Map(GetPlayerMap(index)).left > 0 Then
                Call PlayerWarp(index, Map(GetPlayerMap(index)).left, MAX_MAPX, GetPlayerY(index))
                Exit Sub
            End If

        End If

        If direction = 3 Then

            If Map(GetPlayerMap(index)).Right > 0 Then
                Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                Exit Sub
            End If

        End If

        ' ::::::::::::::::::
        ' :: graplepacket ::
        ' ::::::::::::::::::
     Case PacketID.EndShot

        If Val(Parse(1)) = 0 Then
            Player(index).locked = False
            Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
            Player(index).HookShotX = 0
            Player(index).HookShotY = 0
            Exit Sub
        End If

        If Player(index).HookShotX = 0 Or Player(index).HookShotY = 0 Then
            Call HackingAttempt(index, "")
        End If

        Call PlayerMsg(index, "You carefully cross the wire.", 1)
        Player(index).locked = False

        'This doesn't work... you need to make changes to the game loop
        ' to get an animation effect. This just hangs the server. -Pickle
        'If GetPlayerX(index) < Player(index).HookShotX Then
        'Do While GetPlayerX(index) < Player(index).HookShotX
        '    Call SetPlayerX(index, GetPlayerX(index) + 1)
        '    Call SetPlayerY(index, Player(index).HookShotY)
        'Loop
        'End If

        'If GetPlayerX(index) > Player(index).HookShotX Then
        'Do While GetPlayerX(index) > Player(index).HookShotX
        '    Call SetPlayerX(index, GetPlayerX(index) - 1)
        '    Call SetPlayerY(index, Player(index).HookShotY)
        'Loop
        'End If

        'If GetPlayerY(index) < Player(index).HookShotY Then
        'Do While GetPlayerX(index) < Player(index).HookShotY
        '    Call SetPlayerY(index, GetPlayerY(index) + 1)
        '    Call SetPlayerX(index, Player(index).HookShotX)
        'Loop
        'End If

        'If GetPlayerY(index) > Player(index).HookShotY Then
        'Do While GetPlayerX(index) > Player(index).HookShotY
        '    Call SetPlayerY(index, GetPlayerY(index) - 1)
        '    Call SetPlayerX(index, Player(index).HookShotX)
        'Loop
        'End If

        ' Temporary fix. We'll see if we can do animation later... :) -Pickle
        Call SetPlayerX(index, Player(index).HookShotX)
        Call SetPlayerY(index, Player(index).HookShotY)
        Player(index).HookShotX = 0
        Player(index).HookShotY = 0
        Call SendPlayerData(index)
        Exit Sub

        ' ::::::::::::::::::::::::::
        ' :: Player attack packet ::
        ' ::::::::::::::::::::::::::
     Case PacketID.Attack
        ' Prevent player from moving if they have been script locked

        If Player(index).lockedattack = True Then
            Exit Sub
        End If

        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & index
        End If

        If GetPlayerWeaponSlot(index) > 0 Then
            If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 > 0 Then
                If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Stackable = 0 Then
                    Call SendDataToMap(GetPlayerMap(index), PacketID.CheckArrows & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & END_CHAR)
                 Else
                    Call GrapleHook(index)
                End If

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
                            Damage = GetPlayerDamage(index) - GetPlayerProtection(i)
                            Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                         Else
                            n = GetPlayerDamage(index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                            surge = GetVar("Lang.ini", "Lang", "Surge")

                            If surge <> "" Then
                                Call BattleMsg(index, "You feel a surge of energy upon swinging!", BrightCyan, 0)
                                Call BattleMsg(i, GetPlayerName(index) & " swings with enormous might!", BrightCyan, 1)
                            End If

                            'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                            'Call PlayerMsg(I, GetPlayerName(index) & " swings with enormous might!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                        End If

                        If Damage > 0 Then
                            Call AttackPlayer(index, i, Damage)
                         Else
                            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                            Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If

                     Else
                        Call BattleMsg(index, GetPlayerName(i) & " blocked your hit!", BrightCyan, 0)
                        Call BattleMsg(i, "You blocked " & GetPlayerName(index) & "'s hit!", BrightCyan, 1)

                        'Call PlayerMsg(index, GetPlayerName(I) & "'s " & Trim(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked your hit!", BrightCyan)
                        'Call PlayerMsg(I, "Your " & Trim(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
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
                Player(index).targetnpc = i

                If Not CanPlayerCriticalHit(index) Then
                    Damage = GetPlayerDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).DEF / 2)
                    Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                 Else
                    n = GetPlayerDamage(index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).DEF / 2)
                    Call BattleMsg(index, "You feel a surge of energy upon swinging!", BrightCyan, 0)

                    'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                    Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                End If

                If Damage > 0 Then
                    '//!! Display damage before applying
                    Call SendDataTo(index, PacketID.BlitPlayerDmg & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    Call AttackNpc(index, i, Damage)
                 Else
                    Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)

                    'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                    Call SendDataTo(index, PacketID.BlitPlayerDmg & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                End If

                Exit Sub
            End If

        Next i

        ' Check for skill

        Select Case GetPlayerDir(index)
         Case DIR_UP

            If GetPlayerY(index) > 0 Then

                With Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) - 1)

                    If .Type = TILE_TYPE_SKILL Then
                        Call DoSkill(index, .Data1, .Data2)
                    End If

                End With
            End If

         Case DIR_DOWN

            If GetPlayerY(index) < MAX_MAPY Then

                With Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index) + 1)

                    If .Type = TILE_TYPE_SKILL Then
                        Call DoSkill(index, .Data1, .Data2)
                    End If

                End With
            End If

         Case DIR_LEFT

            If GetPlayerX(index) > 0 Then

                With Map(GetPlayerMap(index)).tile(GetPlayerX(index) - 1, GetPlayerY(index))

                    If .Type = TILE_TYPE_SKILL Then
                        Call DoSkill(index, .Data1, .Data2)
                    End If

                End With
            End If

         Case DIR_RIGHT

            If GetPlayerX(index) < MAX_MAPX Then

                With Map(GetPlayerMap(index)).tile(GetPlayerX(index) + 1, GetPlayerY(index))

                    If .Type = TILE_TYPE_SKILL Then
                        Call DoSkill(index, .Data1, .Data2)
                    End If

                End With
            End If

        End Select

        Call AttackAttributeNpcs(index)
        Exit Sub

        ' ::::::::::::::::::::::
        ' :: Use stats packet ::
        ' ::::::::::::::::::::::
     Case PacketID.UseStatPoint
        PointType = Val(Parse(1))

        ' Prevent hacking

        If (PointType < 0) Or (PointType > 3) Then
            Call HackingAttempt(index, "Invalid Point Type")
            Exit Sub
        End If

        ' Make sure they have points

        If GetPlayerPOINTS(index) > 0 Then
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "UsingStatPoints " & index & "," & PointType
             Else

                Select Case PointType
                 Case 0
                    Call SetPlayerSTR(index, GetPlayerSTR(index) + 1)
                    Call BattleMsg(index, "You have gained more strength!", 15, 0)

                 Case 1
                    Call SetPlayerDEF(index, GetPlayerDEF(index) + 1)
                    Call BattleMsg(index, "You have gained more defense!", 15, 0)

                 Case 2
                    Call SetPlayerMAGI(index, GetPlayerMAGI(index) + 1)
                    Call BattleMsg(index, "You have gained more magic abilities!", 15, 0)

                 Case 3
                    Call SetPlayerSPEED(index, GetPlayerSPEED(index) + 1)
                    Call BattleMsg(index, "You have gained more speed!", 15, 0)
                End Select

                Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)
            End If

         Else
            Call BattleMsg(index, "You have no skill points to train with!", BrightRed, 0)
        End If

        Call SendHP(index)
        Call SendMP(index)
        Call SendSP(index)
        Call SendStats(index)

        Call SendDataTo(index, PacketID.PlayerPoints & SEP_CHAR & GetPlayerPOINTS(index) & SEP_CHAR & END_CHAR)
        Exit Sub

        ' ::::::::::::::::::::::::::::::::
        ' :: Player info request packet ::
        ' ::::::::::::::::::::::::::::::::
     Case PacketID.PlayerInfoRequest
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

        ' :::::::::::::::::::::::
        ' :: Set sprite packet ::
        ' :::::::::::::::::::::::
     Case PacketID.SetSprite
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

        ' ::::::::::::::::::::::::::::::
        ' :: Set player sprite packet ::
        ' ::::::::::::::::::::::::::::::
     Case PacketID.SetPlayerSprite
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

        ' ::::::::::::::::::::::::::
        ' :: Stats request packet ::
        ' ::::::::::::::::::::::::::
     Case PacketID.GetStats
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

     Case PacketID.CanonShoot

        If HasItem(index, Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1) > 0 Then
            Call TakeItem(index, Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1, 1)
            Call SendDataToMap(GetPlayerMap(index), PacketID.ScriptSpellAnim & SEP_CHAR & Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3 & SEP_CHAR & Spell(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3).SpellAnim & SEP_CHAR & Spell(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3).SpellTime & SEP_CHAR & Spell(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3).SpellDone & SEP_CHAR & Val(Parse(1)) & SEP_CHAR & Val(Parse(2)) & SEP_CHAR & Spell(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3).Big & SEP_CHAR & END_CHAR)

            For i = 1 To MAX_PLAYERS

                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
                    If GetPlayerX(i) = Val(Parse(1)) And GetPlayerY(i) = Val(Parse(2)) Then

                        Call SetPlayerHP(index, GetPlayerHP(index) - Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2)
                        Call SendHP(index)

                        If 0 + GetPlayerHP(index) <= 0 Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & index
                            Call SetPlayerHP(index, GetPlayerMaxHP(index))
                            Call SendHP(index)
                        End If

                        Exit Sub
                    End If

                End If
            Next i

            Call GlobalMsg("done", 1)
        End If

        Exit Sub

        ' ::::::::::::::::::::::::::::::::::
        ' :: Player request for a new map ::
        ' ::::::::::::::::::::::::::::::::::
     Case PacketID.RequestNewMap
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
     Case PacketID.MapData
        ' Error Handling
        Err.Clear
        On Error Resume Next
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        n = 1

        MapNum = Val(Parse(n))
        Map(MapNum).Name = Parse(n + 1)
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        Map(MapNum).Moral = Val(Parse(n + 3))
        Map(MapNum).Up = Val(Parse(n + 4))
        Map(MapNum).Down = Val(Parse(n + 5))
        Map(MapNum).left = Val(Parse(n + 6))
        Map(MapNum).Right = Val(Parse(n + 7))
        Map(MapNum).music = Parse(n + 8)
        Map(MapNum).BootMap = Val(Parse(n + 9))
        Map(MapNum).BootX = Val(Parse(n + 10))
        Map(MapNum).BootY = Val(Parse(n + 11))
        Map(MapNum).Indoors = Val(Parse(n + 12))
        Map(MapNum).Weather = Val(Parse(n + 13))

        n = n + 14

        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(MapNum).tile(X, Y).Ground = Val(Parse(n))
                Map(MapNum).tile(X, Y).Mask = Val(Parse(n + 1))
                Map(MapNum).tile(X, Y).Anim = Val(Parse(n + 2))
                Map(MapNum).tile(X, Y).Mask2 = Val(Parse(n + 3))
                Map(MapNum).tile(X, Y).M2Anim = Val(Parse(n + 4))
                Map(MapNum).tile(X, Y).Fringe = Val(Parse(n + 5))
                Map(MapNum).tile(X, Y).FAnim = Val(Parse(n + 6))
                Map(MapNum).tile(X, Y).Fringe2 = Val(Parse(n + 7))
                Map(MapNum).tile(X, Y).F2Anim = Val(Parse(n + 8))
                Map(MapNum).tile(X, Y).Type = Val(Parse(n + 9))
                Map(MapNum).tile(X, Y).Data1 = Val(Parse(n + 10))
                Map(MapNum).tile(X, Y).Data2 = Val(Parse(n + 11))
                Map(MapNum).tile(X, Y).Data3 = Val(Parse(n + 12))
                Map(MapNum).tile(X, Y).String1 = Parse(n + 13)
                Map(MapNum).tile(X, Y).String2 = Parse(n + 14)
                Map(MapNum).tile(X, Y).String3 = Parse(n + 15)
                Map(MapNum).tile(X, Y).light = Val(Parse(n + 16))
                Map(MapNum).tile(X, Y).GroundSet = Val(Parse(n + 17))
                Map(MapNum).tile(X, Y).MaskSet = Val(Parse(n + 18))
                Map(MapNum).tile(X, Y).AnimSet = Val(Parse(n + 19))
                Map(MapNum).tile(X, Y).Mask2Set = Val(Parse(n + 20))
                Map(MapNum).tile(X, Y).M2AnimSet = Val(Parse(n + 21))
                Map(MapNum).tile(X, Y).FringeSet = Val(Parse(n + 22))
                Map(MapNum).tile(X, Y).FAnimSet = Val(Parse(n + 23))
                Map(MapNum).tile(X, Y).Fringe2Set = Val(Parse(n + 24))
                Map(MapNum).tile(X, Y).F2AnimSet = Val(Parse(n + 25))

                n = n + 26
            Next X

        Next Y

        For X = 1 To MAX_MAP_NPCS
            Map(MapNum).Npc(X) = Val(Parse(n))
            n = n + 1
            Call ClearMapNpc(X, MapNum)
        Next X

        ' Clear out it all

        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).X, MapItem(GetPlayerMap(index), i).Y)
            Call ClearMapItem(i, GetPlayerMap(index))
        Next i

        ' Save the map
        Call SaveMap(MapNum)

        ' Respawn
        Call SpawnMapItems(GetPlayerMap(index))

        ' Respawn NPCS

        For i = 1 To MAX_MAP_NPCS
            Call SpawnNPC(i, GetPlayerMap(index))
        Next i

        ' Refresh map for everyone online

        For i = 1 To MAX_PLAYERS

            If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & SEP_CHAR & END_CHAR)
                'Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
            End If

        Next i

        If Err.number = 6 Then
            Call BattleMsg(index, "Error saving map, some parts may not have changed.", BrightRed, 0)
            Err.Clear
         ElseIf Err.number > 0 Then
            Call BattleMsg(index, "Unexpected error occured - map may not have saved.", BrightRed, 0)
            Err.Clear
        End If

        Exit Sub

        ' Reset error handling
        On Error GoTo 0

        ' ::::::::::::::::::::::::::::
        ' :: Need map yes/no packet ::
        ' ::::::::::::::::::::::::::::
     Case PacketID.NeedMap
        ' Get yes/no value
        s = LCase(Parse(1))

        If s = "yes" Then
            Call SendMap(index, GetPlayerMap(index))
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            Player(index).GettingMap = NO
            Call SendDataTo(index, PacketID.MapDone & SEP_CHAR & END_CHAR)
         Else
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            Player(index).GettingMap = NO
            Call SendDataTo(index, PacketID.MapDone & SEP_CHAR & END_CHAR)
        End If

        ' Send everyone player's equipment
        Call SendIndexWornEquipment(index)
        Call SendIndexWornEquipmentFromMap(index)
        Exit Sub

     '//!! Unused packet
     'Case PacketID.NeedMapNum2
     '   Call SendMap(index, GetPlayerMap(index))
     '   Exit Sub

        ' :::::::::::::::::::::::::::::::::::::::::::::::
        ' :: Player trying to pick up something packet ::
        ' :::::::::::::::::::::::::::::::::::::::::::::::
     Case PacketID.MapGetItem
        Call PlayerMapGetItem(index)
        Exit Sub

        ' ::::::::::::::::::::::::::::::::::::::::::::
        ' :: Player trying to drop something packet ::
        ' ::::::::::::::::::::::::::::::::::::::::::::
     Case PacketID.MapDropItem
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

        If Item(GetPlayerInvItemNum(index, InvNum)).Type <> ITEM_TYPE_CURRENCY And Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
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
        Exit Sub

        ' ::::::::::::::::::::::::
        ' :: Respawn map packet ::
        ' ::::::::::::::::::::::::
     Case PacketID.MapRespawn
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        ' Clear out it all

        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).X, MapItem(GetPlayerMap(index), i).Y)
            Call ClearMapItem(i, GetPlayerMap(index))
        Next i

        ' Respawn
        Call SpawnMapItems(GetPlayerMap(index))

        ' Respawn NPCS

        For i = 1 To MAX_MAP_NPCS
            Call SpawnNPC(i, GetPlayerMap(index))
        Next i

        Call PlayerMsg(index, "Map respawned.", Blue)
        Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)
        Exit Sub

        ' ::::::::::::::::::::::::
        ' :: Kick player packet ::
        ' ::::::::::::::::::::::::
     Case PacketID.KickPlayer
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

        ' :::::::::::::::::::::
        ' :: Ban list packet ::
        ' :::::::::::::::::::::
     Case PacketID.Banlist
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

        ' ::::::::::::::::::::::::
        ' :: Ban destroy packet ::
        ' ::::::::::::::::::::::::
     Case PacketID.BanDestroy
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
     Case PacketID.BanPlayer
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

        ' :::::::::::::::::::::::::::::
        ' :: Request edit map packet ::
        ' :::::::::::::::::::::::::::::
     Case PacketID.RequestEditMap
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.EditMap & SEP_CHAR & END_CHAR)
        Exit Sub
        ' :::::::::::::::::::::::::::::
        ' :: Request edit house packet ::
        ' :::::::::::::::::::::::::::::
     Case PacketID.RequestEditHouse
        ' Prevent hacking

        If Map(GetPlayerMap(index)).Moral <> MAP_MORAL_HOUSE Then
            Call PlayerMsg(index, "This is not a house!", BrightRed)
            Exit Sub
        End If

        If Map(GetPlayerMap(index)).owner <> GetPlayerName(index) Then
            Call PlayerMsg(index, "This is not your house!", BrightRed)
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.EditHouse & SEP_CHAR & END_CHAR)
        Exit Sub
        ' ::::::::::::::::::::::::::::::
        ' :: Request edit item packet ::
        ' ::::::::::::::::::::::::::::::
     Case PacketID.RequestEditItem
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.ItemEditor & SEP_CHAR & END_CHAR)
        Exit Sub

        ' ::::::::::::::::::::::
        ' :: Edit item packet ::
        ' ::::::::::::::::::::::
     Case PacketID.EditItem
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

        ' ::::::::::::::::::::::
        ' :: Save item packet ::
        ' ::::::::::::::::::::::
     Case PacketID.SaveItem
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
        Item(n).AttackSpeed = Val(Parse(22))
        Item(n).Price = Val(Parse(23))
        Item(n).Stackable = Val(Parse(24))
        Item(n).Bound = Val(Parse(25))

        ' Save it
        Call SendUpdateItemToAll(n)
        Call SaveItem(n)
        Call AddLog(GetPlayerName(index) & " saved item #" & n & ".", ADMIN_LOG)
        Exit Sub

        ' :::::::::::::::::::::
        ' :: Day/Night Stuff ::
        ' :::::::::::::::::::::
    
     '//!! Unused packet
     'Case PacketID.EnableDayNight
     '   ' Prevent hacking
     '
     '   If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
     '       Call HackingAttempt(index, "Admin Cloning")
     '       Exit Sub
     '   End If
     '
     '   If TimeDisable = False Then
     '       Gamespeed = 0
     '       frmServer.GameTimeSpeed.text = 0
     '       TimeDisable = True
     '       frmServer.Timer1.Enabled = False
     '       frmServer.Command69.caption = "Enable Time"
     '    Else
     '       Gamespeed = 1
     '       frmServer.GameTimeSpeed.text = 1
     '       TimeDisable = False
     '       frmServer.Timer1.Enabled = True
     '       frmServer.Command69.caption = "Disable Time"
     '   End If
     '
     '   Exit Sub

     Case PacketID.DayNight
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
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
     Case PacketID.RequestEditNPC
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.NPCEditor & SEP_CHAR & END_CHAR)
        Exit Sub

        ' :::::::::::::::::::::
        ' :: Edit npc packet ::
        ' :::::::::::::::::::::
     Case PacketID.EditNPC
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

        ' :::::::::::::::::::::
        ' :: Save npc packet ::
        ' :::::::::::::::::::::
     Case PacketID.SaveNPC
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
        Npc(n).Speed = Val(Parse(10))
        Npc(n).Magi = Val(Parse(11))
        Npc(n).Big = Val(Parse(12))
        Npc(n).MaxHp = Val(Parse(13))
        Npc(n).Exp = Val(Parse(14))
        Npc(n).SpawnTime = Val(Parse(15))
        Npc(n).Element = Val(Parse(16))
        Npc(n).Spritesize = Val(Parse(17))

        z = 18

        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = Val(Parse(z))
            Npc(n).ItemNPC(i).ItemNum = Val(Parse(z + 1))
            Npc(n).ItemNPC(i).ItemValue = Val(Parse(z + 2))
            z = z + 3
        Next i

        ' Save it
        Call SendUpdateNpcToAll(n)
        Call SaveNPC(n)
        Call AddLog(GetPlayerName(index) & " saved npc #" & n & ".", ADMIN_LOG)
        Exit Sub

        ' ::::::::::::::::::::::::::::::
        ' :: Request edit shop packet ::
        ' ::::::::::::::::::::::::::::::
     Case PacketID.RequestEditShop
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.ShopEditor & SEP_CHAR & END_CHAR)
        Exit Sub

        ' ::::::::::::::::::::::
        ' :: Edit shop packet ::
        ' ::::::::::::::::::::::
     Case PacketID.EditShop
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

        ' ::::::::::::::::::::::
        ' :: Save shop packet ::
        ' ::::::::::::::::::::::
     Case PacketID.SaveShop
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
        Shop(ShopNum).FixesItems = Val(Parse(3))
        Shop(ShopNum).BuysItems = Val(Parse(4))
        Shop(ShopNum).ShowInfo = Val(Parse(5))
        Shop(ShopNum).currencyItem = Val(Parse(6))

        n = 7

        For z = 1 To MAX_SHOP_ITEMS
            Shop(ShopNum).ShopItem(z).ItemNum = Val(Parse(n))
            Shop(ShopNum).ShopItem(z).Amount = Val(Parse(n + 1))
            Shop(ShopNum).ShopItem(z).Price = Val(Parse(n + 2))
            n = n + 3
        Next z

        ' Save it
        Call SendUpdateShopToAll(ShopNum)
        Call SaveShop(ShopNum)
        Call AddLog(GetPlayerName(index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
        Exit Sub

        ' :::::::::::::::::::::::::::::::
        ' :: Request edit spell packet ::
        ' :::::::::::::::::::::::::::::::
     Case PacketID.RequestEditSpell
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.SpellEditor & SEP_CHAR & END_CHAR)
        Exit Sub

        ' :::::::::::::::::::::::
        ' :: Edit spell packet ::
        ' :::::::::::::::::::::::
     Case PacketID.EditSpell
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

        ' :::::::::::::::::::::::
        ' :: Save spell packet ::
        ' :::::::::::::::::::::::
     Case PacketID.SaveSpell
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
        Spell(n).Range = Val(Parse(11))
        Spell(n).SpellAnim = Val(Parse(12))
        Spell(n).SpellTime = Val(Parse(13))
        Spell(n).SpellDone = Val(Parse(14))
        Spell(n).AE = Val(Parse(15))
        Spell(n).Big = Val(Parse(16))
        Spell(n).Element = Val(Parse(17))

        ' Save it
        Call SendUpdateSpellToAll(n)
        Call SaveSpell(n)
        Call AddLog(GetPlayerName(index) & " saving spell #" & n & ".", ADMIN_LOG)
        Exit Sub

     Case PacketID.ForgetSpell
        ' Spell slot
        n = CLng(Parse(1))

        ' Prevent subscript out of range

        If n <= 0 Or n > MAX_PLAYER_SPELLS Then
            Call HackingAttempt(index, "Invalid Spell Slot")
            Exit Sub
        End If

        With Player(index).Char(Player(index).CharNum)

            If .Spell(n) = 0 Then
                Call PlayerMsg(index, "No spell here.", Red)

             Else
                Call PlayerMsg(index, "You have forgotten the spell """ & Trim$(Spell(.Spell(n)).Name) & """", Green)

                .Spell(n) = 0
                Call SendSpells(index)
            End If

        End With
        Exit Sub

        ' :::::::::::::::::::::::
        ' :: keypressed packet ::
        ' :::::::::::::::::::::::

     Case PacketID.Key

        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "Keys " & index & "," & Val(Parse$(1))
        End If

        Exit Sub

        ' :::::::::::::::::::::::
        ' :: Set access packet ::
        ' :::::::::::::::::::::::
     Case PacketID.SetAccess
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

     Case PacketID.WhosOnline
        Call SendWhosOnline(index)
        Exit Sub

     Case PacketID.OnlineList
        Call SendOnlineList
        Exit Sub

     Case PacketID.SetMOTD

        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call PutVar(App.Path & "\motd.ini", "MOTD", "Msg", Parse(1))
        Call GlobalMsg("MOTD changed to: " & Parse(1), BrightCyan)
        Call AddLog(GetPlayerName(index) & " changed MOTD to: " & Parse(1), ADMIN_LOG)
        Exit Sub

        '********************
        '* Buy Item Packet  *
        '********************
     Case PacketID.Buy
        Dim shopIndex As Integer
        Dim shopItemIndex As Integer

        'The number of the shop
        shopIndex = Val(Parse(1))
        'The number of the shop's item
        shopItemIndex = Val(Parse(2))

        'Error handling
        If shopIndex < 1 Or shopIndex > MAX_SHOPS Or shopItemIndex < 1 Or shopItemIndex > MAX_SHOP_ITEMS Then Exit Sub

        'Check to see if player's inventory is full
        'x is temp var
        X = FindOpenInvSlot(index, Shop(shopIndex).ShopItem(shopItemIndex).ItemNum)

        If X = 0 Then
            Call PlayerMsg(index, GetVar("Lang.ini", "Lang", "FullInv"), Red)
            Exit Sub
        End If

        'Check to see if they have enough currency

        If HasItem(index, Shop(shopIndex).currencyItem) >= Shop(shopIndex).ShopItem(shopItemIndex).Price Then
            'Buy the item
            TakeItem index, Shop(shopIndex).currencyItem, Shop(shopIndex).ShopItem(shopItemIndex).Price
            GiveItem index, Shop(shopIndex).ShopItem(shopItemIndex).ItemNum, Shop(shopIndex).ShopItem(shopItemIndex).Amount

            'Display message, check if it's a stackable item
            Call PlayerMsg(index, "You buy the item(s).", Yellow)
         Else
            'Can't trade
            Call PlayerMsg(index, "You can't afford that!", Red)
        End If

        Exit Sub

        'SELL ITEM PACKET
     Case PacketID.SellItem
        Dim SellItemNum As Long
        Dim SellItemSlot As Integer
        Dim SellItemAmt As Integer

        ShopNum = Val(Parse(1))
        SellItemNum = Val(Parse(2))
        SellItemSlot = Val(Parse(3))
        SellItemAmt = Val(Parse(4))

        If GetPlayerWeaponSlot(index) = Val(Parse(1)) Or GetPlayerArmorSlot(index) = Val(Parse(1)) Or GetPlayerShieldSlot(index) = Val(Parse(1)) Or GetPlayerHelmetSlot(index) = Val(Parse(1)) Or GetPlayerLegsSlot(index) = Val(Parse(1)) Or GetPlayerRingSlot(index) = Val(Parse(1)) Or GetPlayerNecklaceSlot(index) = Val(Parse(1)) Then
            Call PlayerMsg(index, "You cannot sell worn items.", Red)
            Exit Sub
        End If

        If Item(SellItemNum).Stackable = YES Then
            If SellItemAmt > GetPlayerInvItemValue(index, SellItemSlot) Then
                Call PlayerMsg(index, "You don't have enough to sell that many!", Red)
                Exit Sub
            End If

        End If

        If Item(SellItemNum).Price > 0 Then
            Call TakeItem(index, SellItemNum, SellItemAmt)
            Call GiveItem(index, Shop(ShopNum).currencyItem, Item(SellItemNum).Price * SellItemAmt)
            Call PlayerMsg(index, "The shopkeeper hands you " & Item(SellItemNum).Price * SellItemAmt & " " & Trim$(Item(Shop(ShopNum).currencyItem).Name) & ".", Yellow)
         Else
            Call PlayerMsg(index, "This item can't be sold.", Red)
        End If

        Exit Sub

        'FIX ITEM PACKET
     Case PacketID.FixItem
        ' Inv num
        n = Val(Parse(1))

        ' Make sure its a equipable item

        If Item(GetPlayerInvItemNum(index, n)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(index, n)).Type > ITEM_TYPE_NECKLACE Then
            Call PlayerMsg(index, "That item doesn't need to be fixed.", BrightRed)
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

     Case PacketID.Search
        X = Val(Parse(1))
        Y = Val(Parse(2))

        ' Prevent subscript out of range

        If X < 0 Or X > MAX_MAPX Or Y < 0 Or Y > MAX_MAPY Then
            Exit Sub
        End If

        ' Check for a player

        For i = 1 To MAX_PLAYERS

            If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) And GetPlayerX(i) = X And GetPlayerY(i) = Y Then

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

        ' Check for an npc

        For i = 1 To MAX_MAP_NPCS

            If MapNpc(GetPlayerMap(index), i).num > 0 Then
                If MapNpc(GetPlayerMap(index), i).X = X And MapNpc(GetPlayerMap(index), i).Y = Y Then
                    ' Change target
                    Player(index).Target = i
                    Player(index).TargetType = TARGET_TYPE_NPC
                    Call PlayerMsg(index, "Your target is now a " & Trim(Npc(MapNpc(GetPlayerMap(index), i).num).Name) & ".", Yellow)
                    Exit Sub
                End If

            End If
        Next i

        'Check for a onClick Tile

        If Map(GetPlayerMap(index)).tile(X, Y).Type = TILE_TYPE_ONCLICK Then
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnClick " & index & "," & Map(GetPlayerMap(index)).tile(X, Y).Data1
            End If

        End If

        BX = X
        BY = Y

        ' Check for an item

        For i = 1 To MAX_MAP_ITEMS

            If MapItem(GetPlayerMap(index), i).num > 0 Then
                If MapItem(GetPlayerMap(index), i).X = X And MapItem(GetPlayerMap(index), i).Y = Y Then
                    Call PlayerMsg(index, "You see a " & Trim(Item(MapItem(GetPlayerMap(index), i).num).Name) & ".", Yellow)
                    Exit Sub
                End If

            End If
        Next i

        Exit Sub

     Case PacketID.PlayerChat
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

        If Parse(1) = "" Then
            Call PlayerMsg(index, "Click on the player you wish to chat to first.", Pink)
            Exit Sub
        End If

        Call PlayerMsg(index, "Chat request has been sent to " & GetPlayerName(n) & ".", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " wants you to chat with them.  Type /chat to accept, or /chatdecline to decline.", Pink)

        Player(n).ChatPlayer = index
        Player(index).ChatPlayer = n
        Exit Sub

     Case PacketID.AChat
        n = Player(index).ChatPlayer

        If n < 1 Then
            Call PlayerMsg(index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If

        If Player(n).ChatPlayer <> index Then
            Call PlayerMsg(index, "Chat failed.", Pink)
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.PPChatting & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, PacketID.PPChatting & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.DChat
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

     Case PacketID.QChat
        n = Player(index).ChatPlayer

        If n < 1 Then
            Call PlayerMsg(index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.QChat & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, PacketID.QChat & SEP_CHAR & END_CHAR)

        Player(index).ChatPlayer = 0
        Player(index).InChat = 0
        Player(n).ChatPlayer = 0
        Player(n).InChat = 0
        Exit Sub

     Case PacketID.SendChat
        n = Player(index).ChatPlayer

        If n < 1 Then
            Call PlayerMsg(index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If

        Call SendDataTo(n, "sendchat" & SEP_CHAR & Parse(1) & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.PPTrade
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

     Case PacketID.ATrade
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

            Call SendDataTo(index, PacketID.PPTrading & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, PacketID.PPTrading & SEP_CHAR & END_CHAR)

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

     Case PacketID.QTrade
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
        Call SendDataTo(index, PacketID.QTrade & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, PacketID.QTrade & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.DTrade
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

     Case PacketID.UpdateTradeInv
        n = Val(Parse(1))

        Player(index).Trading(n).InvNum = Val(Parse(2))
        Player(index).Trading(n).InvName = Trim(Parse(3))

        If Player(index).Trading(n).InvNum = 0 Then
            Player(index).TradeItemMax = Player(index).TradeItemMax - 1
            Player(index).TradeOk = 0
            Player(n).TradeOk = 0
            Call SendDataTo(index, PacketID.Trading & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, PacketID.Trading & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
         Else
            Player(index).TradeItemMax = Player(index).TradeItemMax + 1
        End If

        Call SendDataTo(Player(index).TradePlayer, PacketID.UpdateTradeItem & SEP_CHAR & n & SEP_CHAR & Player(index).Trading(n).InvNum & SEP_CHAR & Player(index).Trading(n).InvName & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.SwapItems
        n = Player(index).TradePlayer

        If Player(index).TradeOk = 0 Then
            Player(index).TradeOk = 1
            Call SendDataTo(n, PacketID.Trading & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
         ElseIf Player(index).TradeOk = 1 Then
            Player(index).TradeOk = 0
            Call SendDataTo(n, PacketID.Trading & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
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
                    For X = 1 To MAX_INV

                        If GetPlayerInvItemNum(n, X) < 1 Then
                            If Player(index).Trading(i).InvNum > 0 Then
                                Call GiveItem(n, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), 1)
                                Call TakeItem(index, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), 1)
                                Exit For
                            End If

                        End If
                    Next X

                Next i

                For i = 1 To MAX_PLAYER_TRADES
                    For X = 1 To MAX_INV

                        If GetPlayerInvItemNum(index, X) < 1 Then
                            If Player(n).Trading(i).InvNum > 0 Then
                                Call GiveItem(index, GetPlayerInvItemNum(n, Player(n).Trading(i).InvNum), 1)
                                Call TakeItem(n, GetPlayerInvItemNum(n, Player(n).Trading(i).InvNum), 1)
                                Exit For
                            End If

                        End If
                    Next X

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
            Call SendDataTo(index, PacketID.QTrade & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, PacketID.QTrade & SEP_CHAR & END_CHAR)
        End If

        Exit Sub

     Case PacketID.Party
        n = FindPlayer(Parse(1))

        ' Prevent partying with self

        If n = index Then
            Exit Sub
        End If

        ' Check for a full party and if so drop it
        Dim g As Integer
        g = 0

        If Player(index).InParty = True Then

            For i = 1 To MAX_PARTY_MEMBERS
                If Player(index).Party.Member(i) > 0 Then g = g + 1
            Next i

            If g > (MAX_PARTY_MEMBERS - 1) Then
                Call PlayerMsg(index, "Party is full!", Pink)
                Exit Sub
            End If

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

            ' Check to see if player is already in a party

            If Player(n).InParty = False Then
                Call PlayerMsg(index, GetPlayerName(n) & " has been invited to your party.", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " has invited you to join their party.  Type /join to join, or /leave to decline.", Pink)

                Player(n).InvitedBy = index
             Else
                Call PlayerMsg(index, "Player is already in a party!", Pink)
            End If

         Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

        Exit Sub

     Case PacketID.JoinParty
        n = Player(index).InvitedBy

        If n > 0 Then
            ' Check to make sure they aren't the starter
            ' Check to make sure that each of there party players match
            Call PlayerMsg(index, "You have joined " & GetPlayerName(n) & "'s party!", Pink)

            If Player(n).InParty = False Then ' Set the party leader up
                Call SetPMember(n, n) 'Make them the first member and make them the leader
                Player(n).InParty = True 'Set them to be 'InParty' status
                Call SetPShare(n, True)
            End If

            Player(index).InParty = True 'Player joined
            Player(index).Party.Leader = n 'Set party leader
            Call SetPMember(n, index) 'Add the member and update the party

            ' Make sure they are in right level range

            If GetPlayerLevel(index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(index) - 5 > GetPlayerLevel(n) Then
                Call PlayerMsg(index, "There is more then a 5 level gap between you two, you will not share experience.", Pink)
                Call PlayerMsg(n, "There is more then a 5 level gap between you two, you will not share experience.", Pink)
                Call SetPShare(index, False) 'Do not share experience with party
             Else
                Call SetPShare(index, True) 'Share experience with party
            End If

            For i = 1 To MAX_PARTY_MEMBERS
                If Player(index).Party.Member(i) > 0 And Player(index).Party.Member(i) <> index Then Call PlayerMsg(Player(index).Party.Member(i), GetPlayerName(index) & " has joined your party!", Pink)
            Next i

            For i = 1 To MAX_PARTY_MEMBERS

                If Player(index).Party.Member(i) = index Then

                    For n = 1 To MAX_PARTY_MEMBERS
                        Call SendDataTo(n, PacketID.UpdateMembers & SEP_CHAR & i & SEP_CHAR & index & SEP_CHAR & END_CHAR)
                    Next n

                End If
            Next i

            For i = 1 To MAX_PARTY_MEMBERS
                Call SendDataTo(index, PacketID.UpdateMembers & SEP_CHAR & i & SEP_CHAR & Player(index).Party.Member(i) & SEP_CHAR & END_CHAR)
            Next i

         Else
            Call PlayerMsg(index, "You have not been invited into a party!", Pink)
        End If

        Exit Sub

     Case PacketID.LeaveParty
        n = Player(index).InvitedBy

        If n > 0 Or Player(index).Party.Leader = index Then
            If Player(index).InParty = True Then
                Call PlayerMsg(index, "You have left the party.", Pink)

                For i = 1 To MAX_PARTY_MEMBERS
                    If Player(index).Party.Member(i) > 0 Then Call PlayerMsg(Player(index).Party.Member(i), GetPlayerName(index) & " has left the party.", Pink)
                Next i

                Call RemovePMember(index) 'this handles removing them and updating the entire party

             Else
                Call PlayerMsg(index, "Declined party request.", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " declined your request.", Pink)

                Player(index).InParty = False
                Player(index).InvitedBy = 0

            End If
         Else
            Call PlayerMsg(index, "You are not in a party!", Pink)
        End If

        Exit Sub
    
     '//!! Unused packet
     'Case PacketID.PartyChat
     '
     '   For I = 1 To MAX_PARTY_MEMBERS
     '       If Player(index).Party.Member(I) > 0 Then Call PlayerMsg(Player(index).Party.Member(I), Parse(1), Blue)
     '   Next I
     '
     '   Exit Sub

     Case PacketID.Spells
        Call SendPlayerSpells(index)
        Exit Sub

     Case PacketID.HotScript1
        MyScript.ExecuteStatement "Scripts\Main.txt", "HotScript1 " & index
        Exit Sub

     Case "hotscript2"
        MyScript.ExecuteStatement "Scripts\Main.txt", "HotScript2 " & index
        Exit Sub

     Case "hotscript3"
        MyScript.ExecuteStatement "Scripts\Main.txt", "HotScript3 " & index
        Exit Sub

     Case "hotscript4"
        MyScript.ExecuteStatement "Scripts\Main.txt", "HotScript4 " & index
        Exit Sub

     Case PacketID.ScriptTile
        Call SendDataTo(index, PacketID.ScriptTile & SEP_CHAR & GetVar(App.Path & "\Tiles.ini", "Names", "Tile" & Parse(1)) & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.Cast
        n = Val(Parse(1))
        Call CastSpell(index, n)
        Exit Sub

     Case PacketID.RequestLocation

        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call PlayerMsg(index, "Map: " & GetPlayerMap(index) & ", X: " & GetPlayerX(index) & ", Y: " & GetPlayerY(index), Pink)
        Exit Sub

     Case PacketID.Refresh
        Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
        'Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Exit Sub

     Case PacketID.BuySprite
        ' Check if player stepped on sprite changing tile

        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
            Call PlayerMsg(index, "You need to be on a sprite tile to buy it!", BrightRed)
            Exit Sub
        End If

        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2 = 0 Then
            Call SetPlayerSprite(index, Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1)
            Call SendDataToMap(GetPlayerMap(index), PacketID.CheckSprite & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
            Exit Sub
        End If

        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(index, i) = Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2 Then
                If Item(GetPlayerInvItemNum(index, i)).Type = ITEM_TYPE_CURRENCY Then
                    If GetPlayerInvItemValue(index, i) >= Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3 Then
                        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3)

                        If GetPlayerInvItemValue(index, i) <= 0 Then
                            Call SetPlayerInvItemNum(index, i, 0)
                        End If

                        Call PlayerMsg(index, "You have bought a new sprite!", BrightGreen)
                        Call SetPlayerSprite(index, Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                        Call SendDataToMap(GetPlayerMap(index), PacketID.CheckSprite & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(index)
                    End If

                 Else

                    If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerNecklaceSlot(index) <> i Then
                        Call SetPlayerInvItemNum(index, i, 0)
                        Call PlayerMsg(index, "You have bought a new sprite!", BrightGreen)
                        Call SetPlayerSprite(index, Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                        Call SendDataToMap(GetPlayerMap(index), PacketID.CheckSprite & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(index)
                    End If

                End If

                If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerNecklaceSlot(index) <> i Then
                    Exit Sub
                End If

            End If
        Next i

        Call PlayerMsg(index, "You dont have enough to buy this sprite!", BrightRed)
        Exit Sub

     Case PacketID.ClearOwner

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call PlayerMsg(index, "Owner cleared!", BrightRed)
        Map(GetPlayerMap(index)).owner = 0
        Map(GetPlayerMap(index)).Name = "Unonwed House"
        Map(GetPlayerMap(index)).Revision = Map(GetPlayerMap(index)).Revision + 1
        Call SaveMap(GetPlayerMap(index))
        Call SendDataToMap(GetPlayerMap(index), PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Revision + 1) & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.BuyHouse
        ' Check if player stepped on house changing tile

        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_HOUSE Then
            Call PlayerMsg(index, "You need to be on a house tile to buy it!", BrightRed)
            Exit Sub
        End If

        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 0 Then
            Map(GetPlayerMap(index)).owner = GetPlayerName(index)
            Map(GetPlayerMap(index)).Name = GetPlayerName(index) & "'s House"
            Map(GetPlayerMap(index)).Revision = Map(GetPlayerMap(index)).Revision + 1
            Call SaveMap(GetPlayerMap(index))
            Call SendDataToMap(GetPlayerMap(index), PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Revision + 1) & SEP_CHAR & END_CHAR)
            Exit Sub
        End If

        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(index, i) = Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
                If Item(GetPlayerInvItemNum(index, i)).Type = ITEM_TYPE_CURRENCY Then
                    If GetPlayerInvItemValue(index, i) >= Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2 Then
                        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2)

                        If GetPlayerInvItemValue(index, i) <= 0 Then
                            Call SetPlayerInvItemNum(index, i, 0)
                        End If

                        Call PlayerMsg(index, "You have bought a new house!", BrightGreen)
                        Map(GetPlayerMap(index)).owner = GetPlayerName(index)
                        Map(GetPlayerMap(index)).Name = GetPlayerName(index) & "'s House"
                        Map(GetPlayerMap(index)).Revision = Map(GetPlayerMap(index)).Revision + 1
                        Call SaveMap(GetPlayerMap(index))
                        Call SendDataToMap(GetPlayerMap(index), PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Revision + 1) & SEP_CHAR & END_CHAR)
                        Call SendInventory(index)
                    End If

                 Else

                    If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerNecklaceSlot(index) <> i Then
                        Call SetPlayerInvItemNum(index, i, 0)
                        Call PlayerMsg(index, "You now own a new house!", BrightGreen)
                        Map(GetPlayerMap(index)).owner = GetPlayerName(index)
                        Map(GetPlayerMap(index)).Name = GetPlayerName(index) & "'s House"
                        Map(GetPlayerMap(index)).Revision = Map(GetPlayerMap(index)).Revision + 1
                        Call SaveMap(GetPlayerMap(index))
                        Call SendDataToMap(GetPlayerMap(index), PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Revision + 1) & SEP_CHAR & END_CHAR)
                        Call SendInventory(index)
                    End If

                End If

                If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerNecklaceSlot(index) <> i Then
                    Exit Sub
                End If

            End If
        Next i

        Call PlayerMsg(index, "You dont have enough to buy this house!", BrightRed)
        Exit Sub

     Case PacketID.CheckCommands
        s = Parse(1)

        If Scripting = 1 Then
            PutVar App.Path & "\Scripts\Command.ini", "TEMP", "Text" & index, Trim(s)
            MyScript.ExecuteStatement "Scripts\Main.txt", "Commands " & index
         Else
            Call PlayerMsg(index, "Thats not a valid command!", 12)
        End If

        Exit Sub

     Case PacketID.Prompt

        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerPrompt " & index & "," & Val(Parse(1)) & "," & Val(Parse(2))
        End If

        Exit Sub

     Case PacketID.QueryBox

        If Scripting = 1 Then
            Call PutVar(App.Path & "\responses.ini", "Responses", CStr(index), Parse(1))
            MyScript.ExecuteStatement "Scripts\Main.txt", "QueryBox " & index & "," & Val(Parse(2))
        End If

        Exit Sub

     Case PacketID.RequestEditArrow

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.ArrowEditor & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.EditArrow

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse(1))

        If n < 0 Or n > MAX_ARROWS Then
            Call HackingAttempt(index, "Invalid arrow Index")
            Exit Sub
        End If

        Call AddLog(GetPlayerName(index) & " editing arrow #" & n & ".", ADMIN_LOG)
        Call SendEditArrowTo(index, n)
        Exit Sub

     Case PacketID.SaveArrow

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse(1))

        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid arrow Index")
            Exit Sub
        End If

        Arrows(n).Name = Parse(2)
        Arrows(n).Pic = Val(Parse(3))
        Arrows(n).Range = Val(Parse(4))
        Arrows(n).Amount = Val(Parse(5))

        Call SendUpdateArrowToAll(n)
        Call SaveArrow(n)
        Call AddLog(GetPlayerName(index) & " saved arrow #" & n & ".", ADMIN_LOG)
        Exit Sub
    
     Case PacketID.CheckArrows
        n = Arrows(Val(Parse(1))).Pic
     
        Call SendDataToMap(GetPlayerMap(index), PacketID.CheckArrows & SEP_CHAR & index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.RequestEditEmoticon
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.EmoticonEditor & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.EndShot
        ' Prevent hacking

        If Player(index).HookShotX <> 0 Or Player(index).HookShotY <> 0 Then
            Call HackingAttempt(index, "")
            Exit Sub
        End If

        Call SetPlayerX(index, Player(index).HookShotX)
        Call SetPlayerY(index, Player(index).HookShotY)
        Call PlayerMsg(index, "You use your " & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Name & " to carefully cross over to the other side.", 3)
        Exit Sub

     Case PacketID.RequestEditElement
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.ElementEditor & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.RequestEditSkill
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.SkillEditor & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.RequestEditQuest
        ' Prevent hacking

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        Call SendDataTo(index, PacketID.QuestEditor & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.EditEmoticon

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

     Case PacketID.EditElement

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse(1))

        If n < 0 Or n > MAX_ELEMENTS Then
            Call HackingAttempt(index, "Invalid Emoticon Index")
            Exit Sub
        End If

        Call AddLog(GetPlayerName(index) & " editing element #" & n & ".", ADMIN_LOG)
        Call SendEditElementTo(index, n)
        Exit Sub

     Case PacketID.EditSkill

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse(1))

        If n < 0 Or n > MAX_SKILLS Then
            Call HackingAttempt(index, "Invalid Skill Index")
            Exit Sub
        End If

        Call AddLog(GetPlayerName(index) & " editing skill #" & n & ".", ADMIN_LOG)
        Call SendEditSkillTo(index, n)
        Exit Sub

     Case PacketID.EditQuest

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse(1))

        If n < 0 Or n > MAX_QUESTS Then
            Call HackingAttempt(index, "Invalid quest Index")
            Exit Sub
        End If

        Call AddLog(GetPlayerName(index) & " editing quest #" & n & ".", ADMIN_LOG)
        Call SendEditQuestTo(index, n)
        Exit Sub

     Case PacketID.SaveEmoticon

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

     Case PacketID.SaveSkill

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse(1))

        If n < 0 Or n > MAX_SKILLS Then
            Call HackingAttempt(index, "Invalid Skill Index")
            Exit Sub
        End If

        skill(n).Name = Parse$(2)
        skill(n).Action = Parse$(3)
        skill(n).Fail = Parse$(4)
        skill(n).Succes = Parse$(5)
        skill(n).AttemptName = Parse$(6)
        skill(n).Pictop = Val#(Parse$(7))
        skill(n).Picleft = Val#(Parse$(8))

        m = 9

        For j = 1 To MAX_SKILLS_SHEETS
            skill(n).ItemTake1num(j) = Val(Parse(m))
            skill(n).ItemTake2num(j) = Val(Parse(m + 1))
            skill(n).ItemGive1num(j) = Val(Parse(m + 2))
            skill(n).ItemGive2num(j) = Val(Parse(m + 3))
            skill(n).minlevel(j) = Val(Parse(m + 4))
            skill(n).ExpGiven(j) = Val(Parse(m + 5))
            skill(n).base_chance(j) = Val(Parse(m + 6))
            skill(n).ItemTake1val(j) = Val(Parse(m + 7))
            skill(n).ItemTake2val(j) = Val(Parse(m + 8))
            skill(n).ItemGive1val(j) = Val(Parse(m + 9))
            skill(n).ItemGive2val(j) = Val(Parse(m + 10))
            skill(n).itemequiped(j) = Val(Parse(m + 11))
            m = m + 12
        Next j

        Call SendUpdateSkillToAll(n)
        Call SaveSkill(n)
        Call AddLog(GetPlayerName(index) & " saved skill #" & n & ".", ADMIN_LOG)
        Exit Sub

     Case PacketID.SaveQuest

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse(1))

        If n < 0 Or n > MAX_QUESTS Then
            Call HackingAttempt(index, "Invalid quest Index")
            Exit Sub
        End If

        Quest(n).Name = Trim(Parse(2))
        Quest(n).Pictop = Val(Parse(3))
        Quest(n).Picleft = Val(Parse(4))

        m = 5

        For j = 0 To MAX_QUEST_LENGHT
            Quest(n).Map(j) = Val(Parse(m))
            Quest(n).X(j) = Val(Parse(m + 1))
            Quest(n).Y(j) = Val(Parse(m + 2))
            Quest(n).Npc(j) = Val(Parse(m + 3))
            Quest(n).Script(j) = Val(Parse(m + 4))
            Quest(n).ItemTake1num(j) = Val(Parse(m + 5))
            Quest(n).ItemTake2num(j) = Val(Parse(m + 6))
            Quest(n).ItemTake1val(j) = Val(Parse(m + 7))
            Quest(n).ItemTake2val(j) = Val(Parse(m + 8))
            Quest(n).ItemGive1num(j) = Val(Parse(m + 9))
            Quest(n).ItemGive2num(j) = Val(Parse(m + 10))
            Quest(n).ItemGive1val(j) = Val(Parse(m + 11))
            Quest(n).ItemGive2val(j) = Val(Parse(m + 12))
            Quest(n).ExpGiven(j) = Val(Parse(m + 13))
            m = m + 14
        Next j

        Call SendUpdateQuestToAll(n)
        Call SaveQuest(n)
        Call AddLog(GetPlayerName(index) & " saved quest #" & n & ".", ADMIN_LOG)
        Exit Sub

     Case PacketID.SaveElement

        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If

        n = Val(Parse(1))

        If n < 0 Or n > MAX_ELEMENTS Then
            Call HackingAttempt(index, "Invalid Element Index")
            Exit Sub
        End If

        Element(n).Name = Parse(2)
        Element(n).Strong = Val(Parse(3))
        Element(n).Weak = Val(Parse(4))

        Call SendUpdateElementToAll(n)
        Call SaveElement(n)
        Call AddLog(GetPlayerName(index) & " saved element #" & n & ".", ADMIN_LOG)
        Exit Sub

     Case PacketID.CheckEmoticons
        n = Emoticons(Val(Parse(1))).Pic

        Call SendDataToMap(GetPlayerMap(index), PacketID.CheckEmoticons & SEP_CHAR & index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Exit Sub

     Case PacketID.MapReport
        packs = PacketID.MapReport & SEP_CHAR

        For i = 1 To MAX_MAPS
            packs = packs & Map(i).Name & SEP_CHAR
        Next i

        packs = packs & END_CHAR

        Call SendDataTo(index, packs)
        Exit Sub

     Case PacketID.GMTime
        GameTime = Val(Parse(1))
        Call SendTimeToAll
        Exit Sub

     Case PacketID.Weather
        GameWeather = Val(Parse(1))
        Call SendWeatherToAll
        Exit Sub

     Case PacketID.WarpTo
        Call PlayerWarp(index, Val(Parse(1)), GetPlayerX(index), GetPlayerY(index))
        Exit Sub

     Case PacketID.ArrowHit
        n = Val(Parse(1))
        z = Val(Parse(2))
        X = Val(Parse(3))
        Y = Val(Parse(4))

        If n = TARGET_TYPE_PLAYER Then
            ' Make sure we dont try To attack ourselves

            If z <> index Then
                ' Can we attack the player?

                If CanAttackPlayerWithArrow(index, z) Then
                    If Not CanPlayerBlockHit(z) Then
                        ' Get the damage we can Do

                        If Not CanPlayerCriticalHit(index) Then
                            Damage = GetPlayerDamage(index) - GetPlayerProtection(z)
                            Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                         Else
                            n = GetPlayerDamage(index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(z)
                            Call BattleMsg(index, "You feel a surge of energy upon shooting!", BrightCyan, 0)
                            Call BattleMsg(z, GetPlayerName(index) & " shoots With amazing accuracy!", BrightCyan, 1)

                            'Call PlayerMsg(index, "You feel a surge of energy upon shooting!", BrightCyan)
                            'Call PlayerMsg(z, GetPlayerName(index) & " shoots With amazing accuracy!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                        End If

                        If Damage > 0 Then
                            Call AttackPlayer(index, z, Damage)
                         Else
                            Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)
                            Call BattleMsg(z, GetPlayerName(index) & "'s attack did nothing.", BrightRed, 1)

                            'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                            Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If

                     Else
                        Call BattleMsg(index, GetPlayerName(z) & " blocked your hit!", BrightCyan, 0)
                        Call BattleMsg(z, "You blocked " & GetPlayerName(index) & "'s hit!", BrightCyan, 1)

                        'Call PlayerMsg(index, GetPlayerName(z) & "'s " & Trim(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked your hit!", BrightCyan)
                        'Call PlayerMsg(z, "Your " & Trim(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                    End If

                    Exit Sub
                End If

            End If
         ElseIf n = TARGET_TYPE_NPC Then
            ' Can we attack the npc?

            If CanAttackNpcWithArrow(index, z) Then
                ' Get the damage we can Do

                If Not CanPlayerCriticalHit(index) Then
                    Damage = GetPlayerDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).DEF / 2)
                    Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                 Else
                    n = GetPlayerDamage(index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).DEF / 2)
                    Call BattleMsg(index, "You feel a surge of energy upon shooting!", BrightCyan, 0)

                    'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                    Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                End If

                If Damage > 0 Then
                    '//!! Display damage before applying
                    Call SendDataTo(index, PacketID.BlitPlayerDmg & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                    Call AttackNpc(index, z, Damage)
                 Else
                    Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)

                    'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                    Call SendDataTo(index, PacketID.BlitPlayerDmg & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                End If

                Exit Sub
            End If

         ElseIf n = TARGET_TYPE_ATTRIBUTE_NPC Then
            BX = Val(Parse(5))
            BY = Val(Parse(6))

            If Not CanPlayerCriticalHit(index) Then
                Damage = GetPlayerDamage(index) - Int(Npc(MapAttributeNpc(GetPlayerMap(index), z, BX, BY).num).DEF / 2)
                Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
             Else
                n = GetPlayerDamage(index)
                Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapAttributeNpc(GetPlayerMap(index), z, BX, BY).num).DEF / 2)
                Call BattleMsg(index, "You feel a surge of energy upon shooting!", BrightCyan, 0)

                'Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
            End If

            If Damage > 0 Then
                Call AttackAttributeNpc(z, BX, BY, index, Damage)
                'Call SendDataTo(index, PacketID.BlitPlayerDmg & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
             Else
                Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)

                'Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                'Call SendDataTo(index, PacketID.BlitPlayerDmg & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
            End If

            Exit Sub
        End If

        Exit Sub
    End Select

    Select Case Parse(0)
     Case PacketID.BankDeposit
        X = GetPlayerInvItemNum(index, Val(Parse(1)))
        i = FindOpenBankSlot(index, X)

        If i = 0 Then
            Call SendDataTo(index, PacketID.BankMsg & SEP_CHAR & "Bank full!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If

        If Val(Parse(2)) > GetPlayerInvItemValue(index, Val(Parse(1))) Then
            Call SendDataTo(index, PacketID.BankMsg & SEP_CHAR & "You cant deposit more than you have!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If

        If GetPlayerWeaponSlot(index) = Val(Parse(1)) Or GetPlayerArmorSlot(index) = Val(Parse(1)) Or GetPlayerShieldSlot(index) = Val(Parse(1)) Or GetPlayerHelmetSlot(index) = Val(Parse(1)) Or GetPlayerLegsSlot(index) = Val(Parse(1)) Or GetPlayerRingSlot(index) = Val(Parse(1)) Or GetPlayerNecklaceSlot(index) = Val(Parse(1)) Then
            Call SendDataTo(index, PacketID.BankMsg & SEP_CHAR & "You cant deposit worn equipment!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If

        If Item(X).Type = ITEM_TYPE_CURRENCY Or Item(X).Stackable = 1 Then
            If Val(Parse(2)) <= 0 Then
                Call SendDataTo(index, PacketID.BankMsg & SEP_CHAR & "You must deposit more than 0!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If

        End If

        Call TakeItem(index, X, Val(Parse(2)))
        Call GiveBankItem(index, X, Val(Parse(2)), i)

        Call SendBank(index)
        Exit Sub

     Case PacketID.BankWithdraw
        i = GetPlayerBankItemNum(index, Val(Parse(1)))
        TempVal = Val(Parse(2))
        X = FindOpenInvSlot(index, i)

        If X = 0 Then
            Call SendDataTo(index, PacketID.BankMsg & SEP_CHAR & "Inventory full!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If

        If Val(Parse(2)) > GetPlayerBankItemValue(index, Val(Parse(1))) Then
            Call SendDataTo(index, PacketID.BankMsg & SEP_CHAR & "You cant withdraw more than you have!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If

        If Item(i).Type = ITEM_TYPE_CURRENCY Or Item(i).Stackable = 1 Then
            If Val(Parse(2)) <= 0 Then
                Call SendDataTo(index, PacketID.BankMsg & SEP_CHAR & "You must withdraw more than 0!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If

        End If

        Call GiveItem(index, i, TempVal)
        Call TakeBankItem(index, i, TempVal)

        Call SendBank(index)
        Exit Sub

        'Reload the scripts
     Case PacketID.ReloadScripts

        Set MyScript = Nothing
        Set clsScriptCommands = Nothing

        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
        Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)

        Exit Sub

     Case PacketID.CustomMenuClick

        Player(index).custom_title = Parse(3)
        Player(index).custom_msg = Parse(5)

        If Scripting = 1 Then
            'MyScript.ExecuteStatement "Scripts\Main.txt", "CustomMenu " & Parse(1) & "," & Parse(2) & "," & Parse(3) & "," & Parse(4) & "," & Parse(5)
            MyScript.ExecuteStatement "Scripts\Main.txt", "menuscripts " & Parse(1) & "," & Parse(2) & "," & Parse(4)
        End If

        Exit Sub
    
     Case PacketID.ReturningCustomBoxMsg
     
        Player(index).custom_msg = Parse(1)
     
        Exit Sub

    End Select
    Call HackingAttempt(index, "")
    Exit Sub

End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)

  Dim Buffer As String
  Dim packet As String
  Dim top As String * 3
  Dim Start As Long

    If index > 0 Then
        frmServer.Socket(index).GetData Buffer, vbString, DataLength

        If Buffer = "top" Then
            top = STR(TotalOnlinePlayers)
            Call SendDataTo(index, top)
            Exit Sub
        End If

        If Buffer = "top" Then
            top = STR(TotalOnlinePlayers)
            Call SendDataTo(index, top)
            Call CloseSocket(index)
        End If

        Player(index).Buffer = Player(index).Buffer & Buffer

        Start = InStr(Player(index).Buffer, END_CHAR)

        Do While Start > 0
            packet = Mid(Player(index).Buffer, 1, Start - 1)
            Player(index).Buffer = Mid(Player(index).Buffer, Start + 1, Len(Player(index).Buffer))
            Player(index).DataPackets = Player(index).DataPackets + 1
            Start = InStr(Player(index).Buffer, END_CHAR)

            If Len(packet) > 0 Then
                Call HandleData(index, packet)
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

  Dim FileName As String
  Dim fIP As String
  Dim fName As String
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

Function IsConnected(ByVal index As Long) As Boolean

    On Error Resume Next

    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
     Else
        IsConnected = False
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

  Dim packet As String

    packet = PacketID.MapMsg & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR

    Call SendDataToMap(MapNum, packet)

End Sub

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal index As Long)

  Dim packet As String

    packet = PacketID.MapMsg2 & SEP_CHAR & Msg & SEP_CHAR & index & SEP_CHAR & END_CHAR

    Call SendDataToMap(MapNum, packet)

End Sub

Sub PlainMsg(ByVal index As Long, ByVal Msg As String, ByVal num As Long)

  Dim packet As String

    packet = PacketID.PlainMsg & SEP_CHAR & Msg & SEP_CHAR & num & SEP_CHAR & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)

  Dim packet As String

    packet = PacketID.PlayerMsg & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub SendActionNames(ByVal index As Long)

  Dim packet As String
  Dim i As Long

    packet = PacketID.ActionName & SEP_CHAR & ReadINI("ACTION", "max", App.Path & "\Data.ini") & SEP_CHAR

    For i = 1 To ReadINI("ACTION", "max", App.Path & "\Data.ini")
        packet = packet & ReadINI("ACTION", "name" & i, App.Path & "\Data.ini") & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub SendArrows(ByVal index As Long)

  Dim i As Long

    For i = 1 To MAX_ARROWS
        Call SendUpdateArrowTo(index, i)
    Next i

End Sub

Sub SendBank(ByVal index As Long)

  Dim packet As String
  Dim i As Integer

    packet = PacketID.PlayerBank & SEP_CHAR

    For i = 1 To MAX_BANK
        packet = packet & GetPlayerBankItemNum(index, i) & SEP_CHAR & GetPlayerBankItemValue(index, i) & SEP_CHAR & GetPlayerBankItemDur(index, i) & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub SendBankUpdate(ByVal index As Long, ByVal BankSlot As Long)

  Dim packet As String

    packet = PacketID.PlayerBankUpdate & SEP_CHAR & BankSlot & SEP_CHAR & GetPlayerBankItemNum(index, BankSlot) & SEP_CHAR & GetPlayerBankItemValue(index, BankSlot) & SEP_CHAR & GetPlayerBankItemDur(index, BankSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendChars(ByVal index As Long)

  Dim packet As String
  Dim i As Long

    packet = PacketID.AllChars & SEP_CHAR

    For i = 1 To MAX_CHARS
        packet = packet & Trim(Player(index).Char(i).Name) & SEP_CHAR & Trim(Class(Player(index).Char(i).Class).Name) & SEP_CHAR & Player(index).Char(i).Level & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub SendClasses(ByVal index As Long)

  Dim packet As String
  Dim i As Long

    packet = PacketID.ClassesData & SEP_CHAR & MAX_CLASSES & SEP_CHAR

    For i = 0 To MAX_CLASSES
        packet = packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).locked & SEP_CHAR & Class(i).Desc & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub DebugTrackPacket(ByVal Data As String)
Dim PacketID As Byte

    If Data = vbNullString Then Exit Sub

    PacketID = Asc(left$(Data, 1))
    Debug.Print PacketID
    PacketsOut(PacketID) = PacketsOut(PacketID) + Len(Data) - 1

End Sub

Sub SendDataTo(ByVal index As Long, ByVal Data As String)

    If IsConnected(index) Then
        DebugTrackPacket Data
        frmServer.Socket(index).SendData Data
    End If

    DoEvents

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

Sub SendEditArrowTo(ByVal index As Long, ByVal EmoNum As Long)

  Dim packet As String

    packet = PacketID.EditArrow & SEP_CHAR & EmoNum & SEP_CHAR & Trim(Arrows(EmoNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendEditElementTo(ByVal index As Long, ByVal ElementNum As Long)

  Dim packet As String

    packet = PacketID.EditElement & SEP_CHAR & ElementNum & SEP_CHAR & Trim(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendEditEmoticonTo(ByVal index As Long, ByVal EmoNum As Long)

  Dim packet As String

    packet = PacketID.EditEmoticon & SEP_CHAR & EmoNum & SEP_CHAR & Trim(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendEditItemTo(ByVal index As Long, ByVal ItemNum As Long)

  Dim packet As String

    packet = PacketID.EditItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    packet = packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
    packet = packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendEditNpcTo(ByVal index As Long, ByVal NpcNum As Long)

  Dim packet As String
  Dim i As Long

    'Packet = PacketID.EditNPC & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR
    'Packet = Packet & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & END_CHAR
    packet = PacketID.EditNPC & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).Speed & SEP_CHAR & Npc(NpcNum).Magi & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR & Npc(NpcNum).Element & SEP_CHAR & Npc(NpcNum).Spritesize & SEP_CHAR

    For i = 1 To MAX_NPC_DROPS
        packet = packet & Npc(NpcNum).ItemNPC(i).Chance
        packet = packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum
        packet = packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue & SEP_CHAR
    Next i

    packet = packet & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendEditQuestTo(ByVal index As Long, ByVal QuestNum As Long)

  Dim packet As String
  Dim j As Long

    packet = PacketID.EditQuest & SEP_CHAR & QuestNum & SEP_CHAR & Trim$(Quest(QuestNum).Name) & SEP_CHAR & Val(Quest(QuestNum).Pictop) & SEP_CHAR & Val(Quest(QuestNum).Picleft)

    For j = 0 To MAX_QUEST_LENGHT
        packet = packet & SEP_CHAR & Quest(QuestNum).Map(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).X(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).Y(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).Npc(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).Script(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake1num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake2num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake1val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake2val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive1num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive2num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive1val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive2val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ExpGiven(j)
    Next j

    packet = packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendEditShopTo(ByVal index As Long, ByVal ShopNum As Long)

  Dim packet As String
  Dim z As Integer

    packet = PacketID.EditShop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).currencyItem & SEP_CHAR

    For z = 1 To MAX_SHOP_ITEMS
        packet = packet & Shop(ShopNum).ShopItem(z).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(z).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(z).Price & SEP_CHAR
    Next z

    packet = packet & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub SendEditSkillTo(ByVal index As Long, ByVal skillNum As Long)

  Dim packet As String
  Dim j As Long

    packet = PacketID.EditSkill & SEP_CHAR & skillNum & SEP_CHAR & Trim$(skill(skillNum).Name) & SEP_CHAR & Trim$(skill(skillNum).Action) & SEP_CHAR & Trim$(skill(skillNum).Fail) & SEP_CHAR & Trim$(skill(skillNum).Succes) & SEP_CHAR & Trim$(skill(skillNum).AttemptName) & SEP_CHAR & Val(skill(skillNum).Pictop) & SEP_CHAR & Val(skill(skillNum).Picleft)

    For j = 1 To MAX_SKILLS_SHEETS
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake1num(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake2num(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive1num(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive2num(j)
        packet = packet & SEP_CHAR & skill(skillNum).minlevel(j)
        packet = packet & SEP_CHAR & skill(skillNum).ExpGiven(j)
        packet = packet & SEP_CHAR & skill(skillNum).base_chance(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake1val(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake2val(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive1val(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive2val(j)
        packet = packet & SEP_CHAR & skill(skillNum).itemequiped(j)
    Next j

    packet = packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendEditSpellTo(ByVal index As Long, ByVal SpellNum As Long)

  Dim packet As String

    packet = PacketID.EditSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Spell(SpellNum).Element & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendElements(ByVal index As Long)

  Dim i As Long

    For i = 0 To MAX_ELEMENTS
        Call SendUpdateElementTo(index, i)
    Next i

End Sub

Sub SendEmoticons(ByVal index As Long)

  Dim i As Long

    For i = 0 To MAX_EMOTICONS

        If Trim(Emoticons(i).Command) <> "" Then
            Call SendUpdateEmoticonTo(index, i)
        End If

    Next i

End Sub

Sub SendGameClockTo(ByVal index As Long)

  Dim packet As String

    packet = PacketID.GameClock & SEP_CHAR & Seconds & SEP_CHAR & Minutes & SEP_CHAR & Hours & SEP_CHAR & Gamespeed & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendGameClockToAll()

  Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SendGameClockTo(i)
        End If

    Next i

End Sub

Sub SendHP(ByVal index As Long)

  Dim packet As String

    packet = PacketID.PlayerHP & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

    packet = PacketID.PlayerPoints & SEP_CHAR & GetPlayerPOINTS(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendInventory(ByVal index As Long)

  Dim packet As String
  Dim i As Long

    packet = PacketID.PlayerInv & SEP_CHAR & index & SEP_CHAR

    For i = 1 To MAX_INV
        packet = packet & GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(index), packet)

End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)

  Dim packet As String

    packet = PacketID.PlayerInvUpdate & SEP_CHAR & InvSlot & SEP_CHAR & index & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & index & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)

End Sub

Sub SendItems(ByVal index As Long)

  Dim i As Long

    For i = 1 To MAX_ITEMS

        If Trim(Item(i).Name) <> "" Then
            Call SendUpdateItemTo(index, i)
        End If

    Next i

End Sub

Sub SendJoinMap(ByVal index As Long)

  Dim packet As String
  Dim i As Long
  Dim j As Long

    packet = ""

    ' Send all players on current map to index

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> index And GetPlayerMap(i) = GetPlayerMap(index) Then
            packet = PacketID.PlayerData & SEP_CHAR
            packet = packet & i & SEP_CHAR
            packet = packet & GetPlayerName(i) & SEP_CHAR
            packet = packet & GetPlayerSprite(i) & SEP_CHAR
            packet = packet & GetPlayerMap(i) & SEP_CHAR
            packet = packet & GetPlayerX(i) & SEP_CHAR
            packet = packet & GetPlayerY(i) & SEP_CHAR
            packet = packet & GetPlayerDir(i) & SEP_CHAR
            packet = packet & GetPlayerAccess(i) & SEP_CHAR
            packet = packet & GetPlayerPK(i) & SEP_CHAR
            packet = packet & GetPlayerGuild(i) & SEP_CHAR
            packet = packet & GetPlayerGuildAccess(i) & SEP_CHAR
            packet = packet & GetPlayerClass(i) & SEP_CHAR
            packet = packet & GetPlayerHead(i) & SEP_CHAR
            packet = packet & GetPlayerBody(i) & SEP_CHAR
            packet = packet & GetPlayerleg(i) & SEP_CHAR
            packet = packet & GetPlayerPaperdoll(i) & SEP_CHAR
            packet = packet & GetPlayerLevel(i) & SEP_CHAR

            For j = 1 To MAX_SKILLS
                packet = packet & GetPlayerSkillLvl(j, i) & SEP_CHAR
                packet = packet & GetPlayerSkillExp(j, i) & SEP_CHAR
            Next j

            packet = packet & END_CHAR
            Call SendDataTo(index, packet)
        End If

    Next i

    ' Send index's player data to everyone on the map including himself
    packet = PacketID.PlayerData & SEP_CHAR
    packet = packet & index & SEP_CHAR
    packet = packet & GetPlayerName(index) & SEP_CHAR
    packet = packet & GetPlayerSprite(index) & SEP_CHAR
    packet = packet & GetPlayerMap(index) & SEP_CHAR
    packet = packet & GetPlayerX(index) & SEP_CHAR
    packet = packet & GetPlayerY(index) & SEP_CHAR
    packet = packet & GetPlayerDir(index) & SEP_CHAR
    packet = packet & GetPlayerAccess(index) & SEP_CHAR
    packet = packet & GetPlayerPK(index) & SEP_CHAR
    packet = packet & GetPlayerGuild(index) & SEP_CHAR
    packet = packet & GetPlayerGuildAccess(index) & SEP_CHAR
    packet = packet & GetPlayerClass(index) & SEP_CHAR
    packet = packet & GetPlayerHead(index) & SEP_CHAR
    packet = packet & GetPlayerBody(index) & SEP_CHAR
    packet = packet & GetPlayerleg(index) & SEP_CHAR
    packet = packet & GetPlayerPaperdoll(index) & SEP_CHAR
    packet = packet & GetPlayerLevel(index) & SEP_CHAR

    For j = 1 To MAX_SKILLS
        packet = packet & GetPlayerSkillLvl(j, index) & SEP_CHAR
        packet = packet & GetPlayerSkillExp(j, index) & SEP_CHAR
    Next j

    packet = packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)

End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)

  Dim packet As String
  Dim j As Long

    packet = PacketID.PlayerData & SEP_CHAR
    packet = packet & index & SEP_CHAR
    packet = packet & GetPlayerName(index) & SEP_CHAR
    packet = packet & GetPlayerSprite(index) & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & GetPlayerX(index) & SEP_CHAR
    packet = packet & GetPlayerY(index) & SEP_CHAR
    packet = packet & GetPlayerDir(index) & SEP_CHAR
    packet = packet & GetPlayerAccess(index) & SEP_CHAR
    packet = packet & GetPlayerPK(index) & SEP_CHAR
    packet = packet & GetPlayerGuild(index) & SEP_CHAR
    packet = packet & GetPlayerGuildAccess(index) & SEP_CHAR
    packet = packet & GetPlayerClass(index) & SEP_CHAR
    packet = packet & GetPlayerHead(index) & SEP_CHAR
    packet = packet & GetPlayerBody(index) & SEP_CHAR
    packet = packet & GetPlayerleg(index) & SEP_CHAR
    packet = packet & GetPlayerPaperdoll(index) & SEP_CHAR
    packet = packet & GetPlayerLevel(index) & SEP_CHAR

    For j = 1 To MAX_SKILLS
        packet = packet & GetPlayerSkillLvl(j, index) & SEP_CHAR
        packet = packet & GetPlayerSkillExp(j, index) & SEP_CHAR
    Next j

    packet = packet & END_CHAR
    Call SendDataToMapBut(index, MapNum, packet)

End Sub

Sub SendLeftGame(ByVal index As Long)

  Dim packet As String
  Dim j As Long

    packet = PacketID.PlayerData & SEP_CHAR
    packet = packet & index & SEP_CHAR
    packet = packet & "" & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & "" & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & 0 & SEP_CHAR
    packet = packet & 0 & SEP_CHAR

    For j = 1 To MAX_SKILLS
        packet = packet & 0 & SEP_CHAR
        packet = packet & 0 & SEP_CHAR
    Next j

    packet = packet & END_CHAR
    Call SendDataToAllBut(index, packet)

End Sub

Sub ReadyMapPacket(MapIndex As Long)
Dim X As Long
Dim Y As Long

    '//!!
    'This was built to be used with compression, but it is a bit hard when the engine is so damn reliant on END_CHAR / SEP_CHAR
    'This sub will build a map packet and hold it in memory
    'It is only called once unlike before when it was called every time a packet was sent
    'This can be moved to runtime, but I doubt people would enjoy waiting a minute for it all to build

    'If the packet was already built, don't build it again!
    If LenB(MapPackets(MapIndex)) <> 0 Then Exit Sub

    'Build the map packet as normal, skipping the PacketID, first SEP_CHAR, last SEP_CHAR and END_CHAR
    MapPackets(MapIndex) = PacketID.MapData & SEP_CHAR & MapIndex & SEP_CHAR & Trim$(Map(MapIndex).Name) & SEP_CHAR & Map(MapIndex).Revision & SEP_CHAR & Map(MapIndex).Moral & SEP_CHAR & Map(MapIndex).Up & SEP_CHAR & Map(MapIndex).Down & SEP_CHAR & Map(MapIndex).left & SEP_CHAR & Map(MapIndex).Right & SEP_CHAR & Map(MapIndex).music & SEP_CHAR & Map(MapIndex).BootMap & SEP_CHAR & Map(MapIndex).BootX & SEP_CHAR & Map(MapIndex).BootY & SEP_CHAR & Map(MapIndex).Indoors & SEP_CHAR & Map(MapIndex).Weather & SEP_CHAR
    
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
    
            With Map(MapIndex).tile(X, Y)
                MapPackets(MapIndex) = MapPackets(MapIndex) & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & _
                    .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & _
                    .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & _
                    SEP_CHAR & .String3 & SEP_CHAR & .light & SEP_CHAR & .GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & _
                    SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & _
                    .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR
            End With
    
        Next X
    Next Y
    
    For X = 1 To MAX_MAP_NPCS
        MapPackets(MapIndex) = MapPackets(MapIndex) & Map(MapIndex).Npc(X) & SEP_CHAR
    Next X

    MapPackets(MapIndex) = MapPackets(MapIndex) & END_CHAR
    
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)

  Dim packet As String
  Dim P1 As String
  Dim P2 As String
  Dim X As Integer
  Dim Y As Integer

    ReadyMapPacket MapNum

    Call SendDataTo(index, MapPackets(MapNum))

End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Long)

  Dim packet As String
  Dim i As Long

    packet = PacketID.MapItemData & SEP_CHAR

    For i = 1 To MAX_MAP_ITEMS

        If MapNum > 0 Then
            packet = packet & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).Y & SEP_CHAR
        End If

    Next i
    packet = packet & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)

  Dim packet As String
  Dim i As Long

    packet = PacketID.MapItemData & SEP_CHAR

    For i = 1 To MAX_MAP_ITEMS
        packet = packet & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).X & SEP_CHAR & MapItem(MapNum, i).Y & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataToMap(MapNum, packet)

End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)

  Dim packet As String
  Dim i As Long

    packet = PacketID.MapNPCData & SEP_CHAR

    For i = 1 To MAX_MAP_NPCS

        If MapNum > 0 Then
            packet = packet & MapNpc(MapNum, i).num & SEP_CHAR & MapNpc(MapNum, i).X & SEP_CHAR & MapNpc(MapNum, i).Y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
        End If

    Next i
    packet = packet & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)

  Dim packet As String
  Dim i As Long

    packet = PacketID.MapNPCData & SEP_CHAR

    For i = 1 To MAX_MAP_NPCS
        packet = packet & MapNpc(MapNum, i).num & SEP_CHAR & MapNpc(MapNum, i).X & SEP_CHAR & MapNpc(MapNum, i).Y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataToMap(MapNum, packet)

End Sub

Sub SendMP(ByVal index As Long)

  Dim packet As String

    packet = PacketID.PlayerMP & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendNewCharClasses(ByVal index As Long)

  Dim packet As String
  Dim i As Long

    packet = PacketID.NewCharClasses & SEP_CHAR & MAX_CLASSES & SEP_CHAR & ClassesOn & SEP_CHAR

    For i = 0 To MAX_CLASSES
        packet = packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).MaleSprite & SEP_CHAR & Class(i).FemaleSprite & SEP_CHAR & Class(i).locked & SEP_CHAR & Class(i).Desc & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub SendNewsTo(ByVal index As Long)

  Dim packet As String
  Dim Red As Integer
  Dim Green As Integer
  Dim Blue As Integer

    On Error GoTo NewsError
    Red = Val(ReadINI("COLOR", "Red", App.Path & "\News.ini"))
    Green = Val(ReadINI("COLOR", "Green", App.Path & "\News.ini"))
    Blue = Val(ReadINI("COLOR", "Blue", App.Path & "\News.ini"))

    packet = PacketID.News & SEP_CHAR & ReadINI("DATA", "ServerNews", App.Path & "\News.ini") & SEP_CHAR
    packet = packet & Red & SEP_CHAR & Green & SEP_CHAR & Blue & SEP_CHAR & ReadINI("DATA", "Desc", App.Path & "\News.ini") & END_CHAR

    Call SendDataTo(index, packet)
    Exit Sub

NewsError:
    'Error reading the news, so just send white
    Red = 255
    Green = 255
    Blue = 255

    packet = PacketID.News & SEP_CHAR & ReadINI("DATA", "ServerNews", App.Path & "\News.ini") & SEP_CHAR
    packet = packet & Red & SEP_CHAR & Green & SEP_CHAR & Blue & SEP_CHAR & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub SendNpcs(ByVal index As Long)

  Dim i As Long

    For i = 1 To MAX_NPCS

        If Trim(Npc(i).Name) <> "" Then
            Call SendUpdateNpcTo(index, i)
        End If

    Next i

End Sub

Sub SendOnlineList()

  Dim packet As String
  Dim i As Long
  Dim n As Long

    packet = ""
    n = 0

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            packet = packet & SEP_CHAR & GetPlayerName(i) & SEP_CHAR
            n = n + 1
        End If

    Next i

    packet = PacketID.OnlineList & SEP_CHAR & n & packet & END_CHAR

    Call SendDataToAll(packet)

End Sub

Sub SendPlayerData(ByVal index As Long)

  Dim packet As String
  Dim j As Long

    ' Send index's player data to everyone including himself on th emap
    packet = PacketID.PlayerData & SEP_CHAR
    packet = packet & index & SEP_CHAR
    packet = packet & GetPlayerName(index) & SEP_CHAR
    packet = packet & GetPlayerSprite(index) & SEP_CHAR
    packet = packet & GetPlayerMap(index) & SEP_CHAR
    packet = packet & GetPlayerX(index) & SEP_CHAR
    packet = packet & GetPlayerY(index) & SEP_CHAR
    packet = packet & GetPlayerDir(index) & SEP_CHAR
    packet = packet & GetPlayerAccess(index) & SEP_CHAR
    packet = packet & GetPlayerPK(index) & SEP_CHAR
    packet = packet & GetPlayerGuild(index) & SEP_CHAR
    packet = packet & GetPlayerGuildAccess(index) & SEP_CHAR
    packet = packet & GetPlayerClass(index) & SEP_CHAR
    packet = packet & GetPlayerHead(index) & SEP_CHAR
    packet = packet & GetPlayerBody(index) & SEP_CHAR
    packet = packet & GetPlayerleg(index) & SEP_CHAR
    packet = packet & GetPlayerPaperdoll(index) & SEP_CHAR
    packet = packet & GetPlayerLevel(index) & SEP_CHAR

    For j = 1 To MAX_SKILLS
        packet = packet & GetPlayerSkillLvl(j, index) & SEP_CHAR
        packet = packet & GetPlayerSkillExp(j, index) & SEP_CHAR
    Next j

    packet = packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)

End Sub

Sub SendPlayerLevelToAll(ByVal index As Long)

  Dim packet As String

    packet = PacketID.PlayerLevel & SEP_CHAR & index & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)

End Sub

Sub SendPlayerSpells(ByVal index As Long)

  Dim packet As String
  Dim i As Long

    packet = PacketID.Spells & SEP_CHAR

    For i = 1 To MAX_PLAYER_SPELLS
        packet = packet & GetPlayerSpell(index, i) & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub SendPlayerXY(ByVal index As Long)

  Dim packet As String

    packet = PacketID.PlayerXY & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendQuests(ByVal index As Long)

  Dim i As Long

    For i = 1 To MAX_QUESTS
        Call SendUpdateQuestTo(index, i)
    Next i

End Sub

Sub SendShops(ByVal index As Long)

  Dim i As Long

    For i = 1 To MAX_SHOPS

        If Trim(Shop(i).Name) <> "" Then
            Call SendUpdateShopTo(index, i)
        End If

    Next i

End Sub

Sub SendSkills(ByVal index As Long)

  Dim i As Long

    For i = 1 To MAX_SKILLS
        Call SendUpdateSkillTo(index, i)
    Next i

End Sub

Sub SendSP(ByVal index As Long)

  Dim packet As String

    packet = PacketID.PlayerSP & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & SEP_CHAR & END_CHAR
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

Sub Sendsprite(ByVal index As Long, ByVal indexto As Long)

  Dim packet As String

    packet = PacketID.CusSprite & SEP_CHAR & index & SEP_CHAR & Player(index).Char(Player(index).CharNum).head & SEP_CHAR & Player(index).Char(Player(index).CharNum).body & SEP_CHAR & Player(index).Char(Player(index).CharNum).leg & SEP_CHAR & END_CHAR
    Call SendDataTo(indexto, packet)

End Sub

Sub SendStats(ByVal index As Long)

  Dim packet As String

    packet = PacketID.PlayerStatsPacket & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & GetPlayerNextLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendTimeTo(ByVal index As Long)

  Dim packet As String

    packet = PacketID.Time & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendTimeToAll()

  Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SendTimeTo(i)
        End If

    Next i

    Call SpawnAllMapNpcs

End Sub


'Updated with new shop system -Pickle
Sub SendTrade(ByVal index As Long, ByVal ShopNum As Long)

  Dim packet As String

    packet = PacketID.GoShop & SEP_CHAR & ShopNum & SEP_CHAR & END_CHAR
    'All we need are the shop identifier and num - we don't need to send the entire shop every time
    Call SendDataTo(index, packet)

End Sub

Sub SendUpdateArrowTo(ByVal index As Long, ByVal ItemNum As Long)

  Dim packet As String

    packet = PacketID.UpdateArrow & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)

  Dim packet As String

    packet = PacketID.UpdateArrow & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)

End Sub

Sub SendUpdateElementTo(ByVal index As Long, ByVal ElementNum As Long)

  Dim packet As String

    packet = PacketID.UpdateElement & SEP_CHAR & ElementNum & SEP_CHAR & Trim(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendUpdateElementToAll(ByVal ElementNum As Long)

  Dim packet As String

    packet = PacketID.UpdateElement & SEP_CHAR & ElementNum & SEP_CHAR & Trim(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)

End Sub

Sub SendUpdateEmoticonTo(ByVal index As Long, ByVal ItemNum As Long)

  Dim packet As String

    packet = PacketID.UpdateEmoticon & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)

  Dim packet As String

    packet = PacketID.UpdateEmoticon & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)

End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)

  Dim packet As String

    'Packet = PacketID.UpdateItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    packet = PacketID.UpdateItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    packet = packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
    packet = packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)

  Dim packet As String

    'Packet = PacketID.UpdateItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    packet = PacketID.UpdateItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    packet = packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
    packet = packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)

End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NpcNum As Long)

  Dim packet As String

    packet = PacketID.UpdateNPC & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Spritesize & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)

  Dim packet As String

    packet = PacketID.UpdateNPC & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Spritesize & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)

End Sub

Sub SendUpdatePlayerSkill(ByVal index As Integer, ByVal skillNum As Integer)

  Dim packet As String

    packet = PacketID.SkillInfo & SEP_CHAR & skillNum & SEP_CHAR & GetPlayerSkillExp(index, skillNum) & SEP_CHAR & GetPlayerSkillLvl(index, skillNum) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendUpdateQuestTo(ByVal index As Long, ByVal QuestNum As Long)

  Dim packet As String
  Dim j As Long

    packet = PacketID.UpdateQuest & SEP_CHAR & QuestNum & SEP_CHAR & Trim$(Quest(QuestNum).Name) & SEP_CHAR & Val(Quest(QuestNum).Pictop) & SEP_CHAR & Val(Quest(QuestNum).Picleft)

    For j = 0 To MAX_QUEST_LENGHT
        packet = packet & SEP_CHAR & Quest(QuestNum).Map(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).X(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).Y(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).Npc(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).Script(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake1num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake2num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake1val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake2val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive1num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive2num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive1val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive2val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ExpGiven(j)
    Next j

    packet = packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendUpdateQuestToAll(ByVal QuestNum As Long)

  Dim packet As String
  Dim j As Long

    packet = PacketID.UpdateQuest & SEP_CHAR & QuestNum & SEP_CHAR & Trim$(Quest(QuestNum).Name) & SEP_CHAR & Val(Quest(QuestNum).Pictop) & SEP_CHAR & Val(Quest(QuestNum).Picleft)

    For j = 0 To MAX_QUEST_LENGHT
        packet = packet & SEP_CHAR & Quest(QuestNum).Map(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).X(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).Y(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).Npc(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).Script(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake1num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake2num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake1val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake2val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive1num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive2num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive1val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive2val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ExpGiven(j)
    Next j

    packet = packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)

End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum)

  Dim packet As String
  Dim i As Integer

    packet = PacketID.UpdateShop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).currencyItem & SEP_CHAR

    For i = 1 To MAX_SHOP_ITEMS
        packet = packet & Shop(ShopNum).ShopItem(i).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(i).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(i).Price & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)

  Dim packet As String
  Dim i As Integer

    packet = PacketID.UpdateShop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).currencyItem & SEP_CHAR

    For i = 1 To MAX_SHOP_ITEMS
        packet = packet & Shop(ShopNum).ShopItem(i).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(i).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(i).Price & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataToAll(packet)

End Sub

Sub SendUpdateSkillTo(ByVal index As Long, ByVal skillNum As Long)

  Dim packet As String
  Dim j As Long

    packet = PacketID.UpdateSkill & SEP_CHAR & skillNum & SEP_CHAR & Trim$(skill(skillNum).Name) & SEP_CHAR & Trim$(skill(skillNum).Action) & SEP_CHAR & Trim$(skill(skillNum).Fail) & SEP_CHAR & Trim$(skill(skillNum).Succes) & SEP_CHAR & Trim$(skill(skillNum).AttemptName) & SEP_CHAR & Val(skill(skillNum).Pictop) & SEP_CHAR & Val(skill(skillNum).Picleft)

    For j = 1 To MAX_SKILLS_SHEETS
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake1num(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake2num(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive1num(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive2num(j)
        packet = packet & SEP_CHAR & skill(skillNum).minlevel(j)
        packet = packet & SEP_CHAR & skill(skillNum).ExpGiven(j)
        packet = packet & SEP_CHAR & skill(skillNum).base_chance(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake1val(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake2val(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive1val(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive2val(j)
        packet = packet & SEP_CHAR & skill(skillNum).itemequiped(j)
    Next j

    packet = packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendUpdateSkillToAll(ByVal skillNum As Long)

  Dim packet As String
  Dim j As Long

    packet = PacketID.UpdateSkill & SEP_CHAR & skillNum & SEP_CHAR & Trim$(skill(skillNum).Name) & SEP_CHAR & Trim$(skill(skillNum).Action) & SEP_CHAR & Trim$(skill(skillNum).Fail) & SEP_CHAR & Trim$(skill(skillNum).Succes) & SEP_CHAR & Trim$(skill(skillNum).AttemptName) & SEP_CHAR & Val(skill(skillNum).Pictop) & SEP_CHAR & Val(skill(skillNum).Picleft)

    For j = 1 To MAX_SKILLS_SHEETS
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake1num(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake2num(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive1num(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive2num(j)
        packet = packet & SEP_CHAR & skill(skillNum).minlevel(j)
        packet = packet & SEP_CHAR & skill(skillNum).ExpGiven(j)
        packet = packet & SEP_CHAR & skill(skillNum).base_chance(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake1val(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemTake2val(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive1val(j)
        packet = packet & SEP_CHAR & skill(skillNum).ItemGive2val(j)
        packet = packet & SEP_CHAR & skill(skillNum).itemequiped(j)
    Next j

    packet = packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)

End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)

  Dim packet As String

    packet = PacketID.UpdateSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)

  Dim packet As String

    packet = PacketID.UpdateSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(packet)

End Sub

Sub SendWeatherTo(ByVal index As Long)

  Dim packet As String

    If RainIntensity <= 0 Then RainIntensity = 1
    packet = PacketID.Weather & SEP_CHAR & GameWeather & SEP_CHAR & RainIntensity & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

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

    frmServer.Label5.caption = "Current Weather: " & Weather

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SendWeatherTo(i)
        End If

    Next i

End Sub

Sub SendWhosOnline(ByVal index As Long)

  Dim s As String
  Dim n As Long
  Dim i As Long

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

Sub SendWierdTo(ByVal index As Long)

  Dim packet As String

    packet = PacketID.Wierd & SEP_CHAR & Wierd & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendWierdToAll()

  Dim i As Long

    If Wierd = 1 Then
        Wierd = 0
        MsgBox ("Wierdify is turned OFF")
     Else
        Wierd = 1
        MsgBox ("Wierdify is turned ON")
    End If

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SendWierdTo(i)
        End If

    Next i

End Sub

Sub SendWornEquipment(ByVal index As Long)

  Dim packet As String

    If IsPlaying(index) Then
        packet = PacketID.PlayerWornEQ & SEP_CHAR & index & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & GetPlayerLegsSlot(index) & SEP_CHAR & GetPlayerRingSlot(index) & SEP_CHAR & GetPlayerNecklaceSlot(index) & SEP_CHAR & END_CHAR
        Call SendDataToMap(GetPlayerMap(index), packet)
    End If

End Sub

Sub SocketConnected(ByVal index As Long)

    If index <> 0 Then
        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then

        If Not IsBanned(GetPlayerIP(index)) Then
            Call TextAdd(frmServer.txtText(0), "Received connection from " & GetPlayerIP(index) & ".", True)
         Else
            'LMAO PWNED ROFL!!!!11/
            Call AlertMsg(index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
        End If

        'Else
        ' Tried multiple connections
        '    Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
        'End If
    End If

End Sub

Sub UpdateCaption()

    frmServer.caption = GAME_NAME & " - Eclipse Evolution Server"
    frmServer.lblIP.caption = "Ip Address: " & frmServer.Socket(0).LocalIP
    frmServer.lblPort.caption = "Port: " & STR(frmServer.Socket(0).LocalPort)
    frmServer.TPO.caption = "Total Players Online: " & TotalOnlinePlayers

End Sub

