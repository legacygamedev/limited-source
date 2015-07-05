Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = GAME_NAME & " - Ramza Engine Servidor - "
    frmServer.lblIP.Caption = "IP: " & frmServer.Socket(0).LocalIP
    frmServer.lblPort.Caption = "Puerto: " & STR(frmServer.Socket(0).LocalPort)
    frmServer.TPO.Caption = "Jugadores Online: " & TotalOnlinePlayers
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
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
Dim text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
    Call CloseSocket(index)
End Sub

Sub PlainMsg(ByVal index As Long, ByVal Msg As String, ByVal num As Long)
Dim Packet As String

    Packet = "PLAINMSG" & SEP_CHAR & Msg & SEP_CHAR & num & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " fue reiniciado por (" & Reason & ")", White)
        End If
    
        Call AlertMsg(index, "Perdiste tu coneccion con " & GAME_NAME & ".")
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
                Call TextAdd(frmServer.txtText(0), "Reciviendo coneccion de " & GetPlayerIP(index) & ".", True)
            Else
                Call AlertMsg(index, "Fuiste baneado por " & GAME_NAME & ", y no puedes seguir jugando.")
            End If
        'Else
           ' Tried multiple connections
        '    Call AlertMsg(Index, GAME_NAME & " no admite multiples ip.")
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
Dim Amount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Movement As Long
Dim i As Long, n As Long, x As Long, y As Long, f As Long, o As Long, l As Long
Dim c As Long, c1 As Long, c2 As Long, c3 As Long, c4 As Long, c5 As Long, c6 As Long, c7 As Long, c8 As Long
Dim d As Long, d1 As Long, d2 As Long, d3 As Long, d4 As Long, d5 As Long, d6 As Long, d7 As Long, d8 As Long
Dim m As Long, m1 As Long, m2 As Long, m3 As Long, m4 As Long, m5 As Long, m6 As Long, m7 As Long, m8 As Long
Dim MapNum As Long
Dim s As String
Dim tMapStart As Long, tMapEnd As Long
Dim ShopNum As Long, ItemNum As Long
Dim DurNeeded As Long, GoldNeeded As Long
Dim z As Long
Dim Packet As String
Dim BX As Long, BY As Long
Dim TempNum As Long, TempVal As Long

    ' Handle Data
    Parse = Split(Data, SEP_CHAR)

' Parse's Without Being Online
If Not IsPlaying(index) Then
    Select Case LCase(Parse(0))
        Case "gatglasses"
            Call SendNewCharClasses(index)
            Exit Sub
            
        Case "newfaccountied"
            If Not IsLoggedIn(index) Then
                Name = Parse(1)
                Password = Parse(2)
                        
                For i = 1 To Len(Name)
                    n = Asc(Mid(Name, i, 1))
                    
                    If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                    Else
                        Call PlainMsg(index, "Nombre invalido solo letras, numeros, espacios y _ permitidos.", 1)
                        Exit Sub
                    End If
                Next i
                
                If Not AccountExist(Name) Then
                    Call AddAccount(index, Name, Password)
                    Call TextAdd(frmServer.txtText(0), "La cuenta de " & Name & " fue creada.", True)
                    Call AddLog("La cuenta de " & Name & " fue creada.", PLAYER_LOG)
                    Call PlainMsg(index, "Tu cuenta fue creada!", 1)
                Else
                    Call PlainMsg(index, "Ese nombre de cuenta ya esta tomado!", 1)
                End If
            End If
            Exit Sub
        
        Case "delimaccounted"
            If Not IsLoggedIn(index) Then
                Name = Parse(1)
                Password = Parse(2)
                            
                If Not AccountExist(Name) Then
                    Call PlainMsg(index, "Ese nombre de cuenta no existe.", 2)
                    Exit Sub
                End If
                
                If Not PasswordOK(Name, Password) Then
                    Call PlainMsg(index, "Contraseña incorrecta.", 2)
                    Exit Sub
                End If
                            
                Call LoadPlayer(index, Name)
                For i = 1 To MAX_CHARS
                    If Trim(Player(index).Char(i).Name) <> "" Then
                        Call DeleteName(Player(index).Char(i).Name)
                    End If
                Next i
                Call ClearPlayer(index)
                
                Call Kill(App.Path & "\accounts\" & Trim(Name) & ".ini")
                Call AddLog("La cuenta " & Trim(Name) & " fue borrada.", PLAYER_LOG)
                Call PlainMsg(index, "Tu cuenta fue borrada.", 2)
            End If
            Exit Sub
            
        Case "logination"
            If Not IsLoggedIn(index) Then
                Name = Parse(1)
                Password = Parse(2)
                
                For i = 1 To Len(Name)
                    n = Asc(Mid(Name, i, 1))
                    
                    If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                    Else
                        Call PlainMsg(index, "El duplicado de cuentas no esta admitido!", 3)
                    Exit Sub
                    End If
                Next i
            
                If Val(Parse(3)) < CLIENT_MAJOR Or Val(Parse(4)) < CLIENT_MINOR Or Val(Parse(5)) < CLIENT_REVISION Then
                    Call PlainMsg(index, "Version antigua, visita " & Trim(GetVar(App.Path & "\Data.ini", "CONFIG", "WebSite")), 3)
                    Exit Sub
                End If
                            
                If Not AccountExist(Name) Then
                    Call PlainMsg(index, "Ese nombre de cuenta no existe.", 3)
                    Exit Sub
                End If
            
                If Not PasswordOK(Name, Password) Then
                    Call PlainMsg(index, "Contraseña incorrecta.", 3)
                    Exit Sub
                End If
            
                If IsMultiAccounts(Name) Then
                    Call PlainMsg(index, "No esta autorizado conectarse con multiples cuentas.", 3)
                    Exit Sub
                End If
                
                If frmServer.Closed.Value = Checked Then
                    Call PlainMsg(index, "El servidor esta cerrado por el momento!", 3)
                    Exit Sub
                End If
                    
                If Parse(6) = "jwehiehfojcvnvnsdinaoiwheoewyriusdyrflsdjncjkxzncisdughfusyfuapsipiuahfpaijnflkjnvjnuahguiryasbdlfkjblsahgfauygewuifaunfauk" And Parse(7) = "ksisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlasndowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas" And Parse(8) = "saiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi" And Val(Parse(9)) = "88978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672" Then
                Else
                    Call AlertMsg(index, "Hack Attempt Alert!")
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
                           Call PutVar(App.Path & "\banks\" & Trim(Name) & ".ini", "CHAR" & i, "BankItemNum" & n, STR(Player(index).Char(i).Bank(n).num))
                           Call PutVar(App.Path & "\banks\" & Trim(Name) & ".ini", "CHAR" & i, "BankItemVal" & n, STR(Player(index).Char(i).Bank(n).Value))
                           Call PutVar(App.Path & "\banks\" & Trim(Name) & ".ini", "CHAR" & i, "BankItemDur" & n, STR(Player(index).Char(i).Bank(n).Dur))
                       Next n
                   Next i
               End If
        
                Call LoadPlayer(index, Name)
                Call SendChars(index)
        
                Call AddLog(GetPlayerLogin(index) & " se conecto desde " & GetPlayerIP(index) & ".", PLAYER_LOG)
                Call TextAdd(frmServer.txtText(0), GetPlayerLogin(index) & " se conecto desde " & GetPlayerIP(index) & ".", True)
            End If
            Exit Sub
    
        Case "addachara"
                Name = Parse(1)
                Sex = Val(Parse(2))
                Class = Val(Parse(3))
                CharNum = Val(Parse(4))
                        
                If LCase(Trim(Name)) = "Ramza" Then
                    Call PlainMsg(index, "Vamos a dejar una cosa en claro, no sos yo, ok? :)", 4)
                    Exit Sub
                End If
                                
                For i = 1 To Len(Name)
                    n = Asc(Mid(Name, i, 1))
                    
                    If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                    Else
                        Call PlainMsg(index, "Nombre invalido solo letras, numeros, espacios y _ estan permitidos.", 4)
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
                
                If Class < 0 Or Class > Max_Classes Then
                    Call HackingAttempt(index, "Invalid Class")
                    Exit Sub
                End If
    
                If CharExist(index, CharNum) Then
                    Call PlainMsg(index, "El personaje ya existe!", 4)
                    Exit Sub
                End If
    
                If FindChar(Name) Then
                    Call PlainMsg(index, "Perdon, pero ese nombre esta en uso!", 4)
                    Exit Sub
                End If
    
                Call AddChar(index, Name, Sex, Class, CharNum)
                Call SavePlayer(index)
                Call AddLog("El personaje " & Name & " fue agregado a la cuenta de " & GetPlayerLogin(index) & ".", PLAYER_LOG)
                Call SendChars(index)
                Call PlainMsg(index, "El personaje fue creado!", 5)
            Exit Sub
    
        Case "delimbocharu"
                CharNum = Val(Parse(1))
    
                If CharNum < 1 Or CharNum > MAX_CHARS Then
                    Call HackingAttempt(index, "Invalid CharNum")
                    Exit Sub
                End If
                
                Call DelChar(index, CharNum)
                Call AddLog("Personaje borrado en la cuenta de " & GetPlayerLogin(index) & ".", PLAYER_LOG)
                Call SendChars(index)
                Call PlainMsg(index, "El personaje fue borrado!", 5)
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
                            Call PlainMsg(index, "El servidor esta solo abierto para GMs por ahora!", 5)
                            'Call HackingAttempt(Index, "El servidor esta solo abierto para GMs por ahora!")
                            Exit Sub
                        End If
                    End If
                                        
                    Call JoinGame(index)
                
                    CharNum = Player(index).CharNum
                    Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " empezo a jugar " & GAME_NAME & ".", PLAYER_LOG)
                    Call TextAdd(frmServer.txtText(0), GetPlayerLogin(index) & "/" & GetPlayerName(index) & " empezo a jugar " & GAME_NAME & ".", True)
                    Call UpdateCaption
                    
                    If Not FindChar(GetPlayerName(index)) Then
                        f = FreeFile
                        Open App.Path & "\accounts\charlist.txt" For Append As #f
                            Print #f, GetPlayerName(index)
                        Close #f
                    End If
                Else
                    Call PlainMsg(index, "El personaje no existe!", 5)
                End If
            Exit Sub
    End Select
End If
        
' Parse's With Being Online And Playing
If IsPlaying(index) = False Then Exit Sub
If IsConnected(index) = False Then Exit Sub
Select Case LCase(Parse(0))
    ' :::::::::::::::::::
    ' :: Guilds Packet ::
    ' :::::::::::::::::::
    ' Access
    Case "guildchangeaccess"
        ' Check the requirements.
        If FindPlayer(Parse(1)) = 0 Then
            Call PlayerMsg(index, "El jugador esta desconectado", White)
            Exit Sub
        End If
    
        If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(index) Then
            Call PlayerMsg(index, "El jugador no esta en tu Clan", Red)
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
            Call PlayerMsg(index, "El jugador esta desconectado", White)
            Exit Sub
        End If
        
        If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(index) Then
            Call PlayerMsg(index, "El jugador no esta en tu Clan", Red)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(FindPlayer(Parse(1))) > GetPlayerGuildAccess(index) Then
            Call PlayerMsg(index, "El jugador tiene un nivel de Clan mas alto que el tuyo.", Red)
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
            Call PlayerMsg(index, "No estas en un Clan.", Red)
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
            Call PlayerMsg(index, "El jugador esta desconectado", White)
            Exit Sub
        End If
        
        ' Check if they are alredy in a guild
            If GetPlayerGuild(FindPlayer(Parse(1))) <> "" Then
            Call PlayerMsg(index, "El jugador ya pertenece a un Clan", Red)
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
            Call PlayerMsg(index, "El jugador esta desconectado", White)
            Exit Sub
        End If
        
        If GetPlayerGuild(FindPlayer(Parse(1))) <> GetPlayerGuild(index) Then
            Call PlayerMsg(index, "El jugador no esta en tu Clan", Red)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(FindPlayer(Parse(1))) > 1 Then
            Call PlayerMsg(index, "Ese jugador ya fue admitido", Red)
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
            Call PlayerMsg(index, "El jugador esta desconectado", White)
            Exit Sub
        End If
        
        If GetPlayerGuild(FindPlayer(Parse(1))) <> "" Then
            Call PlayerMsg(index, "El jugador ya pertenece a un Clan", Red)
            Exit Sub
        End If
        
        'It is possible, so set the guild to index's guild, and the access level to 0
        Call SetPlayerGuild(FindPlayer(Parse(1)), GetPlayerGuild(index))
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub
        
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    Case "saymsg"
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Say Text Modification")
                Exit Sub
            End If
        Next i
        
        If frmServer.chkM.Value = Unchecked Then
            If GetPlayerAccess(index) <= 0 Then
                Call PlayerMsg(index, "Los mensajes de mapa fueron desactivados por el servidor!", BrightRed)
                Exit Sub
            End If
        End If
        
        Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " : " & Msg & "", PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " : " & Msg & "", SayColor)
        Call MapMsg2(GetPlayerMap(index), Msg, index)
        TextAdd frmServer.txtText(3), GetPlayerName(index) & " En el Mapa " & GetPlayerMap(index) & ": " & Msg, True
        Exit Sub

    Case "emotemsg"
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Emote Text Modification")
                Exit Sub
            End If
        Next i
        
        If frmServer.chkE.Value = Unchecked Then
            If GetPlayerAccess(index) <= 0 Then
                Call PlayerMsg(index, "Los mensajes de emoticones fueron desactivados por el servidor!", BrightRed)
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
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Broadcast Text Modification")
                Exit Sub
            End If
        Next i
        
        If frmServer.chkBC.Value = Unchecked Then
            If GetPlayerAccess(index) <= 0 Then
                Call PlayerMsg(index, "Los mensajes de broadcast fueron desactivados por el servidor!", BrightRed)
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
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Global Text Modification")
                Exit Sub
            End If
        Next i
        
        If frmServer.chkG.Value = Unchecked Then
            If GetPlayerAccess(index) <= 0 Then
                Call PlayerMsg(index, "Los mensajes globales fueron desactivados por el servidor!", BrightRed)
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
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Admin Text Modification")
                Exit Sub
            End If
        Next i
        
        If frmServer.chkA.Value = Unchecked Then
            Call PlayerMsg(index, "Los mensajes de admin fueron desactivados por el servidor!", BrightRed)
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
           If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
               Call HackingAttempt(index, "Player Msg Text Modification")
               Exit Sub
           End If
       Next i
       
       If frmServer.chkP.Value = Unchecked Then
           If GetPlayerAccess(index) <= 0 Then
               Call PlayerMsg(index, "Los mensajes de PM fueron desactivados por el servidor!", BrightRed)
               Exit Sub
           End If
       End If
       
       If FindPlayer(Parse(1)) = 0 Then
           Call PlayerMsg(index, "El jugador esta desconectado", White)
           Exit Sub
       End If
       
       ' Check if they are trying to talk to themselves
       If MsgTo <> index Then
           If MsgTo > 0 Then
               Call AddLog(GetPlayerName(index) & " dice " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
               Call PlayerMsg(MsgTo, GetPlayerName(index) & " te dice, '" & Msg & "'", TellColor)
               Call PlayerMsg(index, "Le dices a " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
           Else
               Call PlayerMsg(index, "El jugador esta desconectado.", White)
           End If
       Else
           Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " begins to mumble to himself, what a wierdo...", PLAYER_LOG)
           Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " begins to mumble to himself, what a wierdo...", Green)
       End If
       TextAdd frmServer.txtText(4), "Para " & GetPlayerName(MsgTo) & " De " & GetPlayerName(index) & ": " & Msg, True
       Exit Sub
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(index, "Player Msg Text Modification")
                Exit Sub
            End If
        Next i
        
        If frmServer.chkP.Value = Unchecked Then
            If GetPlayerAccess(index) <= 0 Then
                Call PlayerMsg(index, "Los mensajes de PM fueron desactivados por el servidor!", BrightRed)
                Exit Sub
            End If
        End If
        
        ' Check if they are trying to talk to themselves
        If MsgTo <> index Then
            If MsgTo > 0 Then
                Call AddLog(GetPlayerName(index) & " dice " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
                Call PlayerMsg(MsgTo, GetPlayerName(index) & " te dice, '" & Msg & "'", TellColor)
                Call PlayerMsg(index, "Le dices a " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
            Else
                Call PlayerMsg(index, "El jugador esta desconectado.", White)
            End If
        Else
            Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " begins to mumble to himself, what a wierdo...", PLAYER_LOG)
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " begins to mumble to himself, what a wierdo...", Green)
        End If
        TextAdd frmServer.txtText(4), "Para " & GetPlayerName(MsgTo) & " De " & GetPlayerName(index) & ": " & Msg, True
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
            If GetTickCount > Player(index).AttackTimer + 100 Then
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
        
    ' :::::::::::::::::::::
    ' :: Use item packet ::
    ' :::::::::::::::::::::
    Case "useitem"
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
            
            Dim n1 As Long, n2 As Long, n3 As Long, n4 As Long, n5 As Long, n6 As Long, n7 As Long
            
            n1 = Item(GetPlayerInvItemNum(index, InvNum)).StrReq
            n2 = Item(GetPlayerInvItemNum(index, InvNum)).DefReq
            n3 = Item(GetPlayerInvItemNum(index, InvNum)).SpeedReq
            n4 = Item(GetPlayerInvItemNum(index, InvNum)).ClassReq
            n5 = Item(GetPlayerInvItemNum(index, InvNum)).AccessReq
            n6 = Item(GetPlayerInvItemNum(index, InvNum)).MagiReq
            n7 = Item(GetPlayerInvItemNum(index, InvNum)).LvlReq
            
            ' Find out what kind of item it is
            Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum <> GetPlayerArmorSlot(index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(index) <> n4 Then
                                Call PlayerMsg(index, "Necesitas ser un/a " & GetClassName(n4) & " para usar este item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Tu acceso necesita ser mas alto que " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Tu Fuerza es muy baja para equiparte este item! Fuerza Requerida (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Tu Defensa es muy baja para equiparte este item! Defensa Requerida (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Tu Velocidad es muy baja para equiparte este item! Velocidad Requerida (" & n3 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerMAGI(index)) < n6 Then
                            Call PlayerMsg(index, "Tu Magia es muy baja para equiparte este item! Magia Requerida (" & n6 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLevel(index)) < n7 Then
                            Call PlayerMsg(index, "Tu Nivel es muy bajo para equiparte este item! Nivel Requerido (" & n7 & ")", BrightRed)
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
                                Call PlayerMsg(index, "Necesitas ser un/a " & GetClassName(n4) & " para usar este item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Tu acceso necesita ser mas alto que " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Tu Fuerza es muy baja para equiparte este item! Fuerza Requerida (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Tu Defensa es muy baja para equiparte este item! Defensa Requerida (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Tu Velocidad es muy baja para equiparte este item! Velocidad Requerida (" & n3 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerMAGI(index)) < n6 Then
                            Call PlayerMsg(index, "Tu Magia es muy baja para equiparte este item! Magia Requerida (" & n6 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLevel(index)) < n7 Then
                            Call PlayerMsg(index, "Tu Nivel es muy bajo para equiparte este item! Nivel Requerido (" & n7 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerWeaponSlot(index, InvNum)
                    Else
                        Call SetPlayerWeaponSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
                        
                    Case ITEM_TYPE_TWO_HAND
                        If InvNum <> GetPlayerWeaponSlot(index) Then
                           If n4 > -1 Then
                            If GetPlayerClass(index) <> n4 Then
                                Call PlayerMsg(index, "Necesitas ser un/a " & GetClassName(n4) & " para usar este item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Tu acceso necesita ser mas alto que " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Tu Fuerza es muy baja para equiparte este item! Fuerza Requerida (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Tu Defensa es muy baja para equiparte este item! Defensa Requerida (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Tu Velocidad es muy baja para equiparte este item! Velocidad Requerida (" & n3 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerMAGI(index)) < n6 Then
                            Call PlayerMsg(index, "Tu Magia es muy baja para equiparte este item! Magia Requerida (" & n6 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLevel(index)) < n7 Then
                            Call PlayerMsg(index, "Tu Nivel es muy bajo para equiparte este item! Nivel Requerido (" & n7 & ")", BrightRed)
                            Exit Sub
                        End If
                         
                            If GetPlayerShieldSlot(index) <> 0 Then
                                Call SetPlayerShieldSlot(index, 0)
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
                                Call PlayerMsg(index, "Necesitas ser un/a " & GetClassName(n4) & " para usar este item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Tu acceso necesita ser mas alto que " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Tu Fuerza es muy baja para equiparte este item! Fuerza Requerida (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Tu Defensa es muy baja para equiparte este item! Defensa Requerida (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Tu Velocidad es muy baja para equiparte este item! Velocidad Requerida (" & n3 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerMAGI(index)) < n6 Then
                            Call PlayerMsg(index, "Tu Magia es muy baja para equiparte este item! Magia Requerida (" & n6 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLevel(index)) < n7 Then
                            Call PlayerMsg(index, "Tu Nivel es muy bajo para equiparte este item! Nivel Requerido (" & n7 & ")", BrightRed)
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
                                Call PlayerMsg(index, "Necesitas ser un/a " & GetClassName(n4) & " para usar este item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Tu acceso necesita ser mas alto que " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Tu Fuerza es muy baja para equiparte este item! Fuerza Requerida (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Tu Defensa es muy baja para equiparte este item! Defensa Requerida (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Tu Velocidad es muy baja para equiparte este item! Velocidad Requerida (" & n3 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerMAGI(index)) < n6 Then
                            Call PlayerMsg(index, "Tu Magia es muy baja para equiparte este item! Magia Requerida (" & n6 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLevel(index)) < n7 Then
                            Call PlayerMsg(index, "Tu Nivel es muy bajo para equiparte este item! Nivel Requerido (" & n7 & ")", BrightRed)
                            Exit Sub
                        End If
                        If GetPlayerWeaponSlot(index) <> 0 Then
                                If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Type = ITEM_TYPE_TWO_HAND Then
                                    Call SetPlayerWeaponSlot(index, 0)
                                End If
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
                                Call PlayerMsg(index, "Necesitas ser un/a " & GetClassName(n4) & " para usar este item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Tu acceso necesita ser mas alto que " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Tu Fuerza es muy baja para equiparte este item! Fuerza Requerida (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Tu Defensa es muy baja para equiparte este item! Defensa Requerida (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Tu Velocidad es muy baja para equiparte este item! Velocidad Requerida (" & n3 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerMAGI(index)) < n6 Then
                            Call PlayerMsg(index, "Tu Magia es muy baja para equiparte este item! Magia Requerida (" & n6 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLevel(index)) < n7 Then
                            Call PlayerMsg(index, "Tu Nivel es muy bajo para equiparte este item! Nivel Requerido (" & n7 & ")", BrightRed)
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
                                Call PlayerMsg(index, "Necesitas ser un/a " & GetClassName(n4) & " para usar este item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Tu acceso necesita ser mas alto que " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Tu Fuerza es muy baja para equiparte este item! Fuerza Requerida (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Tu Defensa es muy baja para equiparte este item! Defensa Requerida (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Tu Velocidad es muy baja para equiparte este item! Velocidad Requerida (" & n3 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerMAGI(index)) < n6 Then
                            Call PlayerMsg(index, "Tu Magia es muy baja para equiparte este item! Magia Requerida (" & n6 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLevel(index)) < n7 Then
                            Call PlayerMsg(index, "Tu Nivel es muy bajo para equiparte este item! Nivel Requerido (" & n7 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerBootsSlot(index, InvNum)
                    Else
                        Call SetPlayerBootsSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
                
                Case ITEM_TYPE_BELT
                    If InvNum <> GetPlayerBeltSlot(index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(index) <> n4 Then
                                Call PlayerMsg(index, "Necesitas ser un/a " & GetClassName(n4) & " para usar este item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Tu acceso necesita ser mas alto que " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Tu Fuerza es muy baja para equiparte este item! Fuerza Requerida (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Tu Defensa es muy baja para equiparte este item! Defensa Requerida (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Tu Velocidad es muy baja para equiparte este item! Velocidad Requerida (" & n3 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerMAGI(index)) < n6 Then
                            Call PlayerMsg(index, "Tu Magia es muy baja para equiparte este item! Magia Requerida (" & n6 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLevel(index)) < n7 Then
                            Call PlayerMsg(index, "Tu Nivel es muy bajo para equiparte este item! Nivel Requerido (" & n7 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerBeltSlot(index, InvNum)
                    Else
                        Call SetPlayerBeltSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
                
                Case ITEM_TYPE_LEGS
                    If InvNum <> GetPlayerLegsSlot(index) Then
                        If n4 > -1 Then
                            If GetPlayerClass(index) <> n4 Then
                                Call PlayerMsg(index, "Necesitas ser un/a " & GetClassName(n4) & " para usar este item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Tu acceso necesita ser mas alto que " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Tu Fuerza es muy baja para equiparte este item! Fuerza Requerida (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Tu Defensa es muy baja para equiparte este item! Defensa Requerida (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Tu Velocidad es muy baja para equiparte este item! Velocidad Requerida (" & n3 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerMAGI(index)) < n6 Then
                            Call PlayerMsg(index, "Tu Magia es muy baja para equiparte este item! Magia Requerida (" & n6 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLevel(index)) < n7 Then
                            Call PlayerMsg(index, "Tu Nivel es muy bajo para equiparte este item! Nivel Requerido (" & n7 & ")", BrightRed)
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
                                Call PlayerMsg(index, "Necesitas ser un/a " & GetClassName(n4) & " para usar este item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Tu acceso necesita ser mas alto que " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Tu Fuerza es muy baja para equiparte este item! Fuerza Requerida (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Tu Defensa es muy baja para equiparte este item! Defensa Requerida (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Tu Velocidad es muy baja para equiparte este item! Velocidad Requerida (" & n3 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerMAGI(index)) < n6 Then
                            Call PlayerMsg(index, "Tu Magia es muy baja para equiparte este item! Magia Requerida (" & n6 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLevel(index)) < n7 Then
                            Call PlayerMsg(index, "Tu Nivel es muy bajo para equiparte este item! Nivel Requerido (" & n7 & ")", BrightRed)
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
                                Call PlayerMsg(index, "Necesitas ser un/a " & GetClassName(n4) & " para usar este item!", BrightRed)
                                Exit Sub
                            End If
                        End If
                        If GetPlayerAccess(index) < n5 Then
                            Call PlayerMsg(index, "Tu acceso necesita ser mas alto que " & n5 & "!", BrightRed)
                            Exit Sub
                        End If
                        If Int(GetPlayerSTR(index)) < n1 Then
                            Call PlayerMsg(index, "Tu Fuerza es muy baja para equiparte este item! Fuerza Requerida (" & n1 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerDEF(index)) < n2 Then
                            Call PlayerMsg(index, "Tu Defensa es muy baja para equiparte este item! Defensa Requerida (" & n2 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                            Call PlayerMsg(index, "Tu Velocidad es muy baja para equiparte este item! Velocidad Requerida (" & n3 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerMAGI(index)) < n6 Then
                            Call PlayerMsg(index, "Tu Magia es muy baja para equiparte este item! Magia Requerida (" & n6 & ")", BrightRed)
                            Exit Sub
                        ElseIf Int(GetPlayerLevel(index)) < n7 Then
                            Call PlayerMsg(index, "Tu Nivel es muy bajo para equiparte este item! Nivel Requerido (" & n7 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerAmuletSlot(index, InvNum)
                    Else
                        Call SetPlayerAmuletSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
                
                Case ITEM_TYPE_POTIONADDHP
                    Call SetPlayerHP(index, GetPlayerHP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendHP(index)
                
                Case ITEM_TYPE_POTIONADDMP
                    Call SetPlayerMP(index, GetPlayerMP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendMP(index)
        
                Case ITEM_TYPE_POTIONADDSP
                    Call SetPlayerSP(index, GetPlayerSP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendSP(index)

                Case ITEM_TYPE_POTIONSUBHP
                    Call SetPlayerHP(index, GetPlayerHP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendHP(index)
                
                Case ITEM_TYPE_POTIONSUBMP
                    Call SetPlayerMP(index, GetPlayerMP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
                    Call SendMP(index)
        
                Case ITEM_TYPE_POTIONSUBSP
                    Call SetPlayerSP(index, GetPlayerSP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
                    Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
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
                                Call MapMsg(GetPlayerMap(index), "Una puerta fue abierta!", White)
                            Else
                                Call MapMsg(GetPlayerMap(index), Trim(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), White)
                            End If
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
                            
                            ' Check if we are supposed to take away the item
                            If Map(GetPlayerMap(index)).Tile(x, y).Data2 = 1 Then
                                Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                Call PlayerMsg(index, "La llave se disuelve.", Yellow)
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
                                Call PlayerMsg(index, "Este hechizo solo puede ser usado por Administadores!", BrightRed)
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
                                        Call PlayerMsg(index, "Estudias el hechizo con cuidado...", Yellow)
                                        Call PlayerMsg(index, "Aprendiste un nuevo hechizo!", White)
                                    Else
                                        Call PlayerMsg(index, "Ya aprendiste este hechizo!", BrightRed)
                                    End If
                                Else
                                    Call PlayerMsg(index, "Ya aprendiste todos los hechizos que puedes aprender!", BrightRed)
                                End If
                            Else
                                Call PlayerMsg(index, "Debes tener nivel " & i & " para aprender este hechizo.", White)
                            End If
                        Else
                            Call PlayerMsg(index, "Este hechizo solo puede ser aprendido por un/a " & GetClassName(Spell(n).ClassReq - 1) & ".", White)
                        End If
                    Else
                        Call PlayerMsg(index, "No hay ninguno hechizo aqui, por favor avisale a un Admin!", White)
                    End If
            End Select
            
            Call SendStats(index)
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
        End If
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
                    If CanPlayerDodgeHit(i) Then
                        Call BattleMsg(index, GetPlayerName(i) & " esquivo tu ataque!", Yellow, 0)
                        Call BattleMsg(i, "Tu esquivaste el ataque de " & GetPlayerName(index) & "!", Yellow, 1)
                    End If
                    If Not CanPlayerBlockHit(i) Then
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(index) Then
                            c = GetPlayerDamage(index) - GetPlayerProtection(i)
                            c1 = GetPlayerFireDamage(index) - GetPlayerFireProtection(i)
                            c2 = GetPlayerWaterDamage(index) - GetPlayerWaterProtection(i)
                            c3 = GetPlayerEarthDamage(index) - GetPlayerEarthProtection(i)
                            c4 = GetPlayerAirDamage(index) - GetPlayerAirProtection(i)
                            c5 = GetPlayerHeatDamage(index) - GetPlayerHeatProtection(i)
                            c6 = GetPlayerColdDamage(index) - GetPlayerColdProtection(i)
                            c7 = GetPlayerLightDamage(index) - GetPlayerLightProtection(i)
                            c8 = GetPlayerDarkDamage(index) - GetPlayerDarkProtection(i)
                        If c < 0 Then
                            c = 0
                        End If
                        If c1 < 0 Then
                            c1 = 0
                        End If
                        If c2 < 0 Then
                            c2 = 0
                        End If
                        If c3 < 0 Then
                            c3 = 0
                        End If
                        If c4 < 0 Then
                            c4 = 0
                        End If
                        If c5 < 0 Then
                            c5 = 0
                        End If
                        If c6 < 0 Then
                            c6 = 0
                        End If
                        If c7 < 0 Then
                            c7 = 0
                        End If
                        If c8 < 0 Then
                            c8 = 0
                        End If
                            Damage = c + c1
                            Damage = Damage + c2
                            Damage = Damage + c3
                            Damage = Damage + c4
                            Damage = Damage + c5
                            Damage = Damage + c6
                            Damage = Damage + c7
                            Damage = Damage + c8
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                        Else
                            m = GetPlayerDamage(index)
                            d = m + Int(Rnd * Int(m / 2)) + 1 - GetPlayerProtection(i)
                            m1 = GetPlayerFireDamage(index)
                            d1 = m1 + Int(Rnd * Int(m1 / 2)) + 1 - GetPlayerFireProtection(i)
                            m2 = GetPlayerWaterDamage(index)
                            d2 = m2 + Int(Rnd * Int(m2 / 2)) + 1 - GetPlayerWaterProtection(i)
                            m3 = GetPlayerEarthDamage(index)
                            d3 = m3 + Int(Rnd * Int(m3 / 2)) + 1 - GetPlayerEarthProtection(i)
                            m4 = GetPlayerAirDamage(index)
                            d4 = m4 + Int(Rnd * Int(m4 / 2)) + 1 - GetPlayerAirProtection(i)
                            m5 = GetPlayerHeatDamage(index)
                            d5 = m5 + Int(Rnd * Int(m5 / 2)) + 1 - GetPlayerHeatProtection(i)
                            m6 = GetPlayerColdDamage(index)
                            d6 = m6 + Int(Rnd * Int(m6 / 2)) + 1 - GetPlayerColdProtection(i)
                            m7 = GetPlayerLightDamage(index)
                            d7 = m7 + Int(Rnd * Int(m7 / 2)) + 1 - GetPlayerLightProtection(i)
                            m8 = GetPlayerDarkDamage(index)
                            d8 = m8 + Int(Rnd * Int(m8 / 2)) + 1 - GetPlayerDarkProtection(i)
                        If d < 0 Then
                            d = 0
                        End If
                        If d1 < 0 Then
                            d1 = 0
                        End If
                        If d2 < 0 Then
                            d2 = 0
                        End If
                        If d3 < 0 Then
                            d3 = 0
                        End If
                        If d4 < 0 Then
                            d4 = 0
                        End If
                        If d5 < 0 Then
                            d5 = 0
                        End If
                        If d6 < 0 Then
                            d6 = 0
                        End If
                        If d7 < 0 Then
                            d7 = 0
                        End If
                        If d8 < 0 Then
                            d8 = 0
                        End If
                            Damage = d + d1
                            Damage = Damage + d2
                            Damage = Damage + d3
                            Damage = Damage + d4
                            Damage = Damage + d5
                            Damage = Damage + d6
                            Damage = Damage + d7
                            Damage = Damage + d8
                            Call BattleMsg(index, "Daño Critico!", Yellow, 0)
                            Call BattleMsg(i, GetPlayerName(index) & " golpea con una gran energia!", Yellow, 1)
                            
                            'Call PlayerMsg(index, "Sientes una gran energia mientras golpeas!", BrightCyan)
                            'Call PlayerMsg(I, GetPlayerName(index) & " golpea con una gran energia!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(index, i, Damage)
                        Else
                            Call PlayerMsg(index, "Tu ataque no hace nada.", BrightRed)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call BattleMsg(index, GetPlayerName(i) & " bloqueo tu golpe!", Yellow, 0)
                        Call BattleMsg(i, "Tu bloqueaste el golpe de " & GetPlayerName(index) & "!", Yellow, 1)
                        
                        'Call PlayerMsg(index, GetPlayerName(I) & " ha bloqueado tu golpe con su " & Trim(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & "!", BrightCyan)
                        'Call PlayerMsg(I, "Tu " & Trim(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " bloqueo el golpe de " & GetPlayerName(index) & "!", BrightCyan)
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
                    c = GetPlayerDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).DEF)
                    c1 = GetPlayerFireDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).FireDEF)
                    c2 = GetPlayerWaterDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).WaterDEF)
                    c3 = GetPlayerEarthDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).EarthDEF)
                    c4 = GetPlayerAirDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).AirDEF)
                    c5 = GetPlayerHeatDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).HeatDEF)
                    c6 = GetPlayerColdDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).ColdDEF)
                    c7 = GetPlayerLightDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).LightDEF)
                    c8 = GetPlayerDarkDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).DarkDEF)
                    If c < 0 Then
                    c = 0
                    End If
                    If c1 < 0 Then
                    c1 = 0
                    End If
                    If c2 < 0 Then
                    c2 = 0
                    End If
                    If c3 < 0 Then
                    c3 = 0
                    End If
                    If c4 < 0 Then
                    c4 = 0
                    End If
                    If c5 < 0 Then
                    c5 = 0
                    End If
                    If c6 < 0 Then
                    c6 = 0
                    End If
                    If c7 < 0 Then
                    c7 = 0
                    End If
                    If c8 < 0 Then
                    c8 = 0
                    End If
                    Damage = c + c1
                    Damage = Damage + c2
                    Damage = Damage + c3
                    Damage = Damage + c4
                    Damage = Damage + c5
                    Damage = Damage + c6
                    Damage = Damage + c7
                    Damage = Damage + c8
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                Else
                    m = GetPlayerDamage(index)
                    d = m + Int(Rnd * Int(m / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).DEF / 2)
                    m1 = GetPlayerFireDamage(index)
                    d1 = m1 + Int(Rnd * Int(m1 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).FireDEF / 2)
                    m2 = GetPlayerWaterDamage(index)
                    d2 = m2 + Int(Rnd * Int(m2 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).WaterDEF / 2)
                    m3 = GetPlayerEarthDamage(index)
                    d3 = m3 + Int(Rnd * Int(m3 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).EarthDEF / 2)
                    m4 = GetPlayerAirDamage(index)
                    d4 = m4 + Int(Rnd * Int(m4 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).AirDEF / 2)
                    m5 = GetPlayerHeatDamage(index)
                    d5 = m5 + Int(Rnd * Int(m5 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).HeatDEF / 2)
                    m6 = GetPlayerColdDamage(index)
                    d6 = m6 + Int(Rnd * Int(m6 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).ColdDEF / 2)
                    m7 = GetPlayerLightDamage(index)
                    d7 = m7 + Int(Rnd * Int(m7 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).LightDEF / 2)
                    m8 = GetPlayerDarkDamage(index)
                    d8 = m8 + Int(Rnd * Int(m8 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).DarkDEF / 2)
                    If d < 0 Then
                    d = 0
                    End If
                    If d1 < 0 Then
                    d1 = 0
                    End If
                    If d2 < 0 Then
                    d2 = 0
                    End If
                    If d3 < 0 Then
                    d3 = 0
                    End If
                    If d4 < 0 Then
                    d4 = 0
                    End If
                    If d5 < 0 Then
                    d5 = 0
                    End If
                    If d6 < 0 Then
                    d6 = 0
                    End If
                    If d7 < 0 Then
                    d7 = 0
                    End If
                    If d8 < 0 Then
                    d8 = 0
                    End If
                    Damage = d + d1
                    Damage = Damage + d2
                    Damage = Damage + d3
                    Damage = Damage + d4
                    Damage = Damage + d5
                    Damage = Damage + d6
                    Damage = Damage + d7
                    Damage = Damage + d8
                    Call BattleMsg(index, "Daño Critico!", Yellow, 0)
                    
                    'Call PlayerMsg(index, "Sientes una gran energia mientras golpeas!", BrightCyan)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(index, i, Damage)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                Else
                    Call BattleMsg(index, "Tu ataque no hace nada.", BrightRed, 0)
                    
                    'Call PlayerMsg(index, "Tu ataque no hace nada.", BrightRed)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                End If
                Exit Sub
            End If
        Next i
        
        Call AttackAttributeNpcs(index)
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
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "UsingStatPoints " & index & "," & PointType
            Else
                Select Case PointType
                    Case 0
                        Call SetPlayerSTR(index, GetPlayerSTR(index) + 1)
                        Call BattleMsg(index, "Has ganado mas fuerza!", 15, 0)
                    Case 1
                        Call SetPlayerDEF(index, GetPlayerDEF(index) + 1)
                        Call BattleMsg(index, "Has ganado mas defensa!", 15, 0)
                    Case 2
                        Call SetPlayerMAGI(index, GetPlayerMAGI(index) + 1)
                        Call BattleMsg(index, "Has ganado mas habilidades magicas!", 15, 0)
                    Case 3
                        Call SetPlayerSPEED(index, GetPlayerSPEED(index) + 1)
                        Call BattleMsg(index, "Has ganado mas velocidad!", 15, 0)
                End Select
                Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)
            End If
        Else
            Call BattleMsg(index, "No tienes puntos de habilidad para entrenar!", BrightRed, 0)
        End If
        
        Call SendHP(index)
        Call SendMP(index)
        Call SendSP(index)
        Call SendStats(index)
        
        Call SendDataTo(index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(index) & SEP_CHAR & END_CHAR)
        Exit Sub
        
    ' ::::::::::::::::::::::::::::::::
    ' :: Player info request packet ::
    ' ::::::::::::::::::::::::::::::::
    Case "playerinforequest"
        Name = Parse(1)
        
        i = FindPlayer(Name)
        If i > 0 Then
            Call PlayerMsg(index, "Cuenta: " & Trim(Player(i).Login) & ", Nombre: " & GetPlayerName(i), Yellow)
            If GetPlayerAccess(index) > ADMIN_MONITER Then
                Call PlayerMsg(index, "-=- Estadisticas para " & GetPlayerName(i) & " -=-", Yellow)
                Call PlayerMsg(index, "Nivel: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), Yellow)
                Call PlayerMsg(index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), Yellow)
                Call PlayerMsg(index, "STR: " & GetPlayerSTR(i) & "  DEF: " & GetPlayerDEF(i) & "  MAGI: " & GetPlayerMAGI(i) & "  SPEED: " & GetPlayerSPEED(i), Yellow)
                n = Int(GetPlayerSTR(i) / 10) + Int(GetPlayerLevel(i) / 10)
                i = Int(GetPlayerDEF(i) / 10) + Int(GetPlayerLevel(i) / 10)
                o = Int(GetPlayerSPEED(i) / 10) + Int(GetPlayerLevel(i) / 10)
                l = Int(GetPlayerMAGI(i) / 10) + Int(GetPlayerLevel(i) / 10)
                If n > 100 Then n = 100
                If i > 100 Then i = 100
                If o > 100 Then o = 100
                If l > 100 Then l = 100
                Call PlayerMsg(index, "Chance de daño critico: " & n & "%, Chance de bloqueo: " & i & "%", Yellow)
                Call PlayerMsg(index, "Chance de esquivar: " & o & "%, Chance de enfoque: " & l & "%", Yellow)
            End If
        Else
            Call PlayerMsg(index, "El jugador esta desconectado.", White)
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
        n = Val(Parse(1))
        
        Call SetPlayerSprite(index, n)
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
        n = Val(Parse(2))
                
        Call SetPlayerSprite(i, n)
        Call SendPlayerData(i)
        Exit Sub
                
    ' ::::::::::::::::::::::::::
    ' :: Stats request packet ::
    ' ::::::::::::::::::::::::::
    Case "getstats"
        Call PlayerMsg(index, "-=- Estadisticas para " & GetPlayerName(index) & " -=-", Yellow)
        Call PlayerMsg(index, "Nivel: " & GetPlayerLevel(index) & "  Exp: " & GetPlayerExp(index) & "/" & GetPlayerNextLevel(index), Yellow)
        Call PlayerMsg(index, "HP: " & GetPlayerHP(index) & "/" & GetPlayerMaxHP(index) & "  MP: " & GetPlayerMP(index) & "/" & GetPlayerMaxMP(index) & "  SP: " & GetPlayerSP(index) & "/" & GetPlayerMaxSP(index), Yellow)
        Call PlayerMsg(index, "Fuer: " & GetPlayerSTR(index) & "  Def: " & GetPlayerDEF(index) & "  Mag: " & GetPlayerMAGI(index) & "  Vel: " & GetPlayerSPEED(index), Yellow)
        n = Int(GetPlayerSTR(index) / 10) + Int(GetPlayerLevel(index) / 10)
        i = Int(GetPlayerDEF(index) / 10) + Int(GetPlayerLevel(index) / 10)
        o = Int(GetPlayerSPEED(index) / 10) + Int(GetPlayerLevel(index) / 10)
        l = Int(GetPlayerMAGI(index) / 10) + Int(GetPlayerLevel(index) / 10)
        If n > 100 Then n = 100
        If i > 100 Then i = 100
        If o > 100 Then o = 100
        If l > 100 Then l = 100
        Call PlayerMsg(index, "Chance de daño critico: " & n & "%, Chance de bloqueo: " & i & "%", Yellow)
        Call PlayerMsg(index, "Chance de esquivar: " & o & "%, Chance de enfoque: " & l & "%", Yellow)
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
            Map(MapNum).Tile(x, y).Light = Val(Parse(n + 16))
            Map(MapNum).Tile(x, y).GroundSet = Val(Parse(n + 17))
            Map(MapNum).Tile(x, y).MaskSet = Val(Parse(n + 18))
            Map(MapNum).Tile(x, y).AnimSet = Val(Parse(n + 19))
            Map(MapNum).Tile(x, y).Mask2Set = Val(Parse(n + 20))
            Map(MapNum).Tile(x, y).M2AnimSet = Val(Parse(n + 21))
            Map(MapNum).Tile(x, y).FringeSet = Val(Parse(n + 22))
            Map(MapNum).Tile(x, y).FAnimSet = Val(Parse(n + 23))
            Map(MapNum).Tile(x, y).Fringe2Set = Val(Parse(n + 24))
            Map(MapNum).Tile(x, y).F2AnimSet = Val(Parse(n + 25))

            n = n + 26
            Next x
        Next y
       
        For x = 1 To MAX_MAP_NPCS
            Map(MapNum).Npc(x) = Val(Parse(n))
            n = n + 1
            Call ClearMapNpc(x, MapNum)
        Next x
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                    For i = 1 To MAX_ATTRIBUTE_NPCS
                        If i <= Map(MapNum).Tile(x, y).Data2 Then
                            Call ClearMapAttributeNpc(i, x, y, GetPlayerMap(index))
                        End If
                    Next i
                End If
            Next x
        Next y
        
        ' Clear out it all
        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
            Call ClearMapItem(i, GetPlayerMap(index))
        Next i
        
        ' Save the map
        Call SaveMap(MapNum)
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(index))
        
        ' Respawn NPCS
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, GetPlayerMap(index))
        Next i
        
        ' Respawn NPCS
        Call SpawnMapAttributeNpcs(GetPlayerMap(index))

        ' Refresh map for everyone online
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & SEP_CHAR & END_CHAR)
                'Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
            End If
        Next i
    
        Exit Sub

    ' ::::::::::::::::::::::::::::
    ' :: Need map yes/no packet ::
    ' ::::::::::::::::::::::::::::
    Case "needmap"
        ' Get yes/no value
        s = LCase(Parse(1))
                
        If s = "yes" Then
            Call SendMap(index, GetPlayerMap(index))
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendMapAttributeNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            Player(index).GettingMap = NO
            Call SendDataTo(index, "MAPDONE" & SEP_CHAR & END_CHAR)
        Else
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendMapAttributeNpcsTo(index, GetPlayerMap(index))
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
        If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
            If Amount <= 0 Then
                Call PlayerMsg(index, "Debes soltar mas que 0!", BrightRed)
                Exit Sub
            End If
            
            If Amount > GetPlayerInvItemValue(index, InvNum) Then
                Call PlayerMsg(index, "No tienes esa cantidad!", BrightRed)
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
        
        Call PlayerMapDropItem(index, InvNum, Amount)
        Call SendStats(index)
        Call SendHP(index)
        Call SendMP(index)
        Call SendSP(index)
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
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
            Call ClearMapItem(i, GetPlayerMap(index))
        Next i
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(index))
        
        ' Respawn NPCS
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, GetPlayerMap(index))
        Next i
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                    For i = 1 To MAX_ATTRIBUTE_NPCS
                        Call ClearMapAttributeNpc(i, x, y, GetPlayerMap(index))
                    Next i
                End If
            Next x
        Next y
        
        Call SpawnMapAttributeNpcs(GetPlayerMap(index))
        
        Call PlayerMsg(index, "Mapa reiniciado.", Blue)
        Call AddLog(GetPlayerName(index) & " ha reiniciado el mapa #" & GetPlayerMap(index), ADMIN_LOG)
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
        n = FindPlayer(Parse(1))
        
        If n <> index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(index) Then
                    Call GlobalMsg(GetPlayerName(n) & " ha sido hechado de " & GAME_NAME & " por " & GetPlayerName(index) & "!", White)
                    Call AddLog(GetPlayerName(index) & " ha hechado a " & GetPlayerName(n) & ".", ADMIN_LOG)
                    Call AlertMsg(n, "Fuiste hechado por " & GetPlayerName(index) & "!")
                Else
                    Call PlayerMsg(index, "Ese es un acceso mas alto que el tuyo!", White)
                End If
            Else
                Call PlayerMsg(index, "El jugador esta desconectado.", White)
            End If
        Else
            Call PlayerMsg(index, "No puedes hecharte a tu mismo!", White)
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
        
        n = 1
        f = FreeFile
        Open App.Path & "\banlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            Input #f, Name
            
            Call PlayerMsg(index, n & ": Banned IP " & s & " por " & Name, White)
            n = n + 1
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
        Call PlayerMsg(index, "Lista de ban destruida.", White)
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
        n = FindPlayer(Parse(1))
        
        If n <> index Then
            If n > 0 Then
                If GetPlayerAccess(n) <= GetPlayerAccess(index) Then
                    Call BanIndex(n, index)
                Else
                    Call PlayerMsg(index, "Ese es un acceso mas alto que el tuyo!", White)
                End If
            Else
                Call PlayerMsg(index, "El jugador esta desconectado.", White)
            End If
        Else
            Call PlayerMsg(index, "No te puedes banear a ti mismo!", White)
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
    Case "saveitem"
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
        Item(n).MagiReq = Val(Parse(11))
        Item(n).LvlReq = Val(Parse(12))
        Item(n).ClassReq = Val(Parse(13))
        Item(n).AccessReq = Val(Parse(14))
        
        Item(n).FireSTR = Val(Parse(15))
        Item(n).FireDEF = Val(Parse(16))
        Item(n).WaterSTR = Val(Parse(17))
        Item(n).WaterDEF = Val(Parse(18))
        Item(n).EarthSTR = Val(Parse(19))
        Item(n).EarthDEF = Val(Parse(20))
        Item(n).AirSTR = Val(Parse(21))
        Item(n).AirDEF = Val(Parse(22))
        Item(n).HeatSTR = Val(Parse(23))
        Item(n).HeatDEF = Val(Parse(24))
        Item(n).ColdSTR = Val(Parse(25))
        Item(n).ColdDEF = Val(Parse(26))
        Item(n).LightSTR = Val(Parse(27))
        Item(n).LightDEF = Val(Parse(28))
        Item(n).DarkSTR = Val(Parse(29))
        Item(n).DarkDEF = Val(Parse(30))
        Item(n).AddHP = Val(Parse(31))
        Item(n).AddMP = Val(Parse(32))
        Item(n).AddSP = Val(Parse(33))
        Item(n).AddStr = Val(Parse(34))
        Item(n).AddDef = Val(Parse(35))
        Item(n).AddMagi = Val(Parse(36))
        Item(n).AddSpeed = Val(Parse(37))
        Item(n).AddEXP = Val(Parse(38))
        Item(n).Desc = Parse(39)
        Item(n).AttackSpeed = Val(Parse(40))
        
        
        ' Save it
        Call SendUpdateItemToAll(n)
        Call SaveItem(n)
        Call AddLog(GetPlayerName(index) & " saved item #" & n & ".", ADMIN_LOG)
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
    Case "savenpc"
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
        Npc(n).FireSTR = Val(Parse(10))
        Npc(n).FireDEF = Val(Parse(11))
        Npc(n).WaterSTR = Val(Parse(12))
        Npc(n).WaterDEF = Val(Parse(13))
        Npc(n).EarthSTR = Val(Parse(14))
        Npc(n).EarthDEF = Val(Parse(15))
        Npc(n).AirSTR = Val(Parse(16))
        Npc(n).AirDEF = Val(Parse(17))
        Npc(n).HeatSTR = Val(Parse(18))
        Npc(n).HeatDEF = Val(Parse(19))
        Npc(n).ColdSTR = Val(Parse(20))
        Npc(n).ColdDEF = Val(Parse(21))
        Npc(n).LightSTR = Val(Parse(22))
        Npc(n).LightDEF = Val(Parse(23))
        Npc(n).DarkSTR = Val(Parse(24))
        Npc(n).DarkDEF = Val(Parse(25))
        Npc(n).Speed = Val(Parse(26))
        Npc(n).Magi = Val(Parse(27))
        Npc(n).Big = Val(Parse(28))
        Npc(n).MaxHp = Val(Parse(29))
        Npc(n).Exp = Val(Parse(30))
        Npc(n).SpawnTime = Val(Parse(31))
        
        
        z = 32
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
        
        n = 6
        For z = 1 To 6
            For i = 1 To MAX_TRADES
                Shop(ShopNum).TradeItem(z).Value(i).GiveItem = Val(Parse(n))
                Shop(ShopNum).TradeItem(z).Value(i).GiveValue = Val(Parse(n + 1))
                Shop(ShopNum).TradeItem(z).Value(i).GetItem = Val(Parse(n + 2))
                Shop(ShopNum).TradeItem(z).Value(i).GetValue = Val(Parse(n + 3))
                n = n + 4
            Next i
        Next z
        
        ' Save it
        Call SendUpdateShopToAll(ShopNum)
        Call SaveShop(ShopNum)
        Call AddLog(GetPlayerName(index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
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
    Case "savespell"
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
        Spell(n).FireSTR = Val(Parse(7))
        Spell(n).WaterSTR = Val(Parse(8))
        Spell(n).EarthSTR = Val(Parse(9))
        Spell(n).AirSTR = Val(Parse(10))
        Spell(n).HeatSTR = Val(Parse(11))
        Spell(n).ColdSTR = Val(Parse(12))
        Spell(n).LightSTR = Val(Parse(13))
        Spell(n).DarkSTR = Val(Parse(14))
        Spell(n).Data2 = Val(Parse(15))
        Spell(n).Data3 = Val(Parse(16))
        Spell(n).MPCost = Val(Parse(17))
        Spell(n).Sound = Val(Parse(18))
        Spell(n).Range = Val(Parse(19))
        Spell(n).SpellAnim = Val(Parse(20))
        Spell(n).SpellTime = Val(Parse(21))
        Spell(n).SpellDone = Val(Parse(22))
        Spell(n).AE = Val(Parse(23))
                
        ' Save it
        Call SendUpdateSpellToAll(n)
        Call SaveSpell(n)
        Call AddLog(GetPlayerName(index) & " saving spell #" & n & ".", ADMIN_LOG)
        Exit Sub
    
    ' :::::::::::::::::::::::
    ' :: Set access packet ::
    ' :::::::::::::::::::::::
    Case "setaccess"
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_CREATOR Then
            Call HackingAttempt(index, "Tratando de usar poderes que no posees")
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
                            Call GlobalMsg(GetPlayerName(n) & " fue bendecido con acceso de administrador.", BrightBlue)
                        End If
                    
                        Call SetPlayerAccess(n, i)
                        Call SendPlayerData(n)
                        Call AddLog(GetPlayerName(index) & " ha modificado el acceso de " & GetPlayerName(n) & ".", ADMIN_LOG)
                    Else
                        Call PlayerMsg(index, "El jugador esta desconectado.", White)
                    End If
                Else
                    Call PlayerMsg(index, "Tu nivel de acceso es mas bajo que el de " & GetPlayerName(n) & ".", Red)
                End If
            Else
                Call PlayerMsg(index, "No puedes cambiar tu acceso.", Red)
            End If
        Else
            Call PlayerMsg(index, "Nivel de acceso invalido.", Red)
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
        
        Call PutVar(App.Path & "\motd.ini", "MOTD", "Msg", Parse(1))
        Call GlobalMsg("MOTD cambiado a: " & Parse(1), BrightCyan)
        Call AddLog(GetPlayerName(index) & " cambio el MOTD a: " & Parse(1), ADMIN_LOG)
        Exit Sub
    
    Case "traderequest"
        ' Trade num
        n = Val(Parse(1))
        z = Val(Parse(2))
        
       ' Prevent hacking
         If (n < 1) Or (n > MAX_TRADES) Then
             Call HackingAttempt(index, "Trade Request Modification")
             Exit Sub
         End If
        
        ' Prevent hacking
        If (z <= 0) Or (z > MAX_TRADES) Then
            Call HackingAttempt(index, "Trade Request Modification")
            Exit Sub
        End If
     
        ' Index for shop
        i = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        
        ' Check if inv full
        If i <= 0 Then Exit Sub
        x = FindOpenInvSlot(index, Shop(i).TradeItem(n).Value(z).GetItem)
        If x = 0 Then
            Call PlayerMsg(index, "Comercio fallido, inventario lleno.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have the item
        If HasItem(index, Shop(i).TradeItem(n).Value(z).GiveItem) >= Shop(i).TradeItem(n).Value(z).GiveValue Then
            Call TakeItem(index, Shop(i).TradeItem(n).Value(z).GiveItem, Shop(i).TradeItem(n).Value(z).GiveValue)
            Call GiveItem(index, Shop(i).TradeItem(n).Value(z).GetItem, Shop(i).TradeItem(n).Value(z).GetValue)
            Call PlayerMsg(index, "El comercio fue realizado!", Yellow)
        Else
            Call PlayerMsg(index, "Comercio cancelado.", BrightRed)
        End If
        Exit Sub

    Case "fixitem"
        ' Inv num
        n = Val(Parse(1))
        
        ' Make sure its a equipable item
        If Item(GetPlayerInvItemNum(index, n)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(index, n)).Type > ITEM_TYPE_SHIELD Then
        If Item(GetPlayerInvItemNum(index, n)).Type < ITEM_TYPE_GLOVES Or Item(GetPlayerInvItemNum(index, n)).Type > ITEM_TYPE_AMULET Then
            Call PlayerMsg(index, "No puedes reparar este item.", BrightRed)
            Exit Sub
        End If
        End If
        
        ' Check if they have a full inventory
        If FindOpenInvSlot(index, GetPlayerInvItemNum(index, n)) <= 0 Then
            Call PlayerMsg(index, "No tienes espacio en el inventario!", BrightRed)
            Exit Sub
        End If
        
        ' Now check the rate of pay
        ItemNum = GetPlayerInvItemNum(index, n)
        i = Int(Item(GetPlayerInvItemNum(index, n)).Data2 / 2)
        If i <= 0 Then i = 1
        
        DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(index, n)
        GoldNeeded = Int(DurNeeded * i / 10000)
        If GoldNeeded <= 0 Then GoldNeeded = 1
        
        ' Check if they even need it repaired
        If DurNeeded <= 0 Then
            Call PlayerMsg(index, "Este item esta en perfectas condiciones!", White)
            Exit Sub
        End If
        
        ' Check if they have enough for at least one point
        If HasItem(index, 1) >= i Then
            ' Check if they have enough for a total restoration
            If HasItem(index, 1) >= GoldNeeded Then
                Call TakeItem(index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(index, n, Item(ItemNum).Data1)
                Call PlayerMsg(index, "El item fue reparado por " & GoldNeeded & "!", BrightBlue)
            Else
                ' They dont so restore as much as we can
                DurNeeded = (HasItem(index, 1) / i)
                GoldNeeded = Int(DurNeeded * i / 10000)
                If GoldNeeded <= 0 Then GoldNeeded = 1
                
                Call TakeItem(index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(index, n, GetPlayerInvItemDur(index, n) + DurNeeded)
                Call PlayerMsg(index, "El item fue reparado parcialmente por " & GoldNeeded & "!", BrightBlue)
            End If
        Else
            Call PlayerMsg(index, "Dinero insuficiente para reparar este item!", BrightRed)
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
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) And GetPlayerX(i) = x And GetPlayerY(i) = y Then
                
                ' Consider the player
                If GetPlayerLevel(i) >= GetPlayerLevel(index) + 10 Then
                    Call PlayerMsg(index, "No tiene oportunidad.", BrightRed)
                Else
                    If GetPlayerLevel(i) > GetPlayerLevel(index) Then
                        Call PlayerMsg(index, "Este jugador parece tener una ventaja sobre ti.", Yellow)
                    Else
                        If GetPlayerLevel(i) = GetPlayerLevel(index) Then
                            Call PlayerMsg(index, "Esta sera una batalla justa.", White)
                        Else
                            If GetPlayerLevel(index) >= GetPlayerLevel(i) + 10 Then
                                Call PlayerMsg(index, "Puedes masacarar a ese jugador.", BrightBlue)
                            Else
                                If GetPlayerLevel(index) > GetPlayerLevel(i) Then
                                    Call PlayerMsg(index, "Tienes una ventaja sobre ese jugador.", Yellow)
                                End If
                            End If
                        End If
                    End If
                End If
            
                ' Change target
                Player(index).Target = i
                Player(index).TargetType = TARGET_TYPE_PLAYER
                Call PlayerMsg(index, "Tu objetivo ahora es " & GetPlayerName(i) & ".", Yellow)
                Exit Sub
            End If
        Next i
        
        ' Check for an npc
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(index), i).num > 0 Then
                If MapNpc(GetPlayerMap(index), i).x = x And MapNpc(GetPlayerMap(index), i).y = y Then
                    ' Change target
                    Player(index).Target = i
                    Player(index).TargetType = TARGET_TYPE_NPC
                    Call PlayerMsg(index, "Tu objetivo ahora es un/a " & Trim(Npc(MapNpc(GetPlayerMap(index), i).num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next i

        BX = x
        BY = y
        For x = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY
                If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                    For i = 1 To MAX_ATTRIBUTE_NPCS
                        If MapAttributeNpc(GetPlayerMap(index), i, x, y).num > 0 Then
                            If MapAttributeNpc(GetPlayerMap(index), i, x, y).x = BX And MapAttributeNpc(GetPlayerMap(index), i, x, y).y = BY Then
                                ' Change target
                                Player(index).Target = i
                                Player(index).TargetType = TARGET_TYPE_ATTRIBUTE_NPC
                                Call PlayerMsg(index, "Tu objetivo ahora es un/a " & Trim(Npc(MapAttributeNpc(GetPlayerMap(index), i, x, y).num).Name) & ".", Yellow)
                                Exit Sub
                            End If
                        End If
                    Next i
                End If
            Next y
        Next x
        
        ' Check for an item
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(GetPlayerMap(index), i).num > 0 Then
                If MapItem(GetPlayerMap(index), i).x = x And MapItem(GetPlayerMap(index), i).y = y Then
                    Call PlayerMsg(index, "Ves un/a " & Trim(Item(MapItem(GetPlayerMap(index), i).num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next i
        Exit Sub
    
    Case "playerchat"
        n = FindPlayer(Parse(1))
        If n < 1 Then
            Call PlayerMsg(index, "El jugador esta desconectado.", White)
            Exit Sub
        End If
        If n = index Then
            Exit Sub
        End If
        If Player(index).InChat = 1 Then
            Call PlayerMsg(index, "Ya estas charlando con un jugador!", Pink)
            Exit Sub
        End If

        If Player(n).InChat = 1 Then
            Call PlayerMsg(index, "El jugador esta charlando con otro jugador!", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "El pedido de charla fue enviado a " & GetPlayerName(n) & ".", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " quiere charlar contigo.  Tipea /charla para aceptar, o /declinarch para declinar.", Pink)
    
        Player(n).ChatPlayer = index
        Player(index).ChatPlayer = n
        Exit Sub
    
    Case "achat"
        n = Player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "Nadie desea charlar contigo.", Pink)
            Exit Sub
        End If
        If Player(n).ChatPlayer <> index Then
            Call PlayerMsg(index, "Charla fallida.", Pink)
            Exit Sub
        End If
                        
        Call SendDataTo(index, "PPCHATTING" & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "PPCHATTING" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        Exit Sub
    
    Case "dchat"
        n = Player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "Nadie desea charlar contigo.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Pedido de charla declinado.", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " declino tu pedido.", Pink)
        
        Player(index).ChatPlayer = 0
        Player(index).InChat = 0
        Player(n).ChatPlayer = 0
        Player(n).InChat = 0
        Exit Sub

    Case "qchat"
        n = Player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "Nadie desea charlar contigo.", Pink)
            Exit Sub
        End If

        Call SendDataTo(index, "qchat" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "qchat" & SEP_CHAR & END_CHAR)
        
        Player(index).ChatPlayer = 0
        Player(index).InChat = 0
        Player(n).ChatPlayer = 0
        Player(n).InChat = 0
        Exit Sub
    
    Case "sendchat"
        n = Player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "Nadie desea charlar contigo.", Pink)
            Exit Sub
        End If

        Call SendDataTo(n, "sendchat" & SEP_CHAR & Parse(1) & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "pptrade"
        n = FindPlayer(Parse(1))
        
        ' Check if player is online
        If n < 1 Then
            Call PlayerMsg(index, "El jugador esta desconectado.", White)
            Exit Sub
        End If
        
        ' Prevent trading with self
        If n = index Then
            Exit Sub
        End If
                
        ' Check if the player is in another trade
        If Player(index).InTrade = 1 Then
            Call PlayerMsg(index, "Ya estas comerciando con alguien mas!", Pink)
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
                Call PlayerMsg(index, "El jugador ya esta comerciando!", Pink)
                Exit Sub
            End If
            
            Call PlayerMsg(index, "El pedido de comercio fue enciado a " & GetPlayerName(n) & ".", Pink)
            Call PlayerMsg(n, GetPlayerName(index) & " quiere comerciar conrigo.  Tipea /aceptar para aceptar, o /declinarco para declinar.", Pink)
        
            Player(n).TradePlayer = index
            Player(index).TradePlayer = n
        Else
            Call PlayerMsg(index, "Necesitas estar al lado del jugador para comerciar!", Pink)
            Call PlayerMsg(n, "El jugador necesita estar a tu lado para comerciar!", Pink)
        End If
        Exit Sub

    Case "atrade"
        n = Player(index).TradePlayer
        
        ' Check if anyone requested a trade
        If n < 1 Then
            Call PlayerMsg(index, "Nadie desea comerciar contigo.", Pink)
            Exit Sub
        End If
        
        ' Check if its the right player
        If Player(n).TradePlayer <> index Then
            Call PlayerMsg(index, "Comercio fallido.", Pink)
            Exit Sub
        End If
        
        ' Check where both players are
        CanTrade = False
        
        If GetPlayerX(index) = GetPlayerX(n) And GetPlayerY(index) + 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) = GetPlayerX(n) And GetPlayerY(index) - 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) + 1 = GetPlayerX(n) And GetPlayerY(index) = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) - 1 = GetPlayerX(n) And GetPlayerY(index) = GetPlayerY(n) Then CanTrade = True
            
        If CanTrade = True Then
            Call PlayerMsg(index, "Estas comerciando con " & GetPlayerName(n) & "!", Pink)
            Call PlayerMsg(n, GetPlayerName(index) & " acepto tu pedido de comercio!", Pink)
            
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
            Call PlayerMsg(index, "El jugador necesita estar a tu lado para comerciar!", Pink)
            Call PlayerMsg(n, "Necesitas estar al lado del jugador para comerciar!", Pink)
        End If
        Exit Sub

    Case "qtrade"
        n = Player(index).TradePlayer
        
        ' Check if anyone trade with player
        If n < 1 Then
            Call PlayerMsg(index, "Nadie desea comerciar contigo.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Comercio cancelado.", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " dejo de comerciar contigo!", Pink)

        Player(index).TradeOk = 0
        Player(n).TradeOk = 0
        Player(index).TradePlayer = 0
        Player(index).InTrade = 0
        Player(n).TradePlayer = 0
        Player(n).InTrade = 0
        Call SendDataTo(index, "qtrade" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "qtrade" & SEP_CHAR & END_CHAR)
        Exit Sub

    Case "dtrade"
        n = Player(index).TradePlayer
        
        ' Check if anyone trade with player
        If n < 1 Then
            Call PlayerMsg(index, "Nadie desea comerciar contigo.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Pedido de comercio declinado.", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " declino tu pedido.", Pink)
        
        Player(index).TradePlayer = 0
        Player(index).InTrade = 0
        Player(n).TradePlayer = 0
        Player(n).InTrade = 0
        Exit Sub

    Case "updatetradeinv"
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
    
    Case "swapitems"
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

                Call PlayerMsg(n, "Comercio completado!", BrightGreen)
                Call PlayerMsg(index, "Comercio completado!", BrightGreen)
                Call SendInventory(n)
                Call SendInventory(index)
            Else
                If Player(index).TradeItemMax2 < Player(index).TradeItemMax Then
                    Call PlayerMsg(index, "Tu inventario esta lleno!", BrightRed)
                    Call PlayerMsg(n, GetPlayerName(index) & " tiene su inventario lleno!", BrightRed)
                End If
                If Player(n).TradeItemMax2 < Player(n).TradeItemMax Then
                    Call PlayerMsg(n, "Tu inventario esta lleno!", BrightRed)
                    Call PlayerMsg(index, GetPlayerName(n) & " tiene su inventario lleno!", BrightRed)
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
        
    Case "party"
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
            Call PlayerMsg(index, "La party esta llena!", Pink)
            Exit Sub
            End If
        End If
       
        If n > 0 Then
            ' Check if its an admin
            If GetPlayerAccess(index) > ADMIN_MONITER Then
                Call PlayerMsg(index, "No puedes unirte en las partys, eres un admin!", BrightBlue)
                Exit Sub
            End If
       
            If GetPlayerAccess(n) > ADMIN_MONITER Then
                Call PlayerMsg(index, "Los admin no pueden unirse en las partys!", BrightBlue)
                Exit Sub
            End If
           
            ' Check to see if player is already in a party
            If Player(n).InParty = False Then
                Call PlayerMsg(index, GetPlayerName(n) & " fue invitado a tu party.", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " te invito a unirte a su party.  Tipea /unirse para unirse, o /dejar para declinar.", Pink)
               
                Player(n).InvitedBy = index
            Else
                Call PlayerMsg(index, "El jugador ya esta en una party!", Pink)
            End If
        Else
            Call PlayerMsg(index, "El jugador esta desconectado.", White)
        End If
        Exit Sub

    Case "joinparty"
        n = Player(index).InvitedBy
       
        If n > 0 Then
            ' Check to make sure they aren't the starter
                ' Check to make sure that each of there party players match
                    Call PlayerMsg(index, "Te has unido a la party de " & GetPlayerName(n) & "!", Pink)
                 
                    If Player(n).InParty = False Then ' Set the party leader up
                    Call SetPMember(n, n) 'Make them the first member and make them the leader
                    Player(n).InParty = True 'Set them to be 'InParty' status
                    Call SetPShare(n, True)
                    End If
                   
                    Player(index).InParty = True 'Player joined
                    Player(index).Party.Leader = n 'Set party leader
                    Call SetPMember(n, index) 'Add the member and update the party
                   
                    ' Make sure they are in right level range
                    If GetPlayerLevel(index) + 7 < GetPlayerLevel(n) Or GetPlayerLevel(index) - 7 > GetPlayerLevel(n) Then
                        Call PlayerMsg(index, "Hay mas de 7 niveles entre ustedes dos, la experiencia no se compartira.", Pink)
                        Call PlayerMsg(n, "Hay mas de 7 niveles entre ustedes dos, la experiencia no se compartira.", Pink)
                        Call SetPShare(index, False) 'Do not share experience with party
                    Else
                        Call SetPShare(index, True) 'Share experience with party
                    End If
                   
                    For i = 1 To MAX_PARTY_MEMBERS
                        If Player(index).Party.Member(i) > 0 And Player(index).Party.Member(i) <> index Then Call PlayerMsg(Player(index).Party.Member(i), GetPlayerName(index) & " se ha unido a tu party!", Pink)
                    Next i
                                       
        Else
            Call PlayerMsg(index, "No fuiste invitado a una party!", Pink)
        End If
        Exit Sub

    Case "leaveparty"
        n = Player(index).InvitedBy
       
        If n > 0 Or Player(index).Party.Leader = index Then
            If Player(index).InParty = True Then
                Call PlayerMsg(index, "Dejaste la party", Pink)
                For i = 1 To MAX_PARTY_MEMBERS
                    If Player(index).Party.Member(i) > 0 Then Call PlayerMsg(Player(index).Party.Member(i), GetPlayerName(index) & " ha dejado la party.", Pink)
                Next i
               
                Call RemovePMember(index) 'this handles removing them and updating the entire party
               
            Else
                Call PlayerMsg(index, "Pedido de party declinado.", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " declino tu pedido.", Pink)
               
                Player(index).InParty = False
                Player(index).InvitedBy = 0
               
            End If
        Else
            Call PlayerMsg(index, "No estas en una party!", Pink)
        End If
        Exit Sub
        
    Case "partychat"
        For i = 1 To MAX_PARTY_MEMBERS
            If Player(index).Party.Member(i) > 0 Then Call PlayerMsg(Player(index).Party.Member(i), Parse(1), Blue)
        Next i
        Exit Sub
    
    Case "spells"
        Call SendPlayerSpells(index)
        Exit Sub
    
    Case "cast"
        n = Val(Parse(1))
        Call CastSpell(index, n)
        Exit Sub

    Case "requestlocation"
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Map: " & GetPlayerMap(index) & ", X: " & GetPlayerX(index) & ", Y: " & GetPlayerY(index), Pink)
        Exit Sub
    
    Case "refresh"
        Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Exit Sub
    
    Case "buysprite"
        ' Check if player stepped on sprite changing tile
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
            Call PlayerMsg(index, "Necesitas estar en una capa de sprite para comprarlo!", BrightRed)
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
                        Call PlayerMsg(index, "Compraste un nuevo sprite!", BrightGreen)
                        Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                        Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(index)
                    End If
                Else
                    If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerGlovesSlot(index) <> i And GetPlayerBootsSlot(index) <> i And GetPlayerBeltSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerAmuletSlot(index) <> i Then
                        Call SetPlayerInvItemNum(index, i, 0)
                        Call PlayerMsg(index, "Compraste un nuevo sprite!", BrightGreen)
                        Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                        Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(index)
                    End If
                End If
                If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerGlovesSlot(index) <> i And GetPlayerBootsSlot(index) <> i And GetPlayerBeltSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerAmuletSlot(index) <> i Then
                    Exit Sub
                End If
            End If
        Next i
        
        Call PlayerMsg(index, "No tienes suficiente para comprar este sprite!", BrightRed)
        Exit Sub
        
    Case "checkcommands"
        s = Parse(1)
        If Scripting = 1 Then
            PutVar App.Path & "\Scripts\Command.ini", "TEMP", "Text" & index, Trim(s)
            MyScript.ExecuteStatement "Scripts\Main.txt", "Commands " & index
        Else
            Call PlayerMsg(index, "Ese no es un comando valido!", 12)
        End If
        Exit Sub
    
    Case "prompt"
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerPrompt " & index & "," & Val(Parse(1)) & "," & Val(Parse(2))
        End If
        Exit Sub
                
    Case "requesteditarrow"
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "arrowEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub

    Case "editarrow"
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

    Case "savearrow"
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

        Call SendUpdateArrowToAll(n)
        Call SaveArrow(n)
        Call AddLog(GetPlayerName(index) & " saved arrow #" & n & ".", ADMIN_LOG)
        Exit Sub
    Case "checkarrows"
        n = Arrows(Val(Parse(1))).Pic
        
        Call SendDataToMap(GetPlayerMap(index), "checkarrows" & SEP_CHAR & index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "requesteditemoticon"
        ' Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "EMOTICONEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub

    Case "editemoticon"
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

    Case "saveemoticon"
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
    
    Case "checkemoticons"
        n = Emoticons(Val(Parse(1))).Pic
        
        Call SendDataToMap(GetPlayerMap(index), "checkemoticons" & SEP_CHAR & index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Exit Sub
    
    Case "mapreport"
        Packs = "mapreport" & SEP_CHAR
        For i = 1 To MAX_MAPS
            Packs = Packs & Map(i).Name & SEP_CHAR
        Next i
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
        
    Case "arrowhit"
        n = Val(Parse(1))
        z = Val(Parse(2))
        x = Val(Parse(3))
        y = Val(Parse(4))
       
        If n = TARGET_TYPE_PLAYER Then
            ' Make sure we dont try To attack ourselves
            If z <> index Then
                ' Can we attack the player?
                If CanAttackPlayerWithArrow(index, z) Then
                    If CanPlayerDodgeHit(z) Then
                    Call BattleMsg(index, GetPlayerName(z) & " esquivo tu ataque!", Yellow, 0)
                    Call BattleMsg(z, "Tu esquivaste el ataque de " & GetPlayerName(index) & "!", Yellow, 1)
                    End If
                    If Not CanPlayerBlockHit(z) Then
                        ' Get the damage we can Do
                        If Not CanPlayerCriticalHit(index) Then
                            c = GetPlayerDamage(index) - GetPlayerProtection(z)
                            c1 = GetPlayerFireDamage(index) - GetPlayerFireProtection(z)
                            c2 = GetPlayerWaterDamage(index) - GetPlayerWaterProtection(z)
                            c3 = GetPlayerEarthDamage(index) - GetPlayerEarthProtection(z)
                            c4 = GetPlayerAirDamage(index) - GetPlayerAirProtection(z)
                            c5 = GetPlayerHeatDamage(index) - GetPlayerHeatProtection(z)
                            c6 = GetPlayerColdDamage(index) - GetPlayerColdProtection(z)
                            c7 = GetPlayerLightDamage(index) - GetPlayerLightProtection(z)
                            c8 = GetPlayerDarkDamage(index) - GetPlayerDarkProtection(z)
                        
                        
                        If c < 0 Then
                            c = 0
                        End If
                        If c1 < 0 Then
                            c1 = 0
                        End If
                        If c2 < 0 Then
                            c2 = 0
                        End If
                        If c3 < 0 Then
                            c3 = 0
                        End If
                        If c4 < 0 Then
                            c4 = 0
                        End If
                        If c5 < 0 Then
                            c5 = 0
                        End If
                        If c6 < 0 Then
                            c6 = 0
                        End If
                        If c7 < 0 Then
                            c7 = 0
                        End If
                        If c8 < 0 Then
                            c8 = 0
                        End If
                            Damage = c + c1
                            Damage = Damage + c2
                            Damage = Damage + c3
                            Damage = Damage + c4
                            Damage = Damage + c5
                            Damage = Damage + c6
                            Damage = Damage + c7
                            Damage = Damage + c8
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                        Else
                            m = GetPlayerDamage(index)
                            d = m + Int(Rnd * Int(m / 2)) + 1 - GetPlayerProtection(i)
                            m1 = GetPlayerFireDamage(index)
                            d1 = m1 + Int(Rnd * Int(m1 / 2)) + 1 - GetPlayerFireProtection(i)
                            m2 = GetPlayerWaterDamage(index)
                            d2 = m2 + Int(Rnd * Int(m2 / 2)) + 1 - GetPlayerWaterProtection(i)
                            m3 = GetPlayerEarthDamage(index)
                            d3 = m3 + Int(Rnd * Int(m3 / 2)) + 1 - GetPlayerEarthProtection(i)
                            m4 = GetPlayerAirDamage(index)
                            d4 = m4 + Int(Rnd * Int(m4 / 2)) + 1 - GetPlayerAirProtection(i)
                            m5 = GetPlayerHeatDamage(index)
                            d5 = m5 + Int(Rnd * Int(m5 / 2)) + 1 - GetPlayerHeatProtection(i)
                            m6 = GetPlayerColdDamage(index)
                            d6 = m6 + Int(Rnd * Int(m6 / 2)) + 1 - GetPlayerColdProtection(i)
                            m7 = GetPlayerLightDamage(index)
                            d7 = m7 + Int(Rnd * Int(m7 / 2)) + 1 - GetPlayerLightProtection(i)
                            m8 = GetPlayerDarkDamage(index)
                            d8 = m8 + Int(Rnd * Int(m8 / 2)) + 1 - GetPlayerDarkProtection(i)
                        If d < 0 Then
                            d = 0
                        End If
                        If d1 < 0 Then
                            d1 = 0
                        End If
                        If d2 < 0 Then
                            d2 = 0
                        End If
                        If d3 < 0 Then
                            d3 = 0
                        End If
                        If d4 < 0 Then
                            d4 = 0
                        End If
                        If d5 < 0 Then
                            d5 = 0
                        End If
                        If d6 < 0 Then
                            d6 = 0
                        End If
                        If d7 < 0 Then
                            d7 = 0
                        End If
                        If d8 < 0 Then
                            d8 = 0
                        End If
                            Damage = d + d1
                            Damage = Damage + d2
                            Damage = Damage + d3
                            Damage = Damage + d4
                            Damage = Damage + d5
                            Damage = Damage + d6
                            Damage = Damage + d7
                            Damage = Damage + d8
                            Call BattleMsg(index, "Daño Critico!", Yellow, 0)
                            Call BattleMsg(z, GetPlayerName(index) & " dispara con una punteria increible!", Yellow, 1)
                           
                            'Call PlayerMsg(index, "Sientes una gran energia mientras disparas!", BrightCyan)
                            'Call PlayerMsg(z, GetPlayerName(index) & " dispara con una punteria increible!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                        End If
                       
                        If Damage > 0 Then
                            Call AttackPlayer(index, z, Damage)
                        Else
                            Call BattleMsg(index, "Tu ataque no hace nada.", BrightRed, 0)
                            Call BattleMsg(z, GetPlayerName(z) & "'s attack did nothing.", BrightRed, 1)
                           
                            'Call PlayerMsg(index, "Tu ataque no hace nada.", BrightRed)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call BattleMsg(index, GetPlayerName(z) & " bloqueo tu golpe!", Yellow, 0)
                        Call BattleMsg(z, "Tu bloqueaste el golpe de " & GetPlayerName(index) & "!", Yellow, 1)
                       
                        'Call PlayerMsg(index, GetPlayerName(z) & " bloqueo tu golpe con su " & Trim(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & "!", BrightCyan)
                        'Call PlayerMsg(z, "Tu " & Trim(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " bloqueo el golpe de " & GetPlayerName(index) & "!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                    End If
                    Exit Sub
                End If
            End If
        ElseIf n = TARGET_TYPE_NPC Then
            ' Can we attack the npc?
            If CanAttackNpcWithArrow(index, z) Then
                ' Get the damage we can Do
                If Not CanPlayerCriticalHit(index) Then
                    c = GetPlayerDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).DEF)
                    c1 = GetPlayerFireDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).FireDEF)
                    c2 = GetPlayerWaterDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).WaterDEF)
                    c3 = GetPlayerEarthDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).EarthDEF)
                    c4 = GetPlayerAirDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).AirDEF)
                    c5 = GetPlayerHeatDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).HeatDEF)
                    c6 = GetPlayerColdDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).ColdDEF)
                    c7 = GetPlayerLightDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).LightDEF)
                    c8 = GetPlayerDarkDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).DarkDEF)
                    If c < 0 Then
                    c = 0
                    End If
                    If c1 < 0 Then
                    c1 = 0
                    End If
                    If c2 < 0 Then
                    c2 = 0
                    End If
                    If c3 < 0 Then
                    c3 = 0
                    End If
                    If c4 < 0 Then
                    c4 = 0
                    End If
                    If c5 < 0 Then
                    c5 = 0
                    End If
                    If c6 < 0 Then
                    c6 = 0
                    End If
                    If c7 < 0 Then
                    c7 = 0
                    End If
                    If c8 < 0 Then
                    c8 = 0
                    End If
                    Damage = c + c1
                    Damage = Damage + c2
                    Damage = Damage + c3
                    Damage = Damage + c4
                    Damage = Damage + c5
                    Damage = Damage + c6
                    Damage = Damage + c7
                    Damage = Damage + c8
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                Else
                    m = GetPlayerDamage(index)
                    d = m + Int(Rnd * Int(m / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).DEF / 2)
                    m1 = GetPlayerFireDamage(index)
                    d1 = m1 + Int(Rnd * Int(m1 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).FireDEF / 2)
                    m2 = GetPlayerWaterDamage(index)
                    d2 = m2 + Int(Rnd * Int(m2 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).WaterDEF / 2)
                    m3 = GetPlayerEarthDamage(index)
                    d3 = m3 + Int(Rnd * Int(m3 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).EarthDEF / 2)
                    m4 = GetPlayerAirDamage(index)
                    d4 = m4 + Int(Rnd * Int(m4 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).AirDEF / 2)
                    m5 = GetPlayerHeatDamage(index)
                    d5 = m5 + Int(Rnd * Int(m5 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).HeatDEF / 2)
                    m6 = GetPlayerColdDamage(index)
                    d6 = m6 + Int(Rnd * Int(m6 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).ColdDEF / 2)
                    m7 = GetPlayerLightDamage(index)
                    d7 = m7 + Int(Rnd * Int(m7 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).LightDEF / 2)
                    m8 = GetPlayerDarkDamage(index)
                    d8 = m8 + Int(Rnd * Int(m8 / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).DarkDEF / 2)
                    If d < 0 Then
                    d = 0
                    End If
                    If d1 < 0 Then
                    d1 = 0
                    End If
                    If d2 < 0 Then
                    d2 = 0
                    End If
                    If d3 < 0 Then
                    d3 = 0
                    End If
                    If d4 < 0 Then
                    d4 = 0
                    End If
                    If d5 < 0 Then
                    d5 = 0
                    End If
                    If d6 < 0 Then
                    d6 = 0
                    End If
                    If d7 < 0 Then
                    d7 = 0
                    End If
                    If d8 < 0 Then
                    d8 = 0
                    End If
                    Damage = d + d1
                    Damage = Damage + d2
                    Damage = Damage + d3
                    Damage = Damage + d4
                    Damage = Damage + d5
                    Damage = Damage + d6
                    Damage = Damage + d7
                    Damage = Damage + d8
                    Call BattleMsg(index, "Daño Critico!", Yellow, 0)
                   
                    'Call PlayerMsg(index, "Sientes una gran energia mientras golpeas!", BrightCyan)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                End If
               
                If Damage > 0 Then
                    Call AttackNpc(index, z, Damage)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                Else
                    Call BattleMsg(index, "Tu ataque no hace nada.", BrightRed, 0)
                   
                    'Call PlayerMsg(index, "Tu ataque no hace nada.", BrightRed)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                End If
                Exit Sub
            End If
        ElseIf n = TARGET_TYPE_ATTRIBUTE_NPC Then
            BX = Val(Parse(5))
            BY = Val(Parse(6))
            If Not CanPlayerCriticalHit(index) Then
                    Damage = GetPlayerDamage(index) - Int(Npc(MapAttributeNpc(GetPlayerMap(index), z, BX, BY).num).DEF / 2)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                Else
                    n = GetPlayerDamage(index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapAttributeNpc(GetPlayerMap(index), z, BX, BY).num).DEF / 2)
                    Call BattleMsg(index, "Daño Critico!", Yellow, 0)
                   
                    'Call PlayerMsg(index, "Sientes una gran energia mientras golpeas!", BrightCyan)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                End If
               
                If Damage > 0 Then
                    Call AttackAttributeNpc(z, BX, BY, index, Damage)
                    'Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                Else
                    Call BattleMsg(index, "Tu ataque no hace nada.", BrightRed, 0)
                   
                    'Call PlayerMsg(index, "Tu ataque no hace nada.", BrightRed)
                    'Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                End If
                Exit Sub
        End If
        Exit Sub
End Select

Select Case LCase(Parse(0))
   Case "bankdeposit"
       x = GetPlayerInvItemNum(index, Val(Parse(1)))
       i = FindOpenBankSlot(index, x)
       If i = 0 Then
           Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Banco lleno" & SEP_CHAR & END_CHAR)
           Exit Sub
       End If
       
       If Val(Parse(2)) > GetPlayerInvItemValue(index, Val(Parse(1))) Then
           Call SendDataTo(index, "bankmsg" & SEP_CHAR & "No puedes depositar mas de lo que tienes!" & SEP_CHAR & END_CHAR)
           Exit Sub
       End If
       
       If GetPlayerWeaponSlot(index) = Val(Parse(1)) Or GetPlayerArmorSlot(index) = Val(Parse(1)) Or GetPlayerShieldSlot(index) = Val(Parse(1)) Or GetPlayerHelmetSlot(index) = Val(Parse(1)) Or GetPlayerGlovesSlot(index) = Val(Parse(1)) Or GetPlayerBootsSlot(index) = Val(Parse(1)) Or GetPlayerBeltSlot(index) = Val(Parse(1)) Or GetPlayerLegsSlot(index) = Val(Parse(1)) Or GetPlayerRingSlot(index) = Val(Parse(1)) Or GetPlayerAmuletSlot(index) = Val(Parse(1)) Then
           Call SendDataTo(index, "bankmsg" & SEP_CHAR & "No puedes depostiar items equipados!" & SEP_CHAR & END_CHAR)
           Exit Sub
       End If
       
       If Item(x).Type = ITEM_TYPE_CURRENCY Then
           If Val(Parse(2)) <= 0 Then
               Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Debes depostiar mas que 0!" & SEP_CHAR & END_CHAR)
               Exit Sub
           End If
       End If
       
       Call TakeItem(index, x, Val(Parse(2)))
       Call GiveBankItem(index, x, Val(Parse(2)), i)
       
       Call SendBank(index)
       Exit Sub
   
   Case "bankwithdraw"
       i = GetPlayerBankItemNum(index, Val(Parse(1)))
       TempVal = Val(Parse(2))
       x = FindOpenInvSlot(index, i)
       If x = 0 Then
           Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Inventario lleno!" & SEP_CHAR & END_CHAR)
           Exit Sub
       End If
       
       If Val(Parse(2)) > GetPlayerBankItemValue(index, Val(Parse(1))) Then
           Call SendDataTo(index, "bankmsg" & SEP_CHAR & "No puedes retirar mas de lo que tienes!" & SEP_CHAR & END_CHAR)
           Exit Sub
       End If
               
       If Item(i).Type = ITEM_TYPE_CURRENCY Then
           If Val(Parse(2)) <= 0 Then
               Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Debes retirar mas que 0!" & SEP_CHAR & END_CHAR)
               Exit Sub
           End If

           If Trim(LCase(Item(GetPlayerInvItemNum(index, x)).Name)) <> "gold" Then
               If GetPlayerInvItemValue(index, x) + Val(Parse(2)) > 1999999999 Then
                   TempVal = 1999999999 - GetPlayerInvItemValue(index, x)
               End If
           End If
       End If
               
       Call GiveItem(index, i, TempVal)
       Call TakeBankItem(index, i, TempVal)
       
       Call SendBank(index)
       Exit Sub
End Select

Call HackingAttempt(index, "")
Exit Sub
End Sub

Sub CloseSocket(ByVal index As Long)
    ' Make sure player was/is playing the game, and if so, save'm.
    If index > 0 Then
        Call LeftGame(index)
    
        Call TextAdd(frmServer.txtText(0), "Coneccion desde " & GetPlayerIP(index) & " termino.", True)
        
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
        s = "No hay otros jugadores online."
    Else
        s = Mid(s, 1, Len(s) - 2)
        s = "Hay " & n & " jugadores online: " & s & "."
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
            Packet = Packet & END_CHAR
            Call SendDataTo(index, Packet)
        End If
    Next i
    
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
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
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
    Packet = Packet & END_CHAR
    Call SendDataToMapBut(index, MapNum, Packet)
End Sub

Sub SendPlayerData(ByVal index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
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
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long
Dim i As Long

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
        With Map(MapNum).Tile(x, y)
            Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .Light & SEP_CHAR
            Packet = Packet & .GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR
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
        If MapNum > 0 Then
            Packet = Packet & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
        End If
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        If MapNum > 0 Then
            Packet = Packet & MapNpc(MapNum, i).num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
        End If
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
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

Sub SendArrows(ByVal index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_ARROWS
        Call SendUpdateArrowTo(index, i)
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
   
 '  Packet = "PLAYERBANKUPDATE" & SEP_CHAR & BankSlot & SEP_CHAR & GetPlayerBankItemNum(index, BankSlot) & SEP_CHAR & GetPlayerBankItemValue(index, BankSlot) & SEP_CHAR & GetPlayerBankItemDur(index, BankSlot) & SEP_CHAR & END_CHAR
   Packet = "PLAYERBANKUPDATE" & SEP_CHAR & BankSlot & SEP_CHAR & index & SEP_CHAR & GetPlayerBankItemNum(index, BankSlot) & SEP_CHAR & GetPlayerBankItemValue(index, BankSlot) & SEP_CHAR & GetPlayerBankItemDur(index, BankSlot) & SEP_CHAR & index & SEP_CHAR & END_CHAR
   Call SendDataTo(index, Packet)
End Sub

Sub SendInventory(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR & index & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & index & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & index & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendWornEquipment(ByVal index As Long)
Dim Packet As String
    
    If IsPlaying(index) Then
        Packet = "PLAYERWORNEQ" & SEP_CHAR & index & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & GetPlayerGlovesSlot(index) & SEP_CHAR & GetPlayerBootsSlot(index) & SEP_CHAR & GetPlayerBeltSlot(index) & SEP_CHAR & GetPlayerLegsSlot(index) & SEP_CHAR & GetPlayerRingSlot(index) & SEP_CHAR & GetPlayerAmuletSlot(index) & SEP_CHAR & END_CHAR
        Call SendDataToMap(GetPlayerMap(index), Packet)
    End If
End Sub

Sub SendHP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
    
    Packet = "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(index) & SEP_CHAR & END_CHAR
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
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendNewCharClasses(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).MaleSprite & SEP_CHAR & Class(i).FemaleSprite & SEP_CHAR & Class(i).Locked & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
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
End Sub

Sub SendPlayerXY(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagiReq & SEP_CHAR & Item(ItemNum).LvlReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).FireSTR & SEP_CHAR & Item(ItemNum).FireDEF & SEP_CHAR & Item(ItemNum).WaterSTR & SEP_CHAR & Item(ItemNum).WaterDEF & SEP_CHAR & Item(ItemNum).EarthSTR & SEP_CHAR & Item(ItemNum).EarthDEF & SEP_CHAR & Item(ItemNum).AirSTR & SEP_CHAR & Item(ItemNum).AirDEF & SEP_CHAR & Item(ItemNum).HeatSTR & SEP_CHAR & Item(ItemNum).HeatDEF & SEP_CHAR & Item(ItemNum).ColdSTR & SEP_CHAR & Item(ItemNum).ColdDEF & SEP_CHAR & Item(ItemNum).LightSTR & SEP_CHAR & Item(ItemNum).LightDEF & SEP_CHAR & Item(ItemNum).DarkSTR & SEP_CHAR & Item(ItemNum).DarkDEF & SEP_CHAR & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    'Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagiReq & SEP_CHAR & Item(ItemNum).LvlReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).FireSTR & SEP_CHAR & Item(ItemNum).FireDEF & SEP_CHAR & Item(ItemNum).WaterSTR & SEP_CHAR & Item(ItemNum).WaterDEF & SEP_CHAR & Item(ItemNum).EarthSTR & SEP_CHAR & Item(ItemNum).EarthDEF & SEP_CHAR & Item(ItemNum).AirSTR & SEP_CHAR & Item(ItemNum).AirDEF & SEP_CHAR & Item(ItemNum).HeatSTR & SEP_CHAR & Item(ItemNum).HeatDEF & SEP_CHAR & Item(ItemNum).ColdSTR & SEP_CHAR & Item(ItemNum).ColdDEF & SEP_CHAR & Item(ItemNum).LightSTR & SEP_CHAR & Item(ItemNum).LightDEF & SEP_CHAR & Item(ItemNum).DarkSTR & SEP_CHAR & Item(ItemNum).DarkDEF & SEP_CHAR & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditItemTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagiReq & SEP_CHAR & Item(ItemNum).LvlReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).FireSTR & SEP_CHAR & Item(ItemNum).FireDEF & SEP_CHAR & Item(ItemNum).WaterSTR & SEP_CHAR & Item(ItemNum).WaterDEF & SEP_CHAR & Item(ItemNum).EarthSTR & SEP_CHAR & Item(ItemNum).EarthDEF & SEP_CHAR & Item(ItemNum).AirSTR & SEP_CHAR & Item(ItemNum).AirDEF & SEP_CHAR & Item(ItemNum).HeatSTR & SEP_CHAR & Item(ItemNum).HeatDEF & SEP_CHAR & Item(ItemNum).ColdSTR & SEP_CHAR & Item(ItemNum).ColdDEF & SEP_CHAR & Item(ItemNum).LightSTR & SEP_CHAR & Item(ItemNum).LightDEF & SEP_CHAR & Item(ItemNum).DarkSTR & SEP_CHAR & Item(ItemNum).DarkDEF & SEP_CHAR & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed
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

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEArrow" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateArrowTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEArrow" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditArrowTo(ByVal index As Long, ByVal EmoNum As Long)
Dim Packet As String

    Packet = "EDITArrow" & SEP_CHAR & EmoNum & SEP_CHAR & Trim(Arrows(EmoNum).Name) & SEP_CHAR & END_CHAR
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
    Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).FireSTR & SEP_CHAR & Npc(NpcNum).FireDEF & SEP_CHAR & Npc(NpcNum).WaterSTR & SEP_CHAR & Npc(NpcNum).WaterDEF & SEP_CHAR & Npc(NpcNum).EarthSTR & SEP_CHAR & Npc(NpcNum).EarthDEF & SEP_CHAR & Npc(NpcNum).AirSTR & SEP_CHAR & Npc(NpcNum).AirDEF & SEP_CHAR & Npc(NpcNum).HeatSTR & SEP_CHAR & Npc(NpcNum).HeatDEF & SEP_CHAR & Npc(NpcNum).ColdSTR & SEP_CHAR & Npc(NpcNum).ColdDEF & SEP_CHAR & Npc(NpcNum).LightSTR & SEP_CHAR & Npc(NpcNum).LightDEF & SEP_CHAR & Npc(NpcNum).DarkSTR & SEP_CHAR & Npc(NpcNum).DarkDEF & SEP_CHAR & Npc(NpcNum).Speed & SEP_CHAR & Npc(NpcNum).Magi & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR
    Packet = Packet & Npc(NpcNum).SpawnTime & Npc(NpcNum).SpawnTime & SEP_CHAR
 
     
    
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
Dim i As Long, z As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To 6
        For z = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(i).Value(z).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GetValue & SEP_CHAR
        Next z
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

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).FireSTR & SEP_CHAR & Spell(SpellNum).WaterSTR & SEP_CHAR & Spell(SpellNum).EarthSTR & SEP_CHAR & Spell(SpellNum).AirSTR & SEP_CHAR & Spell(SpellNum).HeatSTR & SEP_CHAR & Spell(SpellNum).ColdSTR & SEP_CHAR & Spell(SpellNum).LightSTR & SEP_CHAR & Spell(SpellNum).DarkSTR & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendTrade(ByVal index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long, x As Long, y As Long, z As Long, XX As Long

    z = 0
    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To 6
        For XX = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(i).Value(XX).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(XX).GetValue & SEP_CHAR
            
            ' Item #
            x = Shop(ShopNum).TradeItem(i).Value(XX).GetItem
            
            If Item(x).Type = ITEM_TYPE_SPELL Then
                ' Spell class requirement
                y = Spell(Item(x).Data1).ClassReq
                
                If y = 0 Then
                    Call PlayerMsg(index, Trim(Item(x).Name) & " puede ser usado por todas las clases.", Yellow)
                Else
                    Call PlayerMsg(index, Trim(Item(x).Name) & " solo puede ser usado por un/a " & GetClassName(y - 1) & ".", Yellow)
                End If
            End If
            If x < 1 Then
                z = z + 1
            End If
        Next XX
    Next i
    Packet = Packet & END_CHAR
    
    If z = (MAX_TRADES * 6) Then
        Call PlayerMsg(index, "Esta tienda no tiene nada para vender!", BrightRed)
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
    If RainIntensity <= 0 Then RainIntensity = 1
    Packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & RainIntensity & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendWeatherToAll()
Dim i As Long
Dim Weather As String
    
    Select Case GameWeather
        Case 0
            Weather = "Nada"
        Case 1
            Weather = "Lluvia"
        Case 2
            Weather = "Nieve"
        Case 3
            Weather = "Truenos"
    End Select
    frmServer.Label5.Caption = "Clima Actual: " & Weather
    
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
    
    Call SpawnAllMapNpcs
    Call SpawnAllMapAttributeNpcs
End Sub

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal index As Long)
Dim Packet As String

    Packet = "MAPMSG2" & SEP_CHAR & Msg & SEP_CHAR & index & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub
