Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.caption = GAME_NAME & " - Eclipse Evolution Server"
    frmServer.lblIP.caption = "Ip Address: " & frmServer.Socket(0).LocalIP
    frmServer.lblPort.caption = "Port: " & STR(frmServer.Socket(0).LocalPort)
    frmServer.TPO.caption = "Total Players Online: " & TotalOnlinePlayers
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    On Error Resume Next
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
    Dim FileName As String, fIP As String, FName As String
    Dim f As Long
    
    IsBanned = False
    
    FileName = App.Path & "\banlist.txt"
    
'    Check if file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    f = FreeFile
    Open FileName For Input As #f
    
    Do While Not EOF(f)
        Input #f, fIP
        Input #f, FName
        
'        Is banned?
        If Trim(LCase(fIP)) = Trim(LCase(Mid(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If
    Loop
    
    Close #f
End Function

Sub SendDataTo(ByVal index As Long, ByVal Data As String)
    
    If IsConnected(index) Then
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
    Exit Sub
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
'            Whoho, we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If
End Sub

Sub SocketConnected(ByVal index As Long)
    If index <> 0 Then
'        Are they trying to connect more then one connection?
'        If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
        If Not IsBanned(GetPlayerIP(index)) Then
            Call TextAdd(frmServer.txtText(0), "Received connection from " & GetPlayerIP(index) & ".", True)
        Else
'            LMAO PWNED ROFL!!!!11/
            Call AlertMsg(index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
        End If
'        Else
'        Tried multiple connections
'        Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
'        End If
    End If
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
    Dim Buffer As String
    Dim OPlayer As PlayerRec
    Dim Packet As String
    Dim Playernum As Long
    Dim top As String * 3
    Dim Start As Long
    Dim Name As String
    Dim Password As String
    Dim PASSWORD2 As String
    Dim Sex As Long
    Dim Class As Long
    Dim MAXHP As Long
    Dim MAXMP As Long
    Dim MAXSP As Long
    Dim addHPP As Long
    Dim addMPP As Long
    Dim addSPP As Long
    Dim charnum As Long
    Dim Msg As String
    Dim MsgTo As Long
    Dim Dir As Long
    Dim InvNum As Long
    Dim Amount As Long
    Dim Damage As Long
    Dim PointType As Long
    Dim Movement As Long
    Dim i As Long, n As Long, x As Long, y As Long, f As Long
    Dim MapNum As Long
    Dim s As String
    Dim ShopNum As Integer, ItemNum As Long
    Dim DurNeeded As Long, GoldNeeded As Long
    Dim z As Long
    Dim BX As Long, BY As Long
    Dim TempVal As Long
    Dim hfile
    Dim ItemName As String
    Dim m As Long
    Dim j As Long
    Dim Imail As String
    Dim Imail2 As String
    Dim surge As String
    Dim Parse() As String
    Dim q As Long
    Dim p As Long
    Dim count As Long
    Dim output As String
    Dim d As String
    Dim slot As Long
    Dim itemnumber As Long
    Dim chlist As String
    Dim getmax As Double
    Dim topcount As Long
    Dim PlayerNames() As String
    Dim PlayerLevels() As String
    Dim TopNames() As String
    Dim TopLevels() As String
    Dim FileName As String
    
    
    
    If index > 0 Then
        frmServer.Socket(index).GetData Buffer, vbString, DataLength
        
        player(index).Buffer = player(index).Buffer & Buffer
        
        
        Start = InStr(player(index).Buffer, END_CHAR)
        Do While Start > 0
            Packet = Mid(player(index).Buffer, 1, Start - 1)
            player(index).Buffer = Mid(player(index).Buffer, Start + 1, Len(player(index).Buffer))
            player(index).DataPackets = player(index).DataPackets + 1
            Start = InStr(player(index).Buffer, END_CHAR)
            If Len(Packet) > 0 Then
                Call HandleData(index, Packet)
            End If
        Loop
        
        '        Not useful
'        Check if elapsed time has passed
        player(index).DataBytes = player(index).DataBytes + DataLength
        If GetTickCount >= player(index).DataTimer + 1000 Then
            player(index).DataTimer = GetTickCount
            player(index).DataBytes = 0
            player(index).DataPackets = 0
            ' Exit Sub
        End If
        
'        Check for data flooding
        If player(index).DataBytes > 1000 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Data Flooding")
            Exit Sub
        End If
        
'        Check for packet flooding
        If player(index).DataPackets > 25 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Packet Flooding")
            Exit Sub
        End If
    End If
        
        If Buffer = "top" Then
            top = STR(TotalOnlinePlayers)
            Call SendDataTo(index, top)
            Exit Sub
        End If
        
        
        If Buffer = "updatesite" Then
            Call SendDataTo(index, "Done.")
            Exit Sub
        End If
        
        
        If Buffer = "setevents" Then
            If Len(Buffer) > 10 Then
                Buffer = Mid(Buffer, 10, Len(Buffer) - 9)
                Call PutVar(App.Path & "\Events.ini", "DATA", "Events", Buffer)
                Call SendDataTo(index, "Done.")
                Exit Sub
            End If
        End If
        
        If LCase(Mid(Buffer, 1, 19)) = "getplayerspriteinfo" Then
            Name = Mid(Buffer, 21)
            Name = Trim$(Name)
            If FindChar(Name) Then
                If FindPlayer(Name) = 0 Then
                    Call Findcharfile(Name)
                    For q = 1 To MAX_CHARS
                        If LCase(Trim$(tempplayer.Char(q).Name)) = LCase(Name) Then
                            output = "sprite_y=" & tempplayer.Char(q).Sprite * 64
                            If tempplayer.Char(q).WeaponSlot > 0 Then
                                slot = tempplayer.Char(q).WeaponSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&weapon_y=" & Item(itemnumber).Pic * 64
                            Else
                                output = output & "&weapon_y=0"
                            End If
                            If tempplayer.Char(q).ArmorSlot > 0 Then
                                slot = tempplayer.Char(q).ArmorSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&armor_y=" & Item(itemnumber).Pic * 64
                            Else
                                output = output & "&armor_y=0"
                            End If
                            If tempplayer.Char(q).HelmetSlot > 0 Then
                                slot = tempplayer.Char(q).HelmetSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&helmet_y=" & Item(itemnumber).Pic * 64
                            Else
                                output = output & "&helmet_y=0"
                            End If
                            If tempplayer.Char(q).ShieldSlot > 0 Then
                                slot = tempplayer.Char(q).ShieldSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&shield_y=" & Item(itemnumber).Pic * 64
                            Else
                                output = output & "&shield_y=0"
                            End If
                            If tempplayer.Char(q).LegsSlot > 0 Then
                                slot = tempplayer.Char(q).LegsSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&legs_y=" & Item(itemnumber).Pic * 64
                            Else
                                output = output & "&legs_y=0"
                            End If
                            If tempplayer.Char(q).RingSlot > 0 Then
                                slot = tempplayer.Char(q).RingSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&ring_y=" & Item(itemnumber).Pic * 64
                            Else
                                output = output & "&ring_y=0"
                            End If
                            If tempplayer.Char(q).NecklaceSlot > 0 Then
                                slot = tempplayer.Char(q).NecklaceSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&necklace_y=" & Item(itemnumber).Pic * 64
                            Else
                                output = output & "&necklace_y=0"
                            End If
                            Call SendDataTo(index, output)
                            Exit Sub
                        End If
                    Next q
                Else
                    i = FindPlayer(Trim$(Name))
                    output = "sprite_y=" & player(i).Char(player(i).charnum).Sprite * 64
                    If player(i).Char(player(i).charnum).WeaponSlot > 0 Then
                        slot = player(i).Char(player(i).charnum).WeaponSlot
                        itemnumber = Int(player(i).Char(player(i).charnum).Inv(slot).num)
                        output = output & "&weapon_y=" & Item(itemnumber).Pic * 64
                    Else
                        output = output & "&weapon_y=0"
                    End If
                    If player(i).Char(player(i).charnum).ArmorSlot > 0 Then
                        slot = player(i).Char(player(i).charnum).ArmorSlot
                        itemnumber = Int(player(i).Char(player(i).charnum).Inv(slot).num)
                        output = output & "&armor_y=" & Item(itemnumber).Pic * 64
                    Else
                        output = output & "&armor_y=0"
                    End If
                    If player(i).Char(player(i).charnum).HelmetSlot > 0 Then
                        slot = player(i).Char(player(i).charnum).HelmetSlot
                        itemnumber = Int(player(i).Char(player(i).charnum).Inv(slot).num)
                        output = output & "&helmet_y=" & Item(itemnumber).Pic * 64
                    Else
                        output = output & "&helmet_y=0"
                    End If
                    If player(i).Char(player(i).charnum).ShieldSlot > 0 Then
                        slot = player(i).Char(player(i).charnum).ShieldSlot
                        itemnumber = Int(player(i).Char(player(i).charnum).Inv(slot).num)
                        output = output & "&shield_y=" & Item(itemnumber).Pic * 64
                    Else
                        output = output & "&shield_y=0"
                    End If
                    If player(i).Char(player(i).charnum).LegsSlot > 0 Then
                        slot = player(i).Char(player(i).charnum).LegsSlot
                        itemnumber = Int(player(i).Char(player(i).charnum).Inv(slot).num)
                        output = output & "&legs_y=" & Item(itemnumber).Pic * 64
                    Else
                        output = output & "&legs_y=0"
                    End If
                    If player(i).Char(player(i).charnum).RingSlot > 0 Then
                        slot = player(i).Char(player(i).charnum).RingSlot
                        itemnumber = Int(player(i).Char(player(i).charnum).Inv(slot).num)
                        output = output & "&ring_y=" & Item(itemnumber).Pic * 64
                    Else
                        output = output & "&ring_y=0"
                    End If
                    If player(i).Char(player(i).charnum).NecklaceSlot > 0 Then
                        slot = player(i).Char(player(i).charnum).NecklaceSlot
                        itemnumber = Int(player(i).Char(player(i).charnum).Inv(slot).num)
                        output = output & "&necklace_y=" & Item(itemnumber).Pic * 64
                    Else
                        output = output & "&necklace_y=0"
                    End If
                    Call SendDataTo(index, output)
                    Exit Sub
                End If
            Else
                output = "Character does not exist!"
                Call SendDataTo(index, output)
                Exit Sub
            End If
        End If
        
        If Mid(Buffer, 1, 10) = "checklogin" Then
            Dim logparse() As String
            Dim BkUser As String
            Dim BkPass As String
            Dim found As Long
            Dim match As Long
            logparse = Split(Buffer, ",")
            BkUser = logparse(1)
            BkPass = logparse(2)
            Dim fso, fldr, folders, fldrnm
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set fldr = fso.GetFolder(App.Path & "\accounts")
            Set folders = fldr.subfolders
            Call CleartempPlayer
            found = 0
            match = 0
            For Each fldrnm In folders
                If LCase(Trim$(fldrnm.Name)) = LCase(Trim$(BkUser)) Then
                    found = 1
                    If GetVar(App.Path & "\accounts\" & fldrnm.Name & "_info.ini", "ACCESS", "Password") = BkPass Then
                        match = 1
                        Exit For
                    Else
                        match = 0
                        Exit For
                    End If
                End If
            Next fldrnm
            If found = 1 Then
                If match = 1 Then
                    Call SendDataTo(index, "accepted")
                Else
                    Call SendDataTo(index, "rejected")
                End If
            Else
                Call SendDataTo(index, "notfound")
            End If
        End If
        
        If Mid(Buffer, 1, 14) = "changeuserpass" Then
            Dim newpass As String
            logparse = Split(Buffer, ",")
            BkUser = logparse(1)
            BkPass = logparse(2)
            newpass = logparse(3)
            ' Dim fso, fldr, folders, fldrnm
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set fldr = fso.GetFolder(App.Path & "\accounts")
            Set folders = fldr.subfolders
            For Each fldrnm In folders
                If LCase(Trim$(fldrnm.Name)) = LCase(Trim$(BkUser)) Then
                    found = 1
                    If GetVar(App.Path & "\accounts\" & fldrnm.Name & "_info.ini", "ACCESS", "Password") = BkPass Then
                        match = 1
                        BkUser = fldrnm.Name
                        Exit For
                    Else
                        match = 0
                        Exit For
                    End If
                End If
            Next fldrnm
            If found = 1 Then
                If match = 1 Then
                    Call PutVar(App.Path & "\accounts\" & BkUser & "_info.ini", "ACCESS", "Password", newpass)
                    Call SendDataTo(index, "success")
                Else
                    Call SendDataTo(index, "rejected")
                End If
            Else
                Call SendDataTo(index, "notfound")
            End If
        End If
        
        If LCase(Mid(Buffer, 1, 18)) = "getplayerhighscore" Then
            Dim spparse() As String
            spparse = Split(Buffer, ",")
            ReDim PlayerNames(MAX_PLAYERS)
            ReDim PlayerLevels(MAX_PLAYERS)
            ReDim TopNames(MAX_PLAYERS)
            ReDim TopLevels(MAX_PLAYERS)
            Dim getname As String
            f = FreeFile
            getname = spparse(1)
            x = 0
            
            
            FileName = App.Path & "\Accounts\charlist.txt"
            Open FileName For Input As #f
                Do While Not EOF(f)
                    x = x + 1
                    Input #f, PlayerNames(x)
                Loop
            Close #f
            ReDim Preserve PlayerNames(x)
            Dim PlayerClasses() As String
            ReDim Preserve PlayerClasses(MAX_PLAYERS)
            Dim TopClasses() As String
            ReDim Preserve TopClasses(MAX_PLAYERS)
            Dim CCNAME As String
            Dim r As Long
            For r = 1 To UBound(PlayerNames()) + 1
            If r = UBound(PlayerNames()) + 1 Then
            Call SendDataTo(index, "Character does not exist!")
            Exit Sub
            End If
            If Trim$(LCase(PlayerNames(r))) = Trim$(LCase(getname)) Then
            For i = 1 To UBound(PlayerNames())
                CCNAME = PlayerNames(i)
                If FindPlayer(CCNAME) = 0 Then
                    Call Findcharfile(PlayerNames(i))
                    For x = 1 To MAX_CHARS
                        If Trim$(tempplayer.Char(x).Name) = Trim$(PlayerNames(i)) Then
                            f = x
                        End If
                    Next x
                    PlayerLevels(i) = Trim$(tempplayer.Char(f).Level)
                    PlayerClasses(i) = Trim$(GetVar(App.Path & "\Classes\Class" & tempplayer.Char(f).Class & ".ini", "CLASS", "Name"))
                Else
                    f = FindPlayer(PlayerNames(i))
                    PlayerLevels(i) = player(f).Char(player(f).charnum).Level
                    PlayerClasses(i) = Trim$(GetVar(App.Path & "\Classes\Class" & player(f).Char(player(f).charnum).Class & ".ini", "CLASS", "Name"))
                End If
            Next i
            Dim gopos As Long
            For x = 1 To UBound(PlayerNames())
                topcount = UBound(PlayerNames())
                For f = 1 To UBound(PlayerNames())
                    If PlayerLevels(x) > PlayerLevels(f) Then
                        topcount = topcount - 1
                    End If
                Next f
                If TopNames(topcount) = "" Then
                    TopNames(topcount) = PlayerNames(x)
                    TopLevels(topcount) = PlayerLevels(x)
                    TopClasses(topcount) = PlayerClasses(x)
                    If Trim$(LCase(TopNames(topcount))) = Trim$(LCase(getname)) Then
                    gopos = topcount
                    End If
                Else
                    Dim go As Long
                    go = 0
                    Do While go = 0
                    topcount = topcount + 1
                    If UBound(TopNames()) < topcount Then
                        ReDim Preserve TopNames(topcount)
                        ReDim Preserve TopLevels(topcount)
                        ReDim Preserve TopClasses(topcount)
                    End If
                    If TopNames(topcount) = "" Then
                    TopNames(topcount) = PlayerNames(x)
                    TopLevels(topcount) = PlayerLevels(x)
                    TopClasses(topcount) = PlayerClasses(x)
                    If Trim$(LCase(TopNames(topcount))) = Trim$(LCase(getname)) Then
                    gopos = topcount
                    End If
                    go = 1
                    End If
                    Loop
                End If
            Next x
            If getmax > UBound(TopNames()) Then
                ReDim Preserve TopNames(getmax)
                ReDim Preserve TopLevels(getmax)
                ReDim Preserve TopClasses(getmax)
            End If
            output = ""
            Dim spacecount As Long
            spacecount = 0
            For i = 1 To UBound(TopNames())
            x = i
                
                 If Trim$(TopNames(i)) = "" Then
                If x < UBound(TopNames()) Then
                Do While Trim$(TopNames(x)) = ""
                If x < UBound(TopNames) Then
                x = x + 1
                spacecount = spacecount + 1
                Else
                Exit Do
                End If
                Loop
                If LCase(Trim$(TopNames(x))) = getname Then
                output = output & "&name=" & Trim$(TopNames(x)) & "&level=" & Trim$(TopLevels(x)) & "&class=" & Trim$(TopClasses(x)) & "&pos=" & (gopos - spacecount)
                Exit For
                TopNames(x) = ""
                TopLevels(x) = ""
                Else
                TopNames(x) = ""
                TopLevels(x) = ""
                spacecount = 0
                End If
                End If
                Else
                spacecount = 0
                If LCase(Trim$(TopNames(i))) = getname Then
                output = output & "&name=" & Trim$(TopNames(i)) & "&level=" & Trim$(TopLevels(i)) & "&class=" & Trim$(TopClasses(i)) & "&pos=" & i
                Exit For
                End If
                End If
            Next i
            Call SendDataTo(index, output)
            Exit Sub
            End If
            Next r
        End If
        
        If LCase(Mid(Buffer, 1, 13)) = "gethighscores" Then
           ' Dim spparse() As String
            spparse = Split(Buffer, ",")
            ReDim PlayerNames(MAX_PLAYERS)
            ReDim PlayerLevels(MAX_PLAYERS)
            ReDim TopNames(MAX_PLAYERS)
            ReDim TopLevels(MAX_PLAYERS)
            Dim getclass
            f = FreeFile
            If Not IsNumeric(spparse(1)) Then
            Call SendDataTo(index, "Invalid Top!")
            Exit Sub
            Else
            If spparse(1) > 100 Then
            Call SendDataTo(index, "Invalid Top!")
            Exit Sub
            End If
            End If
            getmax = spparse(1)
            getclass = spparse(2)
            x = 0
            
            
            FileName = App.Path & "\Accounts\charlist.txt"
            Open FileName For Input As #f
                Do While Not EOF(f)
                    x = x + 1
                    Input #f, PlayerNames(x)
                Loop
            Close #f
            ReDim Preserve PlayerNames(x)
            ReDim Preserve PlayerLevels(x)
           ' Dim PlayerClasses() As String
            ReDim Preserve PlayerClasses(x)
            ReDim Preserve TopNames(x)
            ReDim Preserve TopLevels(x)
           ' Dim TopClasses() As String
            ReDim Preserve TopClasses(x)
           ' Dim CCNAME As String
            
            For i = 1 To UBound(PlayerNames())
                CCNAME = PlayerNames(i)
                If FindPlayer(CCNAME) = 0 Then
                    Call Findcharfile(PlayerNames(i))
                    For x = 1 To MAX_CHARS
                        If Trim$(tempplayer.Char(x).Name) = Trim$(PlayerNames(i)) Then
                            f = x
                        End If
                    Next x
                    PlayerLevels(i) = Trim$(tempplayer.Char(f).Level)
                    PlayerClasses(i) = Trim$(GetVar(App.Path & "\Classes\Class" & tempplayer.Char(f).Class & ".ini", "CLASS", "Name"))
                Else
                    f = FindPlayer(PlayerNames(i))
                    PlayerLevels(i) = player(f).Char(player(f).charnum).Level
                    PlayerClasses(i) = Trim$(GetVar(App.Path & "\Classes\Class" & player(f).Char(player(f).charnum).Class & ".ini", "CLASS", "Name"))
                End If
            Next i
            For x = 1 To UBound(PlayerNames())
                topcount = 1
                For f = 1 To UBound(PlayerNames())
                    If PlayerLevels(x) < PlayerLevels(f) Then
                        topcount = topcount + 1
                    End If
                Next f
                If TopNames(topcount) = "" Then
                    TopNames(topcount) = PlayerNames(x)
                    TopLevels(topcount) = PlayerLevels(x)
                    TopClasses(topcount) = PlayerClasses(x)
                Else
              '      Dim go As Long
                    go = 0
                    Do While go = 0
                    topcount = topcount + 1
                    If UBound(TopNames()) < topcount Then
                        ReDim Preserve TopNames(topcount)
                        ReDim Preserve TopLevels(topcount)
                        ReDim Preserve TopClasses(topcount)
                    End If
                    If TopNames(topcount) = "" Then
                    TopNames(topcount) = PlayerNames(x)
                    TopLevels(topcount) = PlayerLevels(x)
                    TopClasses(topcount) = PlayerClasses(x)
                    go = 1
                    End If
                    Loop
                End If
            Next x
            If getmax > UBound(TopNames()) Then
                ReDim Preserve TopNames(getmax)
                ReDim Preserve TopLevels(getmax)
                ReDim Preserve TopClasses(getmax)
            End If
            output = "max=" & getmax
            For i = 1 To getmax
            x = i
                
                 If Trim$(TopNames(i)) = "" Then
                If x < UBound(TopNames()) Then
                Do While Trim$(TopNames(x)) = ""
                If x < UBound(TopNames) Then
                x = x + 1
                Else
                Exit Do
                End If
                Loop
                If getclass = "all" Or Trim$(LCase(TopClasses(x))) = Trim$(LCase(getclass)) Then
                output = output & "&names[]=" & Trim$(TopNames(x)) & "&levels[]=" & Trim$(TopLevels(x)) & "&classes[]=" & Trim$(TopClasses(x))
                TopNames(x) = ""
                TopLevels(x) = ""
                Else
                TopNames(x) = ""
                TopLevels(x) = ""
                End If
                End If
                Else
                If getclass = "all" Or Trim$(LCase(TopClasses(i))) = Trim$(LCase(getclass)) Then
                output = output & "&names[]=" & Trim$(TopNames(i)) & "&levels[]=" & Trim$(TopLevels(i)) & "&classes[]=" & Trim$(TopClasses(i))
                End If
                End If
            Next i
            Call SendDataTo(index, output)
            Exit Sub
        End If
        
        If LCase(Mid(Buffer, 1, 12)) = "getitemstats" Then
            ItemName = Mid(Buffer, 14)
            q = 0
            For x = 1 To MAX_ITEMS
                If Trim$(Item(x).Name) = Trim$(ItemName) Then
                    output = "name=" & Trim$(Item(x).Name)
                    output = output & "&icon_y=" & Int(Item(x).Pic / 6)
                    output = output & "&icon_x=" & (Item(x).Pic - Int(Item(x).Pic / 6) * 6)
                    output = output & "&price=" & STR(Item(x).Price)
                    output = output & "&desc=" & Trim$(Item(x).Desc)
                    If Item(x).Type = ITEM_TYPE_NONE Then
                        output = output & "&type=None"
                    ElseIf Item(x).Type = ITEM_TYPE_WEAPON Then
                        output = output & "&type=Weapon"
                    ElseIf Item(x).Type = ITEM_TYPE_TWO_HAND Then
                        output = output & "&type=Two-Handed+Weapon"
                    ElseIf Item(x).Type = ITEM_TYPE_ARMOR Then
                        output = output & "&type=Armor"
                    ElseIf Item(x).Type = ITEM_TYPE_SHIELD Then
                        output = output & "&type=Shield"
                    ElseIf Item(x).Type = ITEM_TYPE_HELMET Then
                        output = output & "&type=Helmet"
                    ElseIf Item(x).Type = ITEM_TYPE_LEGS Then
                        output = output & "&type=Legs"
                    ElseIf Item(x).Type = ITEM_TYPE_RING Then
                        output = output & "&type=Ring"
                    ElseIf Item(x).Type = ITEM_TYPE_NECKLACE Then
                        output = output & "&type=Armor"
                    ElseIf Item(x).Type = ITEM_TYPE_POTIONADDHP Then
                        output = output & "&type=HP+Up+Potion"
                    ElseIf Item(x).Type = ITEM_TYPE_POTIONADDSP Then
                        output = output & "&type=SP+Up+Potion"
                    ElseIf Item(x).Type = ITEM_TYPE_POTIONADDMP Then
                        output = output & "&type=MP+Up+Potion"
                    ElseIf Item(x).Type = ITEM_TYPE_POTIONSUBHP Then
                        output = output & "&type=HP+Down+Potion"
                    ElseIf Item(x).Type = ITEM_TYPE_POTIONSUBSP Then
                        output = output & "&type=SP+Down+Potion"
                    ElseIf Item(x).Type = ITEM_TYPE_POTIONSUBMP Then
                        output = output & "&type=MP+Down+Potion"
                    ElseIf Item(x).Type = ITEM_TYPE_KEY Then
                        output = output & "&type=Key"
                    ElseIf Item(x).Type = ITEM_TYPE_CURRENCY Then
                        output = output & "&type=Currency"
                    ElseIf Item(x).Type = ITEM_TYPE_SPELL Then
                        output = output & "&type=Spell"
                    ElseIf Item(x).Type = ITEM_TYPE_SCRIPTED Then
                        output = output & "&type=Other"
                    End If
                    output = output & "&strreq=" & STR(Item(x).StrReq)
                    output = output & "&defreq=" & STR(Item(x).DefReq)
                    output = output & "&speedreq=" & STR(Item(x).SpeedReq)
                    If Item(x).ClassReq > -1 Then
                    output = output & "&classreq=" & GetVar(App.Path & "\classes\class" & Item(x).ClassReq & ".ini", "CLASS", "Name")
                    Else
                    output = output & "&classreq=None"
                    End If
                    output = output & "&addhp=" & STR(Item(x).addHP)
                    output = output & "&addmp=" & STR(Item(x).addMP)
                    output = output & "&addsp=" & STR(Item(x).addSP)
                    output = output & "&addstr=" & STR(Item(x).AddStr)
                    output = output & "&adddef=" & STR(Item(x).AddDef)
                    output = output & "&addmagi=" & STR(Item(x).AddMagi)
                    output = output & "&addspeed=" & STR(Item(x).AddSpeed)
                    output = output & "&addexp=" & STR(Item(x).AddEXP)
                    Call SendDataTo(index, output)
                    Exit Sub
                    q = 1
                End If
            Next
            If q = 0 Then
            output = "Item doesn't Exist!"
            Call SendDataTo(index, output)
            Exit Sub
            End If
        End If
        
        If LCase(Mid(Buffer, 1, 17)) = "getiniplayerstats" Then
            Name = Mid(Buffer, 19)
            Name = Trim$(Name)
            If FindChar(Name) Then
                If FindPlayer(Name) = 0 Then
                    Call LoadTempPlayerFromINI(Name)
                    For q = 1 To MAX_CHARS
                        If LCase(Trim$(tempplayer.Char(q).Name)) = LCase(Name) Then
                            output = "name=" & Trim$(tempplayer.Char(q).Name) & "&status=Offline" & "&level=" & STR(tempplayer.Char(q).Level) & "&class=" & GetVar(App.Path & "\classes\class" & tempplayer.Char(q).Class & ".ini", "CLASS", "Name") & " "
                            If tempplayer.Char(q).Sex = 0 Then
                                output = output & "&sex=Male"
                            Else
                                output = output & "&sex=Female"
                            End If
                            output = output & "&HP=" & STR(tempplayer.Char(q).MAXHP) & "&MP=" & STR(tempplayer.Char(q).MAXMP) & "&SP=" & STR(tempplayer.Char(q).MAXSP) & "&STR=" & STR(tempplayer.Char(q).STR) & "&DEF=" & STR(tempplayer.Char(q).DEF) & "&MAGI=" & STR(tempplayer.Char(q).Magi) & "&SPEED=" & STR(tempplayer.Char(q).Speed)
                            If tempplayer.Char(q).WeaponSlot > 0 Then
                                slot = tempplayer.Char(q).WeaponSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Weapon=" & Item(itemnumber).Name
                            Else
                                output = output & "&Weapon=None"
                            End If
                            If tempplayer.Char(q).ArmorSlot > 0 Then
                                slot = tempplayer.Char(q).ArmorSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Armor=" & Item(itemnumber).Name
                            Else
                                output = output & "&Armor=None"
                            End If
                            If tempplayer.Char(q).HelmetSlot > 0 Then
                                slot = tempplayer.Char(q).HelmetSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Helmet=" & Item(itemnumber).Name
                            Else
                                output = output & "&Helmet=None"
                            End If
                            If tempplayer.Char(q).ShieldSlot > 0 Then
                                slot = tempplayer.Char(q).ShieldSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Shield=" & Item(itemnumber).Name
                            Else
                                output = output & "&Shield=None"
                            End If
                            If tempplayer.Char(q).LegsSlot > 0 Then
                                slot = tempplayer.Char(q).LegsSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Legs=" & Item(itemnumber).Name
                            Else
                                output = output & "&Legs=None"
                            End If
                            If tempplayer.Char(q).RingSlot > 0 Then
                                slot = tempplayer.Char(q).RingSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Ring=" & Item(itemnumber).Name
                            Else
                                output = output & "&Ring=None"
                            End If
                            If tempplayer.Char(q).NecklaceSlot > 0 Then
                                slot = tempplayer.Char(q).NecklaceSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Necklace=" & Item(itemnumber).Name
                            Else
                                output = output & "&Necklace=None"
                            End If
                            output = output & "&Guild=" & tempplayer.Char(q).Guild
                            For x = 1 To 24
                                itemnumber = Int(tempplayer.Char(q).Inv(x).num)
                                If Item(itemnumber).Name = "" Then
                                    output = output & "&item" & x & "=None"
                                Else
                                    
                                    
                                    If Item(itemnumber).Type = ITEM_TYPE_CURRENCY Or Item(itemnumber).Stackable = 1 Then
                                        output = output & "&item" & x & "=" & Item(itemnumber).Name & "(" & tempplayer.Char(q).Inv(x).Value & ")"
                                    Else
                                        output = output & "&item" & x & "=" & Item(itemnumber).Name
                                    End If
                                End If
                            Next
                            Call SendDataTo(index, output)
                            Exit Sub
                        End If
                    Next q
                Else
                    i = FindPlayer(Name)
                    output = "name=" & GetPlayerName(i) & "&status=Online" & "&level=" & STR(GetPlayerLevel(i)) & "&class=" & GetVar(App.Path & "\classes\class" & player(i).Char(player(i).charnum).Class & ".ini", "CLASS", "Name") & " "
                    If player(i).Char(player(i).charnum).Sex = 0 Then
                        output = output & "&sex=Male "
                    Else
                        output = output & "&sex=Female "
                    End If
                    output = output & "&HP=" & STR(GetPlayerMaxHP(i)) & "&MP=" & STR(GetPlayerMaxMP(i)) & "&SP=" & STR(GetPlayerMaxSP(i)) & "&STR=" & STR(GetPlayerSTR(i)) & "&DEF=" & STR(GetPlayerDEF(i)) & "&MAGI=" & STR(GetPlayerMAGI(i)) & "&SPEED=" & STR(GetPlayerSPEED(i))
                    If GetPlayerWeaponSlot(i) > 0 Then
                        slot = GetPlayerWeaponSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Weapon=" & Item(itemnumber).Name
                    Else
                        output = output & "&Weapon=None"
                    End If
                    If GetPlayerArmorSlot(i) > 0 Then
                        slot = GetPlayerArmorSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Armor=" & Item(itemnumber).Name
                    Else
                        output = output & "&Armor=None"
                    End If
                    If GetPlayerHelmetSlot(i) > 0 Then
                        slot = GetPlayerHelmetSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Helmet=" & Item(itemnumber).Name
                    Else
                        output = output & "&Helmet=None"
                    End If
                    If GetPlayerShieldSlot(i) > 0 Then
                        slot = GetPlayerShieldSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Shield=" & Item(itemnumber).Name
                    Else
                        output = output & "&Shield=None"
                    End If
                    If GetPlayerLegsSlot(i) > 0 Then
                        slot = GetPlayerLegsSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Legs=" & Item(itemnumber).Name
                    Else
                        output = output & "&Legs=None"
                    End If
                    If GetPlayerRingSlot(i) > 0 Then
                        slot = GetPlayerRingSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Ring=" & Item(itemnumber).Name
                    Else
                        output = output & "&Ring=None"
                    End If
                    If GetPlayerNecklaceSlot(i) > 0 Then
                        slot = GetPlayerNecklaceSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Necklace=" & Item(itemnumber).Name
                    Else
                        output = output & "&Necklace=None"
                    End If
                    output = output & "&Guild=" & player(i).Char(player(i).charnum).Guild
                    For x = 1 To 24
                        itemnumber = Int(GetPlayerInvItemNum(i, x))
                        If Item(GetPlayerInvItemNum(i, x)).Name = "" Then
                        output = output & "&item" & x & "=None"
                            
                        Else
                          If Item(itemnumber).Type = ITEM_TYPE_CURRENCY Or Item(itemnumber).Stackable = 1 Then
                                output = output & "&item" & x & "=" & Item(itemnumber).Name & "(" & GetPlayerInvItemValue(i, x) & ")"
                            Else
                                output = output & "&item" & x & "=" & Item(itemnumber).Name
                            End If
                        End If
                    Next
                    Call SendDataTo(index, output)
                    Exit Sub
                End If
            Else
                output = "Could not find " & Name & "!"
                Call SendDataTo(index, output)
                Exit Sub
            End If
        End If
        
        If LCase(Mid(Buffer, 1, 14)) = "getplayerstats" Then
            Name = Mid(Buffer, 16)
            Name = Trim$(Name)
            If FindChar(Name) Then
                If FindPlayer(Name) = 0 Then
                    Call Findcharfile(Name)
                    For q = 1 To MAX_CHARS
                        If LCase(Trim$(tempplayer.Char(q).Name)) = LCase(Name) Then
                            output = "name=" & Trim$(tempplayer.Char(q).Name) & "&status=Offline" & "&level=" & STR(tempplayer.Char(q).Level) & "&class=" & GetVar(App.Path & "\classes\class" & tempplayer.Char(q).Class & ".ini", "CLASS", "Name") & " "
                            If tempplayer.Char(q).Sex = 0 Then
                                output = output & "&sex=Male"
                            Else
                                output = output & "&sex=Female"
                            End If
                            output = output & "&HP=" & STR(tempplayer.Char(q).MAXHP) & "&MP=" & STR(tempplayer.Char(q).MAXMP) & "&SP=" & STR(tempplayer.Char(q).MAXSP) & "&STR=" & STR(tempplayer.Char(q).STR) & "&DEF=" & STR(tempplayer.Char(q).DEF) & "&MAGI=" & STR(tempplayer.Char(q).Magi) & "&SPEED=" & STR(tempplayer.Char(q).Speed)
                            If tempplayer.Char(q).WeaponSlot > 0 Then
                                slot = tempplayer.Char(q).WeaponSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Weapon=" & Item(itemnumber).Name
                            Else
                                output = output & "&Weapon=None"
                            End If
                            If tempplayer.Char(q).ArmorSlot > 0 Then
                                slot = tempplayer.Char(q).ArmorSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Armor=" & Item(itemnumber).Name
                            Else
                                output = output & "&Armor=None"
                            End If
                            If tempplayer.Char(q).HelmetSlot > 0 Then
                                slot = tempplayer.Char(q).HelmetSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Helmet=" & Item(itemnumber).Name
                            Else
                                output = output & "&Helmet=None"
                            End If
                            If tempplayer.Char(q).ShieldSlot > 0 Then
                                slot = tempplayer.Char(q).ShieldSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Shield=" & Item(itemnumber).Name
                            Else
                                output = output & "&Shield=None"
                            End If
                            If tempplayer.Char(q).LegsSlot > 0 Then
                                slot = tempplayer.Char(q).LegsSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Legs=" & Item(itemnumber).Name
                            Else
                                output = output & "&Legs=None"
                            End If
                            If tempplayer.Char(q).RingSlot > 0 Then
                                slot = tempplayer.Char(q).RingSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Ring=" & Item(itemnumber).Name
                            Else
                                output = output & "&Ring=None"
                            End If
                            If tempplayer.Char(q).NecklaceSlot > 0 Then
                                slot = tempplayer.Char(q).NecklaceSlot
                                itemnumber = Int(tempplayer.Char(q).Inv(slot).num)
                                output = output & "&Necklace=" & Item(itemnumber).Name
                            Else
                                output = output & "&Necklace=None"
                            End If
                            output = output & "&Guild=" & tempplayer.Char(q).Guild
                            For x = 1 To 24
                                itemnumber = Int(tempplayer.Char(q).Inv(x).num)
                                If Item(itemnumber).Name = "" Then
                                    output = output & "&item" & x & "=None"
                                Else
                                    
                                    
                                    If Item(itemnumber).Type = ITEM_TYPE_CURRENCY Or Item(itemnumber).Stackable = 1 Then
                                        output = output & "&item" & x & "=" & Item(itemnumber).Name & "(" & tempplayer.Char(q).Inv(x).Value & ")"
                                    Else
                                        output = output & "&item" & x & "=" & Item(itemnumber).Name
                                    End If
                                End If
                            Next
                            Call SendDataTo(index, output)
                            Exit Sub
                        End If
                    Next q
                Else
                    i = FindPlayer(Name)
                    output = "name=" & GetPlayerName(i) & "&status=Online" & "&level=" & STR(GetPlayerLevel(i)) & "&class=" & GetVar(App.Path & "\classes\class" & player(i).Char(player(i).charnum).Class & ".ini", "CLASS", "Name") & " "
                    If player(i).Char(player(i).charnum).Sex = 0 Then
                        output = output & "&sex=Male "
                    Else
                        output = output & "&sex=Female "
                    End If
                    output = output & "&HP=" & STR(GetPlayerMaxHP(i)) & "&MP=" & STR(GetPlayerMaxMP(i)) & "&SP=" & STR(GetPlayerMaxSP(i)) & "&STR=" & STR(GetPlayerSTR(i)) & "&DEF=" & STR(GetPlayerDEF(i)) & "&MAGI=" & STR(GetPlayerMAGI(i)) & "&SPEED=" & STR(GetPlayerSPEED(i))
                    If GetPlayerWeaponSlot(i) > 0 Then
                        slot = GetPlayerWeaponSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Weapon=" & Item(itemnumber).Name
                    Else
                        output = output & "&Weapon=None"
                    End If
                    If GetPlayerArmorSlot(i) > 0 Then
                        slot = GetPlayerArmorSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Armor=" & Item(itemnumber).Name
                    Else
                        output = output & "&Armor=None"
                    End If
                    If GetPlayerHelmetSlot(i) > 0 Then
                        slot = GetPlayerHelmetSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Helmet=" & Item(itemnumber).Name
                    Else
                        output = output & "&Helmet=None"
                    End If
                    If GetPlayerShieldSlot(i) > 0 Then
                        slot = GetPlayerShieldSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Shield=" & Item(itemnumber).Name
                    Else
                        output = output & "&Shield=None"
                    End If
                    If GetPlayerLegsSlot(i) > 0 Then
                        slot = GetPlayerLegsSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Legs=" & Item(itemnumber).Name
                    Else
                        output = output & "&Legs=None"
                    End If
                    If GetPlayerRingSlot(i) > 0 Then
                        slot = GetPlayerRingSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Ring=" & Item(itemnumber).Name
                    Else
                        output = output & "&Ring=None"
                    End If
                    If GetPlayerNecklaceSlot(i) > 0 Then
                        slot = GetPlayerNecklaceSlot(i)
                        itemnumber = Int(GetPlayerInvItemNum(i, slot))
                        output = output & "&Necklace=" & Item(itemnumber).Name
                    Else
                        output = output & "&Necklace=None"
                    End If
                    output = output & "&Guild=" & player(i).Char(player(i).charnum).Guild
                    For x = 1 To 24
                        itemnumber = Int(GetPlayerInvItemNum(i, x))
                        If Item(GetPlayerInvItemNum(i, x)).Name = "" Then
                        output = output & "&item" & x & "=None"
                            
                        Else
                          If Item(itemnumber).Type = ITEM_TYPE_CURRENCY Or Item(itemnumber).Stackable = 1 Then
                                output = output & "&item" & x & "=" & Item(itemnumber).Name & "(" & GetPlayerInvItemValue(i, x) & ")"
                            Else
                                output = output & "&item" & x & "=" & Item(itemnumber).Name
                            End If
                        End If
                    Next
                    Call SendDataTo(index, output)
                    Exit Sub
                End If
            Else
                output = "Could not find " & Name & "!"
                Call SendDataTo(index, output)
                Exit Sub
            End If
        End If
        
        If Buffer = "getplayerlist" Then
            output = "chars[]="
            Open App.Path & "\accounts\charlist.txt" For Input As #43
            Do While Not EOF(43)
                Input #43, d
                output = output & d & "&chars[]="
            Loop
            Close #43
            Call SendDataTo(index, output)
            Exit Sub
        End If
        
        
        
        If LCase(Mid(Buffer, 1, 14)) = "newoaccountied" Then
            If Not IsLoggedIn(index) Then
                Parse = Split(Buffer, ",")
                Name = Parse(1)
                Password = Parse(2)
                Imail = Parse(3)
                
                For i = 1 To Len(Name)
                    n = Asc(Mid(Name, i, 1))
                    
                    If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                    Else
                        Call SendDataTo(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                        Exit Sub
                    End If
                Next i
                
                If Not AccountExist(Name) Then
                    Call AddAccount(index, Name, Password, Imail)
                    Call TextAdd(frmServer.txtText(0), "Account " & Name & " has been created.", True)
                    Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                    Call SendDataTo(index, "Your account has been created!")
                Else
                    Call SendDataTo(index, "Sorry, that account name is already taken!")
                End If
            End If
            Exit Sub
        End If
        
        If Buffer = "getonlineplayers" Then
            output = ""
            x = 0
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                     x = x + 1
                     output = output & Trim$(GetPlayerName(i)) & SEP_CHAR
                End If
            Next i
            output = x & SEP_CHAR & output
            Call SendDataTo(index, output)
            Exit Sub
        End If
            
        
End Sub
Sub HandleData(ByVal index As Long, ByVal Data As String)
    Dim Parse() As String
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim charnum As Long
    Dim Msg As String
    Dim MsgTo As Long
    Dim Dir As Long
    Dim InvNum As Long
    Dim Amount As Long
    Dim Damage As Long
    Dim PointType As Long
    Dim Movement As Long
    Dim i As Long, n As Long, x As Long, y As Long, f As Long
    Dim MapNum As Long
    Dim s As String
    Dim ShopNum As Long, ItemNum As Long
    Dim DurNeeded As Long, GoldNeeded As Long
    Dim z As Long
    Dim BX As Long, BY As Long
    Dim TempVal As Long
    Dim hfile
    Dim m As Long
    Dim j As Long
    Dim Imail As String
    Dim surge As String
    Dim Header As String
    Dim textbody As String
    Dim output As String
    
    
    
'    Handle the data
    Parse = Split(Data, SEP_CHAR)
    
'    Parse's Without Being Online
    If Not IsPlaying(index) Then
        Select Case LCase(Parse(0))
            
        Case "getonlineplayers"
            output = ""
            x = 0
            For i = 1 To MAX_PLAYERS
                If IsPlaying(1) Then
                     x = x + 1
                     output = output & player(i).Char(player(i).charnum).Name & SEP_CHAR
                End If
            Next i
            output = x + SEP_CHAR + output & END_CHAR
            Call SendDataTo(index, output)
            Exit Sub
            
        Case "setclevents"
            Header = "<div align='center'><b>" & Parse(1) & "</b><br><br>" & Parse(2) & "</div>"
            Call PutVar(App.Path & "\Events.ini", "DATA", "Events", Header)
            Exit Sub
            
        Case "gatglasses"
            Call SendNewCharClasses(index)
            Exit Sub
            
        Case "newfaccountied"
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
        Case "delimaccounted"
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
                    If Trim(player(index).Char(i).Name) <> "" Then
                        Call DeleteName(player(index).Char(i).Name)
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
            
        Case "logination"
            If Not IsLoggedIn(index) Then
                Name = Parse(1)
                Password = Parse(2)
                
                If ReadINI("CONFIG", "verified", App.Path & "\Data.ini") = 1 Then
                    If Val(ReadINI("GENERAL", "verified", App.Path & "\accounts\" & Trim(player(index).Login) & ".ini")) = 0 Then
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
                
'                Dim Packs As String
'                Packs = "MAXINFO" & SEP_CHAR
'                Packs = Packs & GAME_NAME & SEP_CHAR
'                Packs = Packs & MAX_PLAYERS & SEP_CHAR
'                Packs = Packs & MAX_ITEMS & SEP_CHAR
'                Packs = Packs & MAX_NPCS & SEP_CHAR
'                Packs = Packs & MAX_SHOPS & SEP_CHAR
'                Packs = Packs & MAX_SPELLS & SEP_CHAR
'                Packs = Packs & MAX_MAPS & SEP_CHAR
'                Packs = Packs & MAX_MAP_ITEMS & SEP_CHAR
'                Packs = Packs & MAX_MAPX & SEP_CHAR
'                Packs = Packs & MAX_MAPY & SEP_CHAR
'                Packs = Packs & MAX_EMOTICONS & SEP_CHAR
'                Packs = Packs & MAX_ELEMENTS & SEP_CHAR
'                Packs = Packs & PAPERDOLL & SEP_CHAR
'                Packs = Packs & SPRITESIZE & SEP_CHAR
'                Packs = Packs & MAX_SCRIPTSPELLS & SEP_CHAR
'                Packs = Packs & ENCRYPT_PASS & SEP_CHAR
'                Packs = Packs & ENCRYPT_TYPE & SEP_CHAR
'                Packs = Packs & END_CHAR
'                Call SendDataTo(index, Packs)
                
                Call LoadPlayer(index, Name)
                Call SendChars(index)
                
                Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
                Call TextAdd(frmServer.txtText(0), GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", True)
            End If
            Exit Sub
            
        Case "givemethemax"
            Dim packs As String
            packs = "MAXINFO" & SEP_CHAR
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
'            Send them the news too
            Call SendNewsTo(index)
            Exit Sub
            
        Case "addachara"
            Dim headc As Long
            Dim bodyc As Long
            Dim legc As Long
            
            Name = Parse(1)
            Sex = Val(Parse(2))
            Class = Val(Parse(3))
            charnum = Val(Parse(4))
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
            
            If charnum < 1 Or charnum > MAX_CHARS Then
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
            
            If CharExist(index, charnum) Then
                Call PlainMsg(index, "Character already exists!", 4)
                Exit Sub
            End If
            
            If FindChar(Name) Then
                Call PlainMsg(index, "Sorry, but that name is in use!", 4)
                Exit Sub
            End If
            
            Call AddChar(index, Name, Sex, Class, charnum, headc, bodyc, legc)
            Call SavePlayer(index)
            Call AddLog("Character " & Name & " added to " & GetPlayerLogin(index) & "'s account.", PLAYER_LOG)
            Call SendChars(index)
            Call PlainMsg(index, "Character has been created!", 5)
            
'            Dunno how useful this would be, but it's there if a future dev wants to work with it. -Pickle
            If Scripting = 1 Then
                Call MyScript.ExecuteStatement("Scripts\Main.txt", "OnNewChar " & index & "," & charnum)
            End If
            Exit Sub
            
        Case "delimbocharu"
            charnum = Val(Parse(1))
            
            If charnum < 1 Or charnum > MAX_CHARS Then
                Call HackingAttempt(index, "Invalid CharNum")
                Exit Sub
            End If
            
            Call DelChar(index, charnum)
            Call AddLog("Character deleted on " & GetPlayerLogin(index) & "'s account.", PLAYER_LOG)
            Call SendChars(index)
            Call PlainMsg(index, "Character has been deleted!", 5)
            Exit Sub
            
        Case "usagakarim"
            charnum = Val(Parse(1))
            
            If charnum < 1 Or charnum > MAX_CHARS Then
                Call HackingAttempt(index, "Invalid CharNum")
                Exit Sub
            End If
            
            If CharExist(index, charnum) Then
                player(index).charnum = charnum
                
                If frmServer.GMOnly.Value = Checked Then
                    If GetPlayerAccess(index) <= 0 Then
                        Call PlainMsg(index, "The server is only available to GMs at the moment!", 5)
'                        Call HackingAttempt(Index, "The server is only available to GMs at the moment!")
                        Exit Sub
                    End If
                End If
                
                Call JoinGame(index)
                
                charnum = player(index).charnum
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
            
        Case "mail"
            Call SendDataTo(index, "mail" & SEP_CHAR & Val(GetVar(App.Path & "\Data.ini", "CONFIG", "Email")) & SEP_CHAR & END_CHAR)
            Exit Sub
        End Select
    End If
    
'    Parse's With Being Online And Playing
    If IsPlaying(index) = False Then Exit Sub
    If IsConnected(index) = False Then Exit Sub
    Select Case LCase(Parse(0))
        
        
        
'        :::::::::::::::::::
'        :: Guilds Packet ::
'        :::::::::::::::::::
'        Access
    Case "guildchangeaccess"
        
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
'        Check the requirements.
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
        
'        Set the player's New access level
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), Parse(2))
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub
        
'        Disown
    Case "guilddisown"
'        Check if all the requirements
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
        
'        Player checks out, take him out of the guild
        Call setplayerguild(FindPlayer(Parse(1)), "")
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub
        
'        Leave Guild
    Case "guildleave"
'        Check if they can leave
        If GetPlayerGuild(index) = "" Then
            Call PlayerMsg(index, "You are not in a guild.", Red)
            Exit Sub
        End If
        
        Call setplayerguild(index, "")
        Call SetPlayerGuildAccess(index, 0)
        Call SendPlayerData(index)
        Exit Sub
        
'        Make A New Guild
    Case "makeguild"
'        Check if the Owner is Online
        If FindPlayer(Parse(1)) = 0 Then
            Call PlayerMsg(index, "Player is offline", White)
            Exit Sub
        End If
        
'        Check if they are alredy in a guild
        If GetPlayerGuild(FindPlayer(Parse(1))) <> "" Then
            Call PlayerMsg(index, "Player is already in a guild", Red)
            Exit Sub
        End If
        
'        If everything is ok then lets make the guild
        Call setplayerguild(FindPlayer(Parse(1)), (Parse(2)))
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 4)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub
        
'        Make A Member
    Case "guildmember"
'        Check if its possible to admit the member
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
        
'        All has gone well, set the guild access to 1
        Call setplayerguild(FindPlayer(Parse(1)), GetPlayerGuild(index))
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 1)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub
        
'        Make A Trainie
    Case "guildtrainee"
'        Check if its possible to induct member
        If FindPlayer(Parse(1)) = 0 Then
            Call PlayerMsg(index, "Player is offline", White)
            Exit Sub
        End If
        
        If GetPlayerGuild(FindPlayer(Parse(1))) <> "" Then
            Call PlayerMsg(index, "Player is already in a guild", Red)
            Exit Sub
        End If
        
'        It is possible, so set the guild to index's guild, and the access level to 0
        Call setplayerguild(FindPlayer(Parse(1)), GetPlayerGuild(index))
        Call SetPlayerGuildAccess(FindPlayer(Parse(1)), 0)
        Call SendPlayerData(FindPlayer(Parse(1)))
        Exit Sub
        
    Case "updatesite"
        Exit Sub
'        ::::::::::::::::::::
'        :: Social packets ::
'        ::::::::::::::::::::
    Case "saymsg"
        Msg = Parse(1)
        
'        Prevent hacking
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
        
'        ASGARD
'        Check for swearing
        Msg = SwearCheck(Msg)
        
        Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " : " & Msg & "", PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " : " & Msg & "", SayColor)
        Call MapMsg2(GetPlayerMap(index), Msg, index)
        TextAdd frmServer.txtText(3), GetPlayerName(index) & " On Map " & GetPlayerMap(index) & ": " & Msg, True
        Exit Sub
        
    Case "emotemsg"
        Msg = Parse(1)
        
'        Prevent hacking
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
        
    Case "broadcastmsg"
        Msg = Parse(1)
        
'        Prevent hacking
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
        
        If player(index).Mute = True Then Exit Sub
        
        Msg = SwearCheck(Msg)
        
        s = GetPlayerName(index) & ": " & Msg
        Call AddLog(s, PLAYER_LOG)
        Call GlobalMsg(s, BroadcastColor)
        Call TextAdd(frmServer.txtText(0), s, True)
        TextAdd frmServer.txtText(1), GetPlayerName(index) & ": " & Msg, True
        Exit Sub
        
    Case "globalmsg"
        Msg = Parse(1)
        
'        Prevent hacking
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
        
        If player(index).Mute = True Then Exit Sub
        
        Msg = SwearCheck(Msg)
        
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
        
'        Prevent hacking
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
        
    Case "playermsg"
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
        
'        Prevent hacking
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
        
'        Check if they are trying to talk to themselves
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
        
'        :::::::::::::::::::::::
'        :: edit main  packet ::
'        :::::::::::::::::::::::
        
    Case "editmain"
'        Prevent hacking
        
        If GetPlayerAccess(index) < ADMIN_MAPPER And 0 + Val(ReadINI(GetPlayerName(index), "editmain", App.Path & "\Acces.ini")) <> 1 Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        hfile = FreeFile
        Open App.Path & "\Scripts\Main.txt" For Input As #hfile
        frmEditor.RT.text = Input$(LOF(hfile), hfile)
        Close #hfile
        Call SendDataTo(index, "MAIN" & SEP_CHAR & frmEditor.RT.text & SEP_CHAR & END_CHAR)
        Exit Sub
        
'        :::::::::::::::::::::::
'        :: save main  packet ::
'        :::::::::::::::::::::::
        
    Case "savemain"
'        Prevent hacking
        
        If GetPlayerAccess(index) < ADMIN_MAPPER And 0 + Val(ReadINI(GetPlayerName(index), "editmain", App.Path & "\Acces.ini")) <> 1 Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        AFileName = "Scripts\Main.txt"
        Open App.Path & "\" & AFileName For Output As #1
        Print #1, Trim$(Parse$(1))
        Close #1
        Exit Sub
        
'        :::::::::::::::::::::::::::::
'        :: Moving character packet ::
'        :::::::::::::::::::::::::::::
    Case "playermove"
        If player(index).GettingMap = YES Then
            Exit Sub
        End If
        
        Dir = Val(Parse(1))
        Movement = Val(Parse(2))
        
'        Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(index, "Invalid Direction")
            Exit Sub
        End If
        
'        Prevent hacking
        If Movement < 1 Or Movement > 2 Then
            Call HackingAttempt(index, "Invalid Movement")
            Exit Sub
        End If
        
'        Prevent player from moving if they have casted a spell
        If player(index).CastedSpell = YES Then
'            Check if they have already casted a spell, and if so we can't let them move
            If GetTickCount > player(index).AttackTimer + 1000 Then
                player(index).CastedSpell = NO
            Else
                Call SendPlayerXY(index)
                Exit Sub
            End If
        End If
        
'        Prevent player from moving if they have been script locked
        If player(index).locked = True Then
            Call SendPlayerXY(index)
            Exit Sub
        End If
        
        Call PlayerMove(index, Dir, Movement)
        Exit Sub
        
'        :::::::::::::::::::::::::::::
'        :: Moving character packet ::
'        :::::::::::::::::::::::::::::
    Case "playerdir"
        If player(index).GettingMap = YES Then
            Exit Sub
        End If
        
        Dir = Val(Parse(1))
        
'        Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(index, "Invalid Direction")
            Exit Sub
        End If
        
        Call SetPlayerDir(index, Dir)
        Call SendDataToMapBut(index, GetPlayerMap(index), "PLAYERDIR" & SEP_CHAR & index & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & END_CHAR)
        Exit Sub
        
'        :::::::::::::::::::::
'        :: Use item packet ::
'        :::::::::::::::::::::
    Case "useitem"
        InvNum = Val(Parse(1))
        charnum = player(index).charnum
'        Prevent player from using an item when they are locked
        
        If player(index).lockeditems = True Then
            Exit Sub
        End If
        
'        Prevent hacking
        If InvNum < 1 Or InvNum > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid InvNum")
            Exit Sub
        End If
        
'        Prevent hacking
        If charnum < 1 Or charnum > MAX_CHARS Then
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
            
'            Find out what kind of item it is
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
                
            Case ITEM_TYPE_TWO_HAND
                
                If InvNum <> GetPlayerWeaponSlot(index) Then
                    If n4 > 0 Then
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
                        Call PlayerMsg(index, "Your strength is too low to equip this item!  Required str (" & n1 & ")", BrightRed)
                        Exit Sub
                    ElseIf Int(GetPlayerDEF(index)) < n2 Then
                        Call PlayerMsg(index, "Your defence is too low to equip this item!  Required Def (" & n2 & ")", BrightRed)
                        Exit Sub
                    ElseIf Int(GetPlayerSPEED(index)) < n3 Then
                        Call PlayerMsg(index, "Your speed is too low to equip this item!  Required Speed (" & n3 & ")", BrightRed)
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
                Call SetPlayerHP(index, GetPlayerHP(index) + Item(player(index).Char(charnum).Inv(InvNum).num).Data1)
                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 1)
                Else
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 0)
                End If
                Call SendHP(index)
                
            Case ITEM_TYPE_POTIONADDMP
                Call SetPlayerMP(index, GetPlayerMP(index) + Item(player(index).Char(charnum).Inv(InvNum).num).Data1)
                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 1)
                Else
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 0)
                End If
                Call SendMP(index)
                
            Case ITEM_TYPE_POTIONADDSP
                Call SetPlayerSP(index, GetPlayerSP(index) + Item(player(index).Char(charnum).Inv(InvNum).num).Data1)
                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 1)
                Else
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 0)
                End If
                Call SendSP(index)
                
            Case ITEM_TYPE_POTIONSUBHP
                Call SetPlayerHP(index, GetPlayerHP(index) - Item(player(index).Char(charnum).Inv(InvNum).num).Data1)
                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 1)
                Else
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 0)
                End If
                Call SendHP(index)
                
            Case ITEM_TYPE_POTIONSUBMP
                Call SetPlayerMP(index, GetPlayerMP(index) - Item(player(index).Char(charnum).Inv(InvNum).num).Data1)
                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 1)
                Else
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 0)
                End If
                Call SendMP(index)
                
            Case ITEM_TYPE_POTIONSUBSP
                Call SetPlayerSP(index, GetPlayerSP(index) - Item(player(index).Char(charnum).Inv(InvNum).num).Data1)
                If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 1)
                Else
                    Call TakeItem(index, player(index).Char(charnum).Inv(InvNum).num, 0)
                End If
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
                
'                Check if a key exists
                If Map(GetPlayerMap(index)).tile(x, y).Type = TILE_TYPE_KEY Then
'                    Check if the key they are using matches the map key
                    If GetPlayerInvItemNum(index, InvNum) = Map(GetPlayerMap(index)).tile(x, y).Data1 Then
                        TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                        TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                        
                        Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                        If Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String1) = "" Then
                            Call MapMsg(GetPlayerMap(index), "A door has been unlocked!", White)
                        Else
                            Call MapMsg(GetPlayerMap(index), Trim(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).String1), White)
                        End If
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
                        
'                        Check if we are supposed to take away the item
                        If Map(GetPlayerMap(index)).tile(x, y).Data2 = 1 Then
                            Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                            Call PlayerMsg(index, "The key disolves.", Yellow)
                        End If
                    End If
                End If
                
                If Map(GetPlayerMap(index)).tile(x, y).Type = TILE_TYPE_DOOR Then
                    TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                    TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
                    
                    Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & SEP_CHAR & END_CHAR)
                End If
                
            Case ITEM_TYPE_SPELL
'                Get the spell num
                n = Item(GetPlayerInvItemNum(index, InvNum)).Data1
                
                If n > 0 Then
'                    Make sure they are the right class
                    If Spell(n).ClassReq - 1 = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                        If Spell(n).LevelReq = 0 And player(index).Char(player(index).charnum).access < 1 Then
                            Call PlayerMsg(index, "This spell can only be used by admins!", BrightRed)
                            Exit Sub
                        End If
                        
'                        Make sure they are the right level
                        i = GetSpellReqLevel(n)
                        If i <= GetPlayerLevel(index) Then
                            i = FindOpenSpellSlot(index)
                            
'                            Make sure they have an open spell slot
                            If i > 0 Then
'                                Make sure they dont already have the spell
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
                MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedItem " & index & "," & Item(player(index).Char(charnum).Inv(InvNum).num).Data1
                
            End Select
            
            Call SendStats(index)
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
            
'            Send everyone player's equipment
            Call SendIndexWornEquipment(index)
        End If
'        End If
        Exit Sub
        
    Case "playermovemouse"
        
        If 0 + Val(ReadINI("CONFIG", "mouse", App.Path & "\Data.ini")) = 1 Then
            Call SendDataTo(index, "mouse" & SEP_CHAR & END_CHAR)
        End If
        
        
        If player(index).GettingMap = YES Then
            Exit Sub
        End If
        
        Dir = Val(Parse(1))
        Movement = 1
'        Prevent hacking
        
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(index, "Invalid Direction")
            Exit Sub
        End If
        
'        Prevent hacking
        
        If Movement < 1 Or Movement > 2 Then
            Call HackingAttempt(index, "Invalid Movement")
            Exit Sub
        End If
        
'        Prevent player from moving if they have casted a spell
        
        If player(index).CastedSpell = YES Then
'            Check if they have already casted a spell, and if so we can't let them move
            
            If GetTickCount > player(index).AttackTimer + 1000 Then
                player(index).CastedSpell = NO
            Else
                Call SendPlayerXY(index)
                Exit Sub
            End If
            
        End If
        
'        Prevent player from moving if they have been script locked
        
        If player(index).locked = True Then
            Call SendPlayerXY(index)
            Exit Sub
        End If
        
        Exit Sub
        
    Case "warp"
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
        
'        ::::::::::::::::::
'        :: graplepacket ::
'        ::::::::::::::::::
    Case "endshot"
        
        If Val(Parse(1)) = 0 Then
            player(index).locked = False
            Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
            player(index).HookShotX = 0
            player(index).HookShotY = 0
            Exit Sub
        End If
        
        If player(index).HookShotX = 0 Or player(index).HookShotY = 0 Then
            Call HackingAttempt(index, "")
        End If
        
        Call PlayerMsg(index, "You carefully cross the wire.", 1)
        player(index).locked = False
        
'        This doesn't work... you need to make changes to the game loop
'        to get an animation effect. This just hangs the server. -Pickle
'        If GetPlayerX(index) < Player(index).HookShotX Then
'        Do While GetPlayerX(index) < Player(index).HookShotX
'        Call SetPlayerX(index, GetPlayerX(index) + 1)
'        Call SetPlayerY(index, Player(index).HookShotY)
'        Loop
'        End If
        
'        If GetPlayerX(index) > Player(index).HookShotX Then
'        Do While GetPlayerX(index) > Player(index).HookShotX
'        Call SetPlayerX(index, GetPlayerX(index) - 1)
'        Call SetPlayerY(index, Player(index).HookShotY)
'        Loop
'        End If
        
'        If GetPlayerY(index) < Player(index).HookShotY Then
'        Do While GetPlayerX(index) < Player(index).HookShotY
'        Call SetPlayerY(index, GetPlayerY(index) + 1)
'        Call SetPlayerX(index, Player(index).HookShotX)
'        Loop
'        End If
        
'        If GetPlayerY(index) > Player(index).HookShotY Then
'        Do While GetPlayerX(index) > Player(index).HookShotY
'        Call SetPlayerY(index, GetPlayerY(index) - 1)
'        Call SetPlayerX(index, Player(index).HookShotX)
'        Loop
'        End If
        
'        Temporary fix. We'll see if we can do animation later... :) -Pickle
        Call SetPlayerX(index, player(index).HookShotX)
        Call SetPlayerY(index, player(index).HookShotY)
        player(index).HookShotX = 0
        player(index).HookShotY = 0
        Call SendPlayerData(index)
        Exit Sub
        
'        ::::::::::::::::::::::::::
'        :: Player attack packet ::
'        ::::::::::::::::::::::::::
    Case "attack"
'        Prevent player from moving if they have been script locked
        
        If player(index).lockedattack = True Then
            Exit Sub
        End If
        
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & index
        End If
        
        If GetPlayerWeaponSlot(index) > 0 Then
            If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 > 0 Then
                If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Stackable = 0 Then
                    Call SendDataToMap(GetPlayerMap(index), "checkarrows" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & END_CHAR)
                Else
                    Call GrapleHook(index)
                End If
                Exit Sub
            End If
        End If
        
'        Try to attack a player
        For i = 1 To MAX_PLAYERS
'            Make sure we dont try to attack ourselves
            If i <> index Then
'                Can we attack the player?
                If CanAttackPlayer(index, i) Then
                    If Not CanPlayerBlockHit(i) Then
'                        Get the damage we can do
                        If Not CanPlayerCriticalHit(index) Then
                            Damage = GetPlayerDamage(index) - GetPlayerProtection(i)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                        Else
                            n = GetPlayerDamage(index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                            surge = GetVar("Lang.ini", "Lang", "Surge")
                            If surge <> "" Then
                                Call BattleMsg(index, "You feel a surge of energy upon swinging!", BrightCyan, 0)
                                Call BattleMsg(i, GetPlayerName(index) & " swings with enormous might!", BrightCyan, 1)
                            End If
                            
'                            Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
'                            Call PlayerMsg(I, GetPlayerName(index) & " swings with enormous might!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
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
                        
'                        Call PlayerMsg(index, GetPlayerName(I) & "'s " & Trim(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked your hit!", BrightCyan)
'                        Call PlayerMsg(I, "Your " & Trim(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                    End If
                    
                    Exit Sub
                End If
            End If
        Next i
        
'        Try to attack a npc
        For i = 1 To MAX_MAP_NPCS
'            Can we attack the npc?
            If CanAttackNpc(index, i) Then
'                Get the damage we can do
                player(index).targetnpc = i
                If Not CanPlayerCriticalHit(index) Then
                    Damage = GetPlayerDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), i).num).DEF / 2)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                Else
                    n = GetPlayerDamage(index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).num).DEF / 2)
                    Call BattleMsg(index, "You feel a surge of energy upon swinging!", BrightCyan, 0)
                    
'                    Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(index, i, Damage)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                Else
                    Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)
                    
'                    Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                End If
                Exit Sub
            End If
        Next i
        
'        Check for skill
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
        
'        ::::::::::::::::::::::
'        :: Use stats packet ::
'        ::::::::::::::::::::::
    Case "usestatpoint"
        PointType = Val(Parse(1))
        
'        Prevent hacking
        If (PointType < 0) Or (PointType > 3) Then
            Call HackingAttempt(index, "Invalid Point Type")
            Exit Sub
        End If
        
'        Make sure they have points
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
        player(index).Char(player(index).charnum).MAXHP = GetPlayerMaxHP(index)
        player(index).Char(player(index).charnum).MAXMP = GetPlayerMaxMP(index)
        player(index).Char(player(index).charnum).MAXSP = GetPlayerMaxSP(index)
        Call SendStats(index)
        
        Call SendDataTo(index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(index) & SEP_CHAR & END_CHAR)
        Exit Sub
        
'        ::::::::::::::::::::::::::::::::
'        :: Player info request packet ::
'        ::::::::::::::::::::::::::::::::
    Case "playerinforequest"
        Name = Parse(1)
        
        i = FindPlayer(Name)
        If i > 0 Then
            Call PlayerMsg(index, "Account: " & Trim(player(i).Login) & ", Name: " & GetPlayerName(i), BrightGreen)
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
        
'        :::::::::::::::::::::::
'        :: Set sprite packet ::
'        :::::::::::::::::::::::
    Case "setsprite"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
'        The sprite
        n = Val(Parse(1))
        
        Call SetPlayerSprite(index, n)
        Call SendPlayerData(index)
        Exit Sub
        
'        ::::::::::::::::::::::::::::::
'        :: Set player sprite packet ::
'        ::::::::::::::::::::::::::::::
    Case "setplayersprite"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
'        The sprite
        i = FindPlayer(Parse(1))
        n = Val(Parse(2))
        
        Call SetPlayerSprite(i, n)
        Call SendPlayerData(i)
        Exit Sub
        
'        ::::::::::::::::::::::::::
'        :: Stats request packet ::
'        ::::::::::::::::::::::::::
    Case "getstats"
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
        
    Case "canonshoot"
        
        If HasItem(index, Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1) > 0 Then
            Call TakeItem(index, Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1, 1)
            Call SendDataToMap(GetPlayerMap(index), "scriptspellanim" & SEP_CHAR & Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3 & SEP_CHAR & Spell(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3).SpellAnim & SEP_CHAR & Spell(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3).SpellTime & SEP_CHAR & Spell(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3).SpellDone & SEP_CHAR & Val(Parse(1)) & SEP_CHAR & Val(Parse(2)) & SEP_CHAR & Spell(Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data3).Big & SEP_CHAR & END_CHAR)
            
            
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
        
        
'        ::::::::::::::::::::::::::::::::::
'        :: Player request for a new map ::
'        ::::::::::::::::::::::::::::::::::
    Case "requestnewmap"
        Dir = Val(Parse(1))
        
'        Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(index, "Invalid Direction")
            Exit Sub
        End If
        
        Call PlayerMove(index, Dir, 1)
        Exit Sub
        
'        :::::::::::::::::::::
'        :: Map data packet ::
'        :::::::::::::::::::::
    Case "mapdata"
'        Error Handling
        Err.Clear
        On Error Resume Next
'        Prevent hacking
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
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(MapNum).tile(x, y).Ground = Val(Parse(n))
                Map(MapNum).tile(x, y).Mask = Val(Parse(n + 1))
                Map(MapNum).tile(x, y).Anim = Val(Parse(n + 2))
                Map(MapNum).tile(x, y).Mask2 = Val(Parse(n + 3))
                Map(MapNum).tile(x, y).M2Anim = Val(Parse(n + 4))
                Map(MapNum).tile(x, y).Fringe = Val(Parse(n + 5))
                Map(MapNum).tile(x, y).FAnim = Val(Parse(n + 6))
                Map(MapNum).tile(x, y).Fringe2 = Val(Parse(n + 7))
                Map(MapNum).tile(x, y).F2Anim = Val(Parse(n + 8))
                Map(MapNum).tile(x, y).Type = Val(Parse(n + 9))
                Map(MapNum).tile(x, y).Data1 = Val(Parse(n + 10))
                Map(MapNum).tile(x, y).Data2 = Val(Parse(n + 11))
                Map(MapNum).tile(x, y).Data3 = Val(Parse(n + 12))
                Map(MapNum).tile(x, y).String1 = Parse(n + 13)
                Map(MapNum).tile(x, y).String2 = Parse(n + 14)
                Map(MapNum).tile(x, y).String3 = Parse(n + 15)
                Map(MapNum).tile(x, y).light = Val(Parse(n + 16))
                Map(MapNum).tile(x, y).GroundSet = Val(Parse(n + 17))
                Map(MapNum).tile(x, y).MaskSet = Val(Parse(n + 18))
                Map(MapNum).tile(x, y).AnimSet = Val(Parse(n + 19))
                Map(MapNum).tile(x, y).Mask2Set = Val(Parse(n + 20))
                Map(MapNum).tile(x, y).M2AnimSet = Val(Parse(n + 21))
                Map(MapNum).tile(x, y).FringeSet = Val(Parse(n + 22))
                Map(MapNum).tile(x, y).FAnimSet = Val(Parse(n + 23))
                Map(MapNum).tile(x, y).Fringe2Set = Val(Parse(n + 24))
                Map(MapNum).tile(x, y).F2AnimSet = Val(Parse(n + 25))
                
                
                n = n + 26
            Next x
        Next y
        
        For x = 1 To MAX_MAP_NPCS
            Map(MapNum).Npc(x) = Val(Parse(n))
            n = n + 1
            Call ClearMapNpc(x, MapNum)
        Next x
        
'        Clear out it all
        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
            Call ClearMapItem(i, GetPlayerMap(index))
        Next i
        
'        Save the map
        Call SaveMap(MapNum)
        
'        Respawn
        Call SpawnMapItems(GetPlayerMap(index))
        
'        Respawn NPCS
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, GetPlayerMap(index))
        Next i
        
        
'        Refresh map for everyone online
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & SEP_CHAR & END_CHAR)
'                Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
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
        
'        Reset error handling
        On Error GoTo 0
        
'        ::::::::::::::::::::::::::::
'        :: Need map yes/no packet ::
'        ::::::::::::::::::::::::::::
    Case "needmap"
'        Get yes/no value
        s = LCase(Parse(1))
        
        If s = "yes" Then
            Call SendMap(index, GetPlayerMap(index))
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            player(index).GettingMap = NO
            Call SendDataTo(index, "MAPDONE" & SEP_CHAR & END_CHAR)
        Else
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            player(index).GettingMap = NO
            Call SendDataTo(index, "MAPDONE" & SEP_CHAR & END_CHAR)
        End If
        
'        Send everyone player's equipment
        
'        Call SendIndexWornEquipmentFromMap(index)
'        Call SendWornEquipment(index)
'        Call GetMapWornEquipment(index)
        Dim iindex As Long
        For iindex = 1 To MAX_PLAYERS
            If IsPlaying(iindex) Then
                If iindex <> index Then
'                    Call JoinWarp(iindex, GetPlayerMap(iindex), GetPlayerX(iindex), GetPlayerY(iindex))
                End If
                Call SendIndexWornEquipment(iindex)
                Call SendWornEquipment(iindex)
            End If
        Next iindex
        
        Exit Sub
        
    Case "mapdone2"
        
        Exit Sub
    Case "needmapnum2"
        Call SendMap(index, GetPlayerMap(index))
        Exit Sub
        
'        :::::::::::::::::::::::::::::::::::::::::::::::
'        :: Player trying to pick up something packet ::
'        :::::::::::::::::::::::::::::::::::::::::::::::
    Case "mapgetitem"
        Call PlayerMapGetItem(index)
        Exit Sub
    
    Case "getplayerhp"
        Call SendDataTo(index, "playerhpreturn" & SEP_CHAR & Val(Parse(1)) & SEP_CHAR & Val(player(Parse(1)).Char(player(Parse(1)).charnum).HP) & SEP_CHAR & GetPlayerMaxHP(Parse(1)) & SEP_CHAR & END_CHAR)
        Exit Sub
'        ::::::::::::::::::::::::::::::::::::::::::::
'        :: Player trying to drop something packet ::
'        ::::::::::::::::::::::::::::::::::::::::::::
    Case "mapdropitem"
        InvNum = Val(Parse(1))
        Amount = Val(Parse(2))
        
'        Prevent hacking
        If InvNum < 1 Or InvNum > MAX_INV Then
            Call HackingAttempt(index, "Invalid InvNum")
            Exit Sub
        End If
        
        
'        Prevent hacking
        If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
'            Check if money and if it is we want to make sure that they aren't trying to drop 0 value
            If Amount <= 0 Then
                Call PlayerMsg(index, "You must drop more than 0!", BrightRed)
                Exit Sub
            End If
            
            If Amount > GetPlayerInvItemValue(index, InvNum) Then
                Call PlayerMsg(index, "You dont have that much to drop!", BrightRed)
                Exit Sub
            End If
        End If
        
'        Prevent hacking
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
        
'        ::::::::::::::::::::::::
'        :: Respawn map packet ::
'        ::::::::::::::::::::::::
    Case "maprespawn"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
'        Clear out it all
        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
            Call ClearMapItem(i, GetPlayerMap(index))
        Next i
        
'        Respawn
        Call SpawnMapItems(GetPlayerMap(index))
        
'        Respawn NPCS
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, GetPlayerMap(index))
        Next i
        
        
        Call PlayerMsg(index, "Map respawned.", Blue)
        Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)
        Exit Sub
        
        
'        ::::::::::::::::::::::::
'        :: Kick player packet ::
'        ::::::::::::::::::::::::
    Case "kickplayer"
'        Prevent hacking
        If GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
'        The player index
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
        
'        :::::::::::::::::::::
'        :: Ban list packet ::
'        :::::::::::::::::::::
    Case "banlist"
'        Prevent hacking
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
        
'        ::::::::::::::::::::::::
'        :: Ban destroy packet ::
'        ::::::::::::::::::::::::
    Case "bandestroy"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_CREATOR Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call Kill(App.Path & "\banlist.txt")
        Call PlayerMsg(index, "Ban list destroyed.", White)
        Exit Sub
        
'        :::::::::::::::::::::::
'        :: Ban player packet ::
'        :::::::::::::::::::::::
    Case "banplayer"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
'        The player index
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
        
'        :::::::::::::::::::::::::::::
'        :: Request edit map packet ::
'        :::::::::::::::::::::::::::::
    Case "requesteditmap"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_MAPPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "EDITMAP" & SEP_CHAR & END_CHAR)
        Exit Sub
'        :::::::::::::::::::::::::::::
'        :: Request edit house packet ::
'        :::::::::::::::::::::::::::::
    Case "requestedithouse"
'        Prevent hacking
        If Map(GetPlayerMap(index)).Moral <> MAP_MORAL_HOUSE Then
            Call PlayerMsg(index, "This is not a house!", BrightRed)
            Exit Sub
        End If
        If Map(GetPlayerMap(index)).owner <> GetPlayerName(index) Then
            Call PlayerMsg(index, "This is not your house!", BrightRed)
            Exit Sub
        End If
        
        
        
        Call SendDataTo(index, "EDITHOUSE" & SEP_CHAR & END_CHAR)
        Exit Sub
'        ::::::::::::::::::::::::::::::
'        :: Request edit item packet ::
'        ::::::::::::::::::::::::::::::
    Case "requestedititem"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "ITEMEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
        
'        ::::::::::::::::::::::
'        :: Edit item packet ::
'        ::::::::::::::::::::::
    Case "edititem"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
'        The item #
        n = Val(Parse(1))
        
'        Prevent hacking
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid Item Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing item #" & n & ".", ADMIN_LOG)
        Call SendEditItemTo(index, n)
        Exit Sub
        
'        ::::::::::::::::::::::
'        :: Save item packet ::
'        ::::::::::::::::::::::
    Case "saveitem"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(index, "Invalid Item Index")
            Exit Sub
        End If
        
'        Update the item
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
        
        Item(n).addHP = Val(Parse(13))
        Item(n).addMP = Val(Parse(14))
        Item(n).addSP = Val(Parse(15))
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
        
'        Save it
        Call SendUpdateItemToAll(n)
        Call SaveItem(n)
        Call AddLog(GetPlayerName(index) & " saved item #" & n & ".", ADMIN_LOG)
        Exit Sub
        
'        :::::::::::::::::::::
'        :: Day/Night Stuff ::
'        :::::::::::::::::::::
        
    Case "enabledaynight"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        If TimeDisable = False Then
            Gamespeed = 0
            frmServer.GameTimeSpeed.text = 0
            TimeDisable = True
            frmServer.Timer1.Enabled = False
            frmServer.Command69.caption = "Enable Time"
        Else
            Gamespeed = 1
            frmServer.GameTimeSpeed.text = 1
            TimeDisable = False
            frmServer.Timer1.Enabled = True
            frmServer.Command69.caption = "Disable Time"
        End If
        
        Exit Sub
        
    Case "daynight"
'        Prevent hacking
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
        
'        :::::::::::::::::::::::::::::
'        :: Request edit npc packet ::
'        :::::::::::::::::::::::::::::
    Case "requesteditnpc"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "NPCEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
        
'        :::::::::::::::::::::
'        :: Edit npc packet ::
'        :::::::::::::::::::::
    Case "editnpc"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
'        The npc #
        n = Val(Parse(1))
        
'        Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(index, "Invalid NPC Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing npc #" & n & ".", ADMIN_LOG)
        Call SendEditNpcTo(index, n)
        Exit Sub
        
'        :::::::::::::::::::::
'        :: Save npc packet ::
'        :::::::::::::::::::::
    Case "savenpc"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        
'        Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(index, "Invalid NPC Index")
            Exit Sub
        End If
        
'        Update the npc
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
        Npc(n).MAXHP = Val(Parse(13))
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
        
'        Save it
        Call SendUpdateNpcToAll(n)
        Call SaveNpc(n)
        Call AddLog(GetPlayerName(index) & " saved npc #" & n & ".", ADMIN_LOG)
        Exit Sub
        
'        ::::::::::::::::::::::::::::::
'        :: Request edit shop packet ::
'        ::::::::::::::::::::::::::::::
    Case "requesteditshop"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "SHOPEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
        
'        ::::::::::::::::::::::
'        :: Edit shop packet ::
'        ::::::::::::::::::::::
    Case "editshop"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
'        The shop #
        n = Val(Parse(1))
        
'        Prevent hacking
        If n < 0 Or n > MAX_SHOPS Then
            Call HackingAttempt(index, "Invalid Shop Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing shop #" & n & ".", ADMIN_LOG)
        Call SendEditShopTo(index, n)
        Exit Sub
        
'        ::::::::::::::::::::::
'        :: Save shop packet ::
'        ::::::::::::::::::::::
    Case "saveshop"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        ShopNum = Val(Parse(1))
        
'        Prevent hacking
        If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
            Call HackingAttempt(index, "Invalid Shop Index")
            Exit Sub
        End If
        
'        Update the shop
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
        
        
'        Save it
        Call SendUpdateShopToAll(ShopNum)
        Call SaveShop(ShopNum)
        Call AddLog(GetPlayerName(index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
        Exit Sub
        
'        :::::::::::::::::::::::::::::::
'        :: Request edit spell packet ::
'        :::::::::::::::::::::::::::::::
    Case "requesteditspell"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "SPELLEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
        
'        :::::::::::::::::::::::
'        :: Edit spell packet ::
'        :::::::::::::::::::::::
    Case "editspell"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
'        The spell #
        n = Val(Parse(1))
        
'        Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(index, "Invalid Spell Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(index) & " editing spell #" & n & ".", ADMIN_LOG)
        Call SendEditSpellTo(index, n)
        Exit Sub
        
'        :::::::::::::::::::::::
'        :: Save spell packet ::
'        :::::::::::::::::::::::
    Case "savespell"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
'        Spell #
        n = Val(Parse(1))
        
'        Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(index, "Invalid Spell Index")
            Exit Sub
        End If
        
'        Update the spell
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
        
'        Save it
        Call SendUpdateSpellToAll(n)
        Call SaveSpell(n)
        Call AddLog(GetPlayerName(index) & " saving spell #" & n & ".", ADMIN_LOG)
        Exit Sub
        
    Case "forgetspell"
'        Spell slot
        n = CLng(Parse(1))
        
'        Prevent subscript out of range
        If n <= 0 Or n > MAX_PLAYER_SPELLS Then
            Call HackingAttempt(index, "Invalid Spell Slot")
            Exit Sub
        End If
        
        With player(index).Char(player(index).charnum)
            If .Spell(n) = 0 Then
                Call PlayerMsg(index, "No spell here.", Red)
                
            Else
                Call PlayerMsg(index, "You have forgotten the spell """ & Trim$(Spell(.Spell(n)).Name) & """", Green)
                
                .Spell(n) = 0
                Call SendSpells(index)
            End If
        End With
        Exit Sub
        
'        :::::::::::::::::::::::
'        :: keypressed packet ::
'        :::::::::::::::::::::::
        
    Case "key"
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "Keys " & index & "," & Val(Parse$(1))
        End If
        Exit Sub
        
'        :::::::::::::::::::::::
'        :: Set access packet ::
'        :::::::::::::::::::::::
    Case "setaccess"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_CREATOR Then
            Call HackingAttempt(index, "Trying to use powers not available")
            Exit Sub
        End If
        
'        The index
        n = FindPlayer(Parse(1))
'        The access
        i = Val(Parse(2))
        
        
'        Check for invalid access level
        If i >= 0 Or i <= 3 Then
            If GetPlayerName(index) <> GetPlayerName(n) Then
                If GetPlayerAccess(index) > GetPlayerAccess(n) Then
'                    Check if player is on
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
                    Call PlayerMsg(index, "Your access level is lower than " & GetPlayerName(n) & "´.", Red)
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
        
        Call PutVar(App.Path & "\motd.ini", "MOTD", "Msg", Parse(1))
        Call GlobalMsg("MOTD changed to: " & Parse(1), BrightCyan)
        Call AddLog(GetPlayerName(index) & " changed MOTD to: " & Parse(1), ADMIN_LOG)
        Exit Sub
        
        
'        ********************
'        * Buy Item Packet  *
'        ********************
    Case "buy"
        Dim shopIndex As Integer
        Dim shopItemIndex As Integer
        
'        The number of the shop
        shopIndex = Val(Parse(1))
'        The number of the shop's item
        shopItemIndex = Val(Parse(2))
        
'        Error handling
        If shopIndex < 1 Or shopIndex > MAX_SHOPS Or shopItemIndex < 1 Or shopItemIndex > MAX_SHOP_ITEMS Then Exit Sub
        
'        Check to see if player's inventory is full
'        x is temp var
        x = FindOpenInvSlot(index, Shop(shopIndex).ShopItem(shopItemIndex).ItemNum)
        If x = 0 Then
            Call PlayerMsg(index, GetVar("Lang.ini", "Lang", "FullInv"), Red)
            Exit Sub
        End If
        
'        Check to see if they have enough currency
        If HasItem(index, Shop(shopIndex).currencyItem) >= Shop(shopIndex).ShopItem(shopItemIndex).Price Then
'            Buy the item
            TakeItem index, Shop(shopIndex).currencyItem, Shop(shopIndex).ShopItem(shopItemIndex).Price
            GiveItem index, Shop(shopIndex).ShopItem(shopItemIndex).ItemNum, Shop(shopIndex).ShopItem(shopItemIndex).Amount
            
'            Display message, check if it's a stackable item
            Call PlayerMsg(index, "You buy the item(s).", Yellow)
        Else
'            Can't trade
            Call PlayerMsg(index, "You can't afford that!", Red)
        End If
        Exit Sub
        
'        SELL ITEM PACKET
    Case "sellitem"
        Dim SellItemNum As Long
        Dim SellItemSlot As Integer
        Dim SellItemAmt As Integer
        Dim snumber As Integer
        
        snumber = Val(Parse(1))
        SellItemNum = Val(Parse(2))
        SellItemSlot = Val(Parse(3))
        SellItemAmt = Val(Parse(4))
        
        If GetPlayerWeaponSlot(index) = Val(Parse(2)) Or GetPlayerArmorSlot(index) = Val(Parse(2)) Or GetPlayerShieldSlot(index) = Val(Parse(2)) Or GetPlayerHelmetSlot(index) = Val(Parse(2)) Or GetPlayerLegsSlot(index) = Val(Parse(2)) Or GetPlayerRingSlot(index) = Val(Parse(2)) Or GetPlayerNecklaceSlot(index) = Val(Parse(2)) Then
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
            Call GiveItem(index, Shop(snumber).currencyItem, Item(SellItemNum).Price * SellItemAmt)
            Call PlayerMsg(index, "The shopkeeper hands you " & Item(SellItemNum).Price * SellItemAmt & " " & Trim$(Item(Shop(snumber).currencyItem).Name) & ".", Yellow)
        Else
            Call PlayerMsg(index, "This item can't be sold.", Red)
        End If
        Exit Sub
        
        
'        FIX ITEM PACKET
    Case "fixitem"
'        Inv num
        n = Val(Parse(1))
        
'        Make sure its a equipable item
        If Item(GetPlayerInvItemNum(index, n)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(index, n)).Type > ITEM_TYPE_NECKLACE Then
            Call PlayerMsg(index, "That item doesn't need to be fixed.", BrightRed)
            Exit Sub
        End If
        
'        Check if they have a full inventory
        If FindOpenInvSlot(index, GetPlayerInvItemNum(index, n)) <= 0 Then
            Call PlayerMsg(index, "You have no inventory space left!", BrightRed)
            Exit Sub
        End If
        
'        Now check the rate of pay
        ItemNum = GetPlayerInvItemNum(index, n)
        i = Int(Item(GetPlayerInvItemNum(index, n)).Data2 / 5)
        If i <= 0 Then i = 1
        
        DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(index, n)
        GoldNeeded = Int(DurNeeded * i / 2)
        If GoldNeeded <= 0 Then GoldNeeded = 1
        
'        Check if they even need it repaired
        If DurNeeded <= 0 Then
            Call PlayerMsg(index, "This item is in perfect condition!", White)
            Exit Sub
        End If
        
'        Check if they have enough for at least one point
        If HasItem(index, 1) >= i Then
'            Check if they have enough for a total restoration
            If HasItem(index, 1) >= GoldNeeded Then
                Call TakeItem(index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(index, n, Item(ItemNum).Data1)
                Call PlayerMsg(index, "Item has been totally restored for " & GoldNeeded & " gold!", BrightBlue)
            Else
'                They dont so restore as much as we can
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
        
    Case "search"
        x = Val(Parse(1))
        y = Val(Parse(2))
        
'        Prevent subscript out of range
        If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then
            Exit Sub
        End If
        
'        Check for a player
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(index) = GetPlayerMap(i) And GetPlayerX(i) = x And GetPlayerY(i) = y Then
                
'                Consider the player
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
                
'                Change target
                player(index).Target = i
                player(index).TargetType = TARGET_TYPE_PLAYER
                Call PlayerMsg(index, "Your target is now " & GetPlayerName(i) & ".", Yellow)
                Exit Sub
            End If
        Next i
        
'        Check for an npc
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(index), i).num > 0 Then
                If MapNpc(GetPlayerMap(index), i).x = x And MapNpc(GetPlayerMap(index), i).y = y Then
'                    Change target
                    player(index).Target = i
                    player(index).TargetType = TARGET_TYPE_NPC
                    Call PlayerMsg(index, "Your target is now a " & Trim(Npc(MapNpc(GetPlayerMap(index), i).num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next i
        
'        Check for a onClick Tile
        If Map(GetPlayerMap(index)).tile(x, y).Type = TILE_TYPE_ONCLICK Then
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnClick " & index & "," & Map(GetPlayerMap(index)).tile(x, y).Data1
            End If
        End If
        
        BX = x
        BY = y
        
'        Check for an item
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(GetPlayerMap(index), i).num > 0 Then
                If MapItem(GetPlayerMap(index), i).x = x And MapItem(GetPlayerMap(index), i).y = y Then
                    Call PlayerMsg(index, "You see a " & Trim(Item(MapItem(GetPlayerMap(index), i).num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next i
        
        
        Exit Sub
        
    Case "playerchat"
        n = FindPlayer(Parse(1))
        If n < 1 Then
            Call PlayerMsg(index, "Player is not online.", White)
            Exit Sub
        End If
        If n = index Then
            Exit Sub
        End If
        If player(index).InChat = 1 Then
            Call PlayerMsg(index, "Your already in a chat with another player!", Pink)
            Exit Sub
        End If
        
        If player(n).InChat = 1 Then
            Call PlayerMsg(index, "Player is already in a chat with another player!", Pink)
            Exit Sub
        End If
        
        If Parse(1) = "" Then
            Call PlayerMsg(index, "Click on the player you wish to chat to first.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Chat request has been sent to " & GetPlayerName(n) & ".", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " wants you to chat with them.  Type /chat to accept, or /chatdecline to decline.", Pink)
        
        player(n).ChatPlayer = index
        player(index).ChatPlayer = n
        Exit Sub
        
    Case "achat"
        n = player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If
        If player(n).ChatPlayer <> index Then
            Call PlayerMsg(index, "Chat failed.", Pink)
            Exit Sub
        End If
        
        Call SendDataTo(index, "PPCHATTING" & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "PPCHATTING" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "dchat"
        n = player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Declined chat request.", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " declined your request.", Pink)
        
        player(index).ChatPlayer = 0
        player(index).InChat = 0
        player(n).ChatPlayer = 0
        player(n).InChat = 0
        Exit Sub
        
    Case "qchat"
        n = player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If
        
        Call SendDataTo(index, "qchat" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "qchat" & SEP_CHAR & END_CHAR)
        
        player(index).ChatPlayer = 0
        player(index).InChat = 0
        player(n).ChatPlayer = 0
        player(n).InChat = 0
        Exit Sub
        
    Case "sendchat"
        n = player(index).ChatPlayer
        
        If n < 1 Then
            Call PlayerMsg(index, "No one requested to chat with you.", Pink)
            Exit Sub
        End If
        
        Call SendDataTo(n, "sendchat" & SEP_CHAR & Parse(1) & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "pptrade"
        n = FindPlayer(Parse(1))
        
'        Check if player is online
        If n < 1 Then
            Call PlayerMsg(index, "Player is not online.", White)
            Exit Sub
        End If
        
'        Prevent trading with self
        If n = index Then
            Exit Sub
        End If
        
'        Check if the player is in another trade
        If player(index).InTrade = 1 Then
            Call PlayerMsg(index, "Your already in a trade with someone else!", Pink)
            Exit Sub
        End If
        
'        Check where both players are
        Dim CanTrade As Boolean
        CanTrade = False
        
        If GetPlayerX(index) = GetPlayerX(n) And GetPlayerY(index) + 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) = GetPlayerX(n) And GetPlayerY(index) - 1 = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) + 1 = GetPlayerX(n) And GetPlayerY(index) = GetPlayerY(n) Then CanTrade = True
        If GetPlayerX(index) - 1 = GetPlayerX(n) And GetPlayerY(index) = GetPlayerY(n) Then CanTrade = True
        
        If CanTrade = True Then
'            Check to see if player is already in a trade
            If player(n).InTrade = 1 Then
                Call PlayerMsg(index, "Player is already in a trade!", Pink)
                Exit Sub
            End If
            
            Call PlayerMsg(index, "Trade request has been sent to " & GetPlayerName(n) & ".", Pink)
            Call PlayerMsg(n, GetPlayerName(index) & " wants you to trade with them.  Type /accept to accept, or /decline to decline.", Pink)
            
            player(n).TradePlayer = index
            player(index).TradePlayer = n
        Else
            Call PlayerMsg(index, "You need to be beside the player to trade!", Pink)
            Call PlayerMsg(n, "The player needs to be beside you to trade!", Pink)
        End If
        Exit Sub
        
    Case "atrade"
        n = player(index).TradePlayer
        
'        Check if anyone requested a trade
        If n < 1 Then
            Call PlayerMsg(index, "No one requested a trade with you.", Pink)
            Exit Sub
        End If
        
'        Check if its the right player
        If player(n).TradePlayer <> index Then
            Call PlayerMsg(index, "Trade failed.", Pink)
            Exit Sub
        End If
        
'        Check where both players are
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
                player(index).Trading(i).InvNum = 0
                player(index).Trading(i).InvName = ""
                player(n).Trading(i).InvNum = 0
                player(n).Trading(i).InvName = ""
            Next i
            
            player(index).InTrade = 1
            player(index).TradeItemMax = 0
            player(index).TradeItemMax2 = 0
            player(n).InTrade = 1
            player(n).TradeItemMax = 0
            player(n).TradeItemMax2 = 0
        Else
            Call PlayerMsg(index, "The player needs to be beside you to trade!", Pink)
            Call PlayerMsg(n, "You need to be beside the player to trade!", Pink)
        End If
        Exit Sub
        
    Case "qtrade"
        n = player(index).TradePlayer
        
'        Check if anyone trade with player
        If n < 1 Then
            Call PlayerMsg(index, "No one requested a trade with you.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Stopped trading.", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " stopped trading with you!", Pink)
        
        player(index).TradeOk = 0
        player(n).TradeOk = 0
        player(index).TradePlayer = 0
        player(index).InTrade = 0
        player(n).TradePlayer = 0
        player(n).InTrade = 0
        Call SendDataTo(index, "qtrade" & SEP_CHAR & END_CHAR)
        Call SendDataTo(n, "qtrade" & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "dtrade"
        n = player(index).TradePlayer
        
'        Check if anyone trade with player
        If n < 1 Then
            Call PlayerMsg(index, "No one requested a trade with you.", Pink)
            Exit Sub
        End If
        
        Call PlayerMsg(index, "Declined trade request.", Pink)
        Call PlayerMsg(n, GetPlayerName(index) & " declined your request.", Pink)
        
        player(index).TradePlayer = 0
        player(index).InTrade = 0
        player(n).TradePlayer = 0
        player(n).InTrade = 0
        Exit Sub
        
    Case "updatetradeinv"
        n = Val(Parse(1))
        
        player(index).Trading(n).InvNum = Val(Parse(2))
        player(index).Trading(n).InvName = Trim(Parse(3))
        If player(index).Trading(n).InvNum = 0 Then
            player(index).TradeItemMax = player(index).TradeItemMax - 1
            player(index).TradeOk = 0
            player(n).TradeOk = 0
            Call SendDataTo(index, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        Else
            player(index).TradeItemMax = player(index).TradeItemMax + 1
        End If
        
        Call SendDataTo(player(index).TradePlayer, "updatetradeitem" & SEP_CHAR & n & SEP_CHAR & player(index).Trading(n).InvNum & SEP_CHAR & player(index).Trading(n).InvName & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "swapitems"
        n = player(index).TradePlayer
        
        If player(index).TradeOk = 0 Then
            player(index).TradeOk = 1
            Call SendDataTo(n, "trading" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
        ElseIf player(index).TradeOk = 1 Then
            player(index).TradeOk = 0
            Call SendDataTo(n, "trading" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        End If
        
        If player(index).TradeOk = 1 And player(n).TradeOk = 1 Then
            player(index).TradeItemMax2 = 0
            player(n).TradeItemMax2 = 0
            
            For i = 1 To MAX_INV
                If player(index).TradeItemMax = player(index).TradeItemMax2 Then
                    Exit For
                End If
                If GetPlayerInvItemNum(n, i) < 1 Then
                    player(index).TradeItemMax2 = player(index).TradeItemMax2 + 1
                End If
            Next i
            
            For i = 1 To MAX_INV
                If player(n).TradeItemMax = player(n).TradeItemMax2 Then
                    Exit For
                End If
                If GetPlayerInvItemNum(index, i) < 1 Then
                    player(n).TradeItemMax2 = player(n).TradeItemMax2 + 1
                End If
            Next i
            
            If player(index).TradeItemMax2 = player(index).TradeItemMax And player(n).TradeItemMax2 = player(n).TradeItemMax Then
                For i = 1 To MAX_PLAYER_TRADES
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(n, x) < 1 Then
                            If player(index).Trading(i).InvNum > 0 Then
                                Call GiveItem(n, GetPlayerInvItemNum(index, player(index).Trading(i).InvNum), 1)
                                Call TakeItem(index, GetPlayerInvItemNum(index, player(index).Trading(i).InvNum), 1)
                                Exit For
                            End If
                        End If
                    Next x
                Next i
                
                For i = 1 To MAX_PLAYER_TRADES
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(index, x) < 1 Then
                            If player(n).Trading(i).InvNum > 0 Then
                                Call GiveItem(index, GetPlayerInvItemNum(n, player(n).Trading(i).InvNum), 1)
                                Call TakeItem(n, GetPlayerInvItemNum(n, player(n).Trading(i).InvNum), 1)
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
                If player(index).TradeItemMax2 < player(index).TradeItemMax Then
                    Call PlayerMsg(index, "Your inventory is full!", BrightRed)
                    Call PlayerMsg(n, GetPlayerName(index) & "'s inventory is full!", BrightRed)
                End If
                If player(n).TradeItemMax2 < player(n).TradeItemMax Then
                    Call PlayerMsg(n, "Your inventory is full!", BrightRed)
                    Call PlayerMsg(index, GetPlayerName(n) & "'s inventory is full!", BrightRed)
                End If
            End If
            
            player(index).TradePlayer = 0
            player(index).InTrade = 0
            player(index).TradeOk = 0
            player(n).TradePlayer = 0
            player(n).InTrade = 0
            player(n).TradeOk = 0
            Call SendDataTo(index, "qtrade" & SEP_CHAR & END_CHAR)
            Call SendDataTo(n, "qtrade" & SEP_CHAR & END_CHAR)
        End If
        Exit Sub
        
    Case "party"
        n = FindPlayer(Parse(1))
        
'        Prevent partying with self
        If n = index Then
            Exit Sub
        End If
        
'        Check for a full party and if so drop it
        Dim g As Integer
        g = 0
        If player(index).InParty = True Then
            For i = 1 To MAX_PARTY_MEMBERS
                If player(index).Party.Member(i) > 0 Then g = g + 1
            Next i
            If g > (MAX_PARTY_MEMBERS - 1) Then
                Call PlayerMsg(index, "Party is full!", Pink)
                Exit Sub
            End If
        End If
        
        If n > 0 Then
'            Check if its an admin
            If GetPlayerAccess(index) > ADMIN_MONITER Then
                Call PlayerMsg(index, "You can't join a party, you are an admin!", BrightBlue)
                Exit Sub
            End If
            
            If GetPlayerAccess(n) > ADMIN_MONITER Then
                Call PlayerMsg(index, "Admins cannot join parties!", BrightBlue)
                Exit Sub
            End If
            
'            Check to see if player is already in a party
            If player(n).InParty = False Then
                Call PlayerMsg(index, GetPlayerName(n) & " has been invited to your party.", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " has invited you to join their party.  Type /join to join, or /leave to decline.", Pink)
                
                player(n).InvitedBy = index
            Else
                Call PlayerMsg(index, "Player is already in a party!", Pink)
            End If
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If
        Exit Sub
        
    Case "joinparty"
        n = player(index).InvitedBy
        
        If n > 0 Then
'            Check to make sure they aren't the starter
'            Check to make sure that each of there party players match
            Call PlayerMsg(index, "You have joined " & GetPlayerName(n) & "'s party!", Pink)
            
            If player(n).InParty = False Then ' Set the party leader up
                Call SetPMember(n, n) 'Make them the first member and make them the leader
                player(n).InParty = True 'Set them to be 'InParty' status
                Call SetPShare(n, True)
            End If
            
            player(index).InParty = True 'Player joined
            player(index).Party.Leader = n 'Set party leader
            Call SetPMember(n, index) 'Add the member and update the party
            
'            Make sure they are in right level range
            If GetPlayerLevel(index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(index) - 5 > GetPlayerLevel(n) Then
                Call PlayerMsg(index, "There is more then a 5 level gap between you two, you will not share experience.", Pink)
                Call PlayerMsg(n, "There is more then a 5 level gap between you two, you will not share experience.", Pink)
                Call SetPShare(index, False) 'Do not share experience with party
            Else
                Call SetPShare(index, True) 'Share experience with party
            End If
            
            For i = 1 To MAX_PARTY_MEMBERS
                If player(index).Party.Member(i) > 0 And player(index).Party.Member(i) <> index Then Call PlayerMsg(player(index).Party.Member(i), GetPlayerName(index) & " has joined your party!", Pink)
            Next i
            
            For i = 1 To MAX_PARTY_MEMBERS
                If player(index).Party.Member(i) = index Then
                    For n = 1 To MAX_PARTY_MEMBERS
                        Call SendDataTo(n, "updatemembers" & SEP_CHAR & i & SEP_CHAR & index & SEP_CHAR & END_CHAR)
                    Next n
                End If
            Next i
            
            For i = 1 To MAX_PARTY_MEMBERS
                Call SendDataTo(index, "updatemembers" & SEP_CHAR & i & SEP_CHAR & player(index).Party.Member(i) & SEP_CHAR & END_CHAR)
            Next i
            
        Else
            Call PlayerMsg(index, "You have not been invited into a party!", Pink)
        End If
        Exit Sub
        
    Case "leaveparty"
        n = player(index).InvitedBy
        
        If n > 0 Or player(index).Party.Leader = index Then
            If player(index).InParty = True Then
                Call PlayerMsg(index, "You have left the party.", Pink)
                For i = 1 To MAX_PARTY_MEMBERS
                    If player(index).Party.Member(i) > 0 Then Call PlayerMsg(player(index).Party.Member(i), GetPlayerName(index) & " has left the party.", Pink)
                Next i
                
                Call RemovePMember(index) 'this handles removing them and updating the entire party
                Call SendDataTo(index, "leaveparty211")
            Else
                Call PlayerMsg(index, "Declined party request.", Pink)
                Call PlayerMsg(n, GetPlayerName(index) & " declined your request.", Pink)
                
                player(index).InParty = False
                player(index).InvitedBy = 0
                
            End If
        Else
            Call PlayerMsg(index, "You are not in a party!", Pink)
        End If
        Exit Sub
        
    Case "partychat"
        For i = 1 To MAX_PARTY_MEMBERS
            If player(index).Party.Member(i) > 0 Then Call PlayerMsg(player(index).Party.Member(i), Parse(1), Blue)
        Next i
        Exit Sub
        
    Case "spells"
        Call SendPlayerSpells(index)
        Exit Sub
        
    Case "hotscript1"
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
        
    Case "scripttile"
        Call SendDataTo(index, "SCRIPTTILE" & SEP_CHAR & GetVar(App.Path & "\Tiles.ini", "Names", "Tile" & Parse(1)) & SEP_CHAR & END_CHAR)
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
        Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
'        Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Exit Sub
        
    Case "buysprite"
'        Check if player stepped on sprite changing tile
        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
            Call PlayerMsg(index, "You need to be on a sprite tile to buy it!", BrightRed)
            Exit Sub
        End If
        
        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data2 = 0 Then
            Call SetPlayerSprite(index, Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1)
            Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
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
                        Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
                        Call SendInventory(index)
                    End If
                Else
                    If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerNecklaceSlot(index) <> i Then
                        Call SetPlayerInvItemNum(index, i, 0)
                        Call PlayerMsg(index, "You have bought a new sprite!", BrightGreen)
                        Call SetPlayerSprite(index, Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                        Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & END_CHAR)
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
        
    Case "clearowner"
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        Call PlayerMsg(index, "Owner cleared!", BrightRed)
        Map(GetPlayerMap(index)).owner = 0
        Map(GetPlayerMap(index)).Name = "Unonwed House"
        Map(GetPlayerMap(index)).Revision = Map(GetPlayerMap(index)).Revision + 1
        Call SaveMap(GetPlayerMap(index))
        Call SendDataToMap(GetPlayerMap(index), "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Revision + 1) & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "buyhouse"
'        Check if player stepped on house changing tile
        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_HOUSE Then
            Call PlayerMsg(index, "You need to be on a house tile to buy it!", BrightRed)
            Exit Sub
        End If
        
        If Map(GetPlayerMap(index)).tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 0 Then
            Map(GetPlayerMap(index)).owner = GetPlayerName(index)
            Map(GetPlayerMap(index)).Name = GetPlayerName(index) & "'s House"
            Map(GetPlayerMap(index)).Revision = Map(GetPlayerMap(index)).Revision + 1
            Call SaveMap(GetPlayerMap(index))
            Call SendDataToMap(GetPlayerMap(index), "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Revision + 1) & SEP_CHAR & END_CHAR)
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
                        Call SendDataToMap(GetPlayerMap(index), "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Revision + 1) & SEP_CHAR & END_CHAR)
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
                        Call SendDataToMap(GetPlayerMap(index), "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Revision + 1) & SEP_CHAR & END_CHAR)
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
        
    Case "checkcommands"
        s = Parse(1)
        If Scripting = 1 Then
            PutVar App.Path & "\Scripts\Command.ini", "TEMP", "Text" & index, Trim(s)
            MyScript.ExecuteStatement "Scripts\Main.txt", "Commands " & index
        Else
            Call PlayerMsg(index, "Thats not a valid command!", 12)
        End If
        Exit Sub
        
    Case "prompt"
        If Scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerPrompt " & index & "," & Val(Parse(1)) & "," & Val(Parse(2))
        End If
        Exit Sub
        
    Case "querybox"
        If Scripting = 1 Then
            Call PutVar(App.Path & "\responses.ini", "Responses", CStr(index), Parse(1))
            MyScript.ExecuteStatement "Scripts\Main.txt", "QueryBox " & index & "," & Val(Parse(2))
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
        Arrows(n).Amount = Val(Parse(5))
        
        Call SendUpdateArrowToAll(n)
        Call SaveArrow(n)
        Call AddLog(GetPlayerName(index) & " saved arrow #" & n & ".", ADMIN_LOG)
        Exit Sub
    Case "checkarrows"
        n = Arrows(Val(Parse(1))).Pic
        
        Call SendDataToMap(GetPlayerMap(index), "checkarrows" & SEP_CHAR & index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "requesteditemoticon"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "EMOTICONEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "endshot"
'        Prevent hacking
        If player(index).HookShotX <> 0 Or player(index).HookShotY <> 0 Then
            Call HackingAttempt(index, "")
            Exit Sub
        End If
        
        Call SetPlayerX(index, player(index).HookShotX)
        Call SetPlayerY(index, player(index).HookShotY)
        Call PlayerMsg(index, "You use your " & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Name & " to carefully cross over to the other side.", 3)
        Exit Sub
        
    Case "requesteditelement"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "ELEMENTEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "requesteditskill"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "SKILLEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "requesteditquest"
'        Prevent hacking
        If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(index, "QUESTEDITOR" & SEP_CHAR & END_CHAR)
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
        
    Case "editelement"
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
        
    Case "editskill"
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
        
    Case "editquest"
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
        
    Case "saveemoticon"
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
        
    Case "saveskill"
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
        
    Case "savequest"
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
            Quest(n).x(j) = Val(Parse(m + 1))
            Quest(n).y(j) = Val(Parse(m + 2))
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
        
    Case "saveelement"
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
        
        
    Case "checkemoticons"
        n = Emoticons(Val(Parse(1))).Pic
        
        Call SendDataToMap(GetPlayerMap(index), "checkemoticons" & SEP_CHAR & index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        Exit Sub
        
    Case "mapreport"
        packs = "mapreport" & SEP_CHAR
        For i = 1 To MAX_MAPS
            packs = packs & Map(i).Name & SEP_CHAR
        Next i
        packs = packs & END_CHAR
        
        Call SendDataTo(index, packs)
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
'            Make sure we dont try To attack ourselves
            If z <> index Then
'                Can we attack the player?
                If CanAttackPlayerWithArrow(index, z) Then
                    If Not CanPlayerBlockHit(z) Then
'                        Get the damage we can Do
                        If Not CanPlayerCriticalHit(index) Then
                            Damage = GetPlayerDamage(index) - GetPlayerProtection(z)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                        Else
                            n = GetPlayerDamage(index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(z)
                            Call BattleMsg(index, "You feel a surge of energy upon shooting!", BrightCyan, 0)
                            Call BattleMsg(z, GetPlayerName(index) & " shoots With amazing accuracy!", BrightCyan, 1)
                            
'                            Call PlayerMsg(index, "You feel a surge of energy upon shooting!", BrightCyan)
'                            Call PlayerMsg(z, GetPlayerName(index) & " shoots With amazing accuracy!", BrightCyan)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(index, z, Damage)
                        Else
                            Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)
                            Call BattleMsg(z, GetPlayerName(index) & "'s attack did nothing.", BrightRed, 1)
                            
'                            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call BattleMsg(index, GetPlayerName(z) & " blocked your hit!", BrightCyan, 0)
                        Call BattleMsg(z, "You blocked " & GetPlayerName(index) & "'s hit!", BrightCyan, 1)
                        
'                        Call PlayerMsg(index, GetPlayerName(z) & "'s " & Trim(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked your hit!", BrightCyan)
'                        Call PlayerMsg(z, "Your " & Trim(Item(GetPlayerInvItemNum(z, GetPlayerShieldSlot(z))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & SEP_CHAR & END_CHAR)
                    End If
                    Exit Sub
                End If
            End If
        ElseIf n = TARGET_TYPE_NPC Then
'            Can we attack the npc?
            If CanAttackNpcWithArrow(index, z) Then
'                Get the damage we can Do
                If Not CanPlayerCriticalHit(index) Then
                    Damage = GetPlayerDamage(index) - Int(Npc(MapNpc(GetPlayerMap(index), z).num).DEF / 2)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                Else
                    n = GetPlayerDamage(index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), z).num).DEF / 2)
                    Call BattleMsg(index, "You feel a surge of energy upon shooting!", BrightCyan, 0)
                    
'                    Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(index, z, Damage)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
                Else
                    Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)
                    
'                    Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
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
                Call BattleMsg(index, "You feel a surge of energy upon shooting!", BrightCyan, 0)
                
'                Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & SEP_CHAR & END_CHAR)
            End If
            
            If Damage > 0 Then
                Call AttackAttributeNpc(z, BX, BY, index, Damage)
'                Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
            Else
                Call BattleMsg(index, "Your attack does nothing.", BrightRed, 0)
                
'                Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
'                Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & END_CHAR)
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
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Bank full!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        If Val(Parse(2)) > GetPlayerInvItemValue(index, Val(Parse(1))) Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You cant deposit more than you have!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        If GetPlayerWeaponSlot(index) = Val(Parse(1)) Or GetPlayerArmorSlot(index) = Val(Parse(1)) Or GetPlayerShieldSlot(index) = Val(Parse(1)) Or GetPlayerHelmetSlot(index) = Val(Parse(1)) Or GetPlayerLegsSlot(index) = Val(Parse(1)) Or GetPlayerRingSlot(index) = Val(Parse(1)) Or GetPlayerNecklaceSlot(index) = Val(Parse(1)) Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You cant deposit worn equipment!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        If Item(x).Type = ITEM_TYPE_CURRENCY Or Item(x).Stackable = 1 Then
            If Val(Parse(2)) <= 0 Then
                Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You must deposit more than 0!" & SEP_CHAR & END_CHAR)
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
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Inventory full!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        If Val(Parse(2)) > GetPlayerBankItemValue(index, Val(Parse(1))) Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You cant withdraw more than you have!" & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        If Item(i).Type = ITEM_TYPE_CURRENCY Or Item(i).Stackable = 1 Then
            If Val(Parse(2)) <= 0 Then
                Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You must withdraw more than 0!" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
        
        Call GiveItem(index, i, TempVal)
        Call TakeBankItem(index, i, TempVal)
        
        Call SendBank(index)
        Exit Sub
        
'        Reload the scripts
    Case "reloadscripts"
        
        Set MyScript = Nothing
        Set clsScriptCommands = Nothing
        
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
        Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
        
        Exit Sub
        
    Case "custommenuclick"
        
        player(index).custom_title = Parse(3)
        player(index).custom_msg = Parse(5)
        
        If Scripting = 1 Then
'            MyScript.ExecuteStatement "Scripts\Main.txt", "CustomMenu " & Parse(1) & "," & Parse(2) & "," & Parse(3) & "," & Parse(4) & "," & Parse(5)
            MyScript.ExecuteStatement "Scripts\Main.txt", "menuscripts " & Parse(1) & "," & Parse(2) & "," & Parse(4)
        End If
        
        Exit Sub
        
    Case "returningcustomboxmsg"
        
        player(index).custom_msg = Parse(1)
        
        Exit Sub
        
    End Select
    Call HackingAttempt(index, "")
    Exit Sub
End Sub

Sub CloseSocket(ByVal index As Long)
'    Make sure player was/is playing the game, and if so, save'm.
    If index > 0 Then
        Call LeftGame(index)
        
        Call TextAdd(frmServer.txtText(0), "Connection from " & GetPlayerIP(index) & " has been terminated.", True)
        
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
        Packet = Packet & Trim(player(index).Char(i).Name) & SEP_CHAR & Trim(Class(player(index).Char(i).Class).Name) & SEP_CHAR & player(index).Char(i).Level & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendJoinMap(ByVal index As Long)
    Dim Packet As String
    Dim i As Long
    Dim j As Long
    
    Packet = ""
    
'    Send all players on current map to index
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
            Packet = Packet & GetPlayerHead(i) & SEP_CHAR
            Packet = Packet & GetPlayerBody(i) & SEP_CHAR
            Packet = Packet & GetPlayerleg(i) & SEP_CHAR
            Packet = Packet & GetPlayerPaperdoll(i) & SEP_CHAR
            Packet = Packet & GetPlayerLevel(i) & SEP_CHAR
            For j = 1 To MAX_SKILLS
                Packet = Packet & GetPlayerSkillLvl(j, i) & SEP_CHAR
                Packet = Packet & GetPlayerSkillExp(j, i) & SEP_CHAR
            Next j
            Packet = Packet & END_CHAR
            Call SendDataTo(index, Packet)
        End If
    Next i
    
'    Send index's player data to everyone on the map including himself
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
    Packet = Packet & GetPlayerHead(index) & SEP_CHAR
    Packet = Packet & GetPlayerBody(index) & SEP_CHAR
    Packet = Packet & GetPlayerleg(index) & SEP_CHAR
    Packet = Packet & GetPlayerPaperdoll(index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(index) & SEP_CHAR
    For j = 1 To MAX_SKILLS
        Packet = Packet & GetPlayerSkillLvl(j, index) & SEP_CHAR
        Packet = Packet & GetPlayerSkillExp(j, index) & SEP_CHAR
    Next j
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
    Dim Packet As String
    Dim j As Long
    
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
    Packet = Packet & GetPlayerHead(index) & SEP_CHAR
    Packet = Packet & GetPlayerBody(index) & SEP_CHAR
    Packet = Packet & GetPlayerleg(index) & SEP_CHAR
    Packet = Packet & GetPlayerPaperdoll(index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(index) & SEP_CHAR
    For j = 1 To MAX_SKILLS
        Packet = Packet & GetPlayerSkillLvl(j, index) & SEP_CHAR
        Packet = Packet & GetPlayerSkillExp(j, index) & SEP_CHAR
    Next j
    Packet = Packet & END_CHAR
    Call SendDataToMapBut(index, MapNum, Packet)
End Sub

Sub SendPlayerData(ByVal index As Long)
    Dim Packet As String
    Dim j As Long
    
'    Send index's player data to everyone including himself on th emap
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
    Packet = Packet & GetPlayerHead(index) & SEP_CHAR
    Packet = Packet & GetPlayerBody(index) & SEP_CHAR
    Packet = Packet & GetPlayerleg(index) & SEP_CHAR
    Packet = Packet & GetPlayerPaperdoll(index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(index) & SEP_CHAR
    For j = 1 To MAX_SKILLS
        Packet = Packet & GetPlayerSkillLvl(j, index) & SEP_CHAR
        Packet = Packet & GetPlayerSkillExp(j, index) & SEP_CHAR
    Next j
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub


Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
    Dim Packet As String, P1 As String, P2 As String
    Dim x As Integer
    Dim y As Integer
    
    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR & Map(MapNum).Weather & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map(MapNum).tile(x, y)
                Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .light & SEP_CHAR
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
    Dim i As Long
    
    For i = 1 To MAX_ITEMS
        If Trim(Item(i).Name) <> "" Then
            Call SendUpdateItemTo(index, i)
        End If
    Next i
End Sub

Sub SendSkills(ByVal index As Long)
    Dim i As Long
    
    For i = 1 To MAX_SKILLS
        Call SendUpdateSkillTo(index, i)
    Next i
End Sub

Sub SendQuests(ByVal index As Long)
    Dim i As Long
    
    For i = 1 To MAX_QUESTS
        Call SendUpdateQuestTo(index, i)
    Next i
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

Sub SendArrows(ByVal index As Long)
    Dim i As Long
    
    For i = 1 To MAX_ARROWS
        Call SendUpdateArrowTo(index, i)
    Next i
End Sub

Sub SendNpcs(ByVal index As Long)
    Dim i As Long
    
    For i = 1 To MAX_NPCS
        If Trim(Npc(i).Name) <> "" Then
            Call SendUpdateNpcTo(index, i)
        End If
    Next i
End Sub
Sub SendBank(ByVal index As Long)
    Dim Packet As String
    Dim i As Integer
    
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
        Packet = "PLAYERWORNEQ" & SEP_CHAR & index & SEP_CHAR & player(index).Char(player(index).charnum).ArmorSlot & SEP_CHAR & player(index).Char(player(index).charnum).WeaponSlot & SEP_CHAR & player(index).Char(player(index).charnum).HelmetSlot & SEP_CHAR & player(index).Char(player(index).charnum).ShieldSlot & SEP_CHAR & player(index).Char(player(index).charnum).LegsSlot & SEP_CHAR & player(index).Char(player(index).charnum).RingSlot & SEP_CHAR & player(index).Char(player(index).charnum).NecklaceSlot & SEP_CHAR & END_CHAR
        Call SendDataToMap(GetPlayerMap(index), Packet)
    End If
End Sub

Sub GetMapWornEquipment(ByVal index As Long)
    Dim Packet As String
    Dim i As Long
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If player(i).Char(player(i).charnum).Map = player(index).Char(player(index).charnum).Map Then
                Packet = "PLAYERWORNEQ" & SEP_CHAR & i & SEP_CHAR & player(i).Char(player(i).charnum).ArmorSlot & SEP_CHAR & player(i).Char(player(i).charnum).WeaponSlot & SEP_CHAR & player(i).Char(player(i).charnum).HelmetSlot & SEP_CHAR & player(i).Char(player(i).charnum).ShieldSlot & SEP_CHAR & player(i).Char(player(i).charnum).LegsSlot & SEP_CHAR & player(i).Char(player(i).charnum).RingSlot & SEP_CHAR & player(i).Char(player(i).charnum).NecklaceSlot & SEP_CHAR & END_CHAR
                Call SendDataTo(index, Packet)
            End If
        End If
    Next i
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

Sub SendPlayerLevelToAll(ByVal index As Long)
    Dim Packet As String
    
    Packet = "PLAYERLEVEL" & SEP_CHAR & index & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendClasses(ByVal index As Long)
    Dim Packet As String
    Dim i As Long
    
    Packet = "CLASSESDATA" & SEP_CHAR & MAX_CLASSES & SEP_CHAR
    For i = 0 To MAX_CLASSES
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).locked & SEP_CHAR & Class(i).Desc & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendNewCharClasses(ByVal index As Long)
    Dim Packet As String
    Dim i As Long
    
    Packet = "NEWCHARCLASSES" & SEP_CHAR & MAX_CLASSES & SEP_CHAR & ClassesOn & SEP_CHAR
    For i = 0 To MAX_CLASSES
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).Speed & SEP_CHAR & Class(i).Magi & SEP_CHAR & Class(i).MaleSprite & SEP_CHAR & Class(i).FemaleSprite & SEP_CHAR & Class(i).locked & SEP_CHAR & Class(i).Desc & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendLeftGame(ByVal index As Long)
    Dim Packet As String
    Dim j As Long
    
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
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    For j = 1 To MAX_SKILLS
        Packet = Packet & 0 & SEP_CHAR
        Packet = Packet & 0 & SEP_CHAR
    Next j
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
    
'    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).addHP & SEP_CHAR & Item(ItemNum).addMP & SEP_CHAR & Item(ItemNum).addSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
    Dim Packet As String
    
'    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).addHP & SEP_CHAR & Item(ItemNum).addMP & SEP_CHAR & Item(ItemNum).addSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditItemTo(ByVal index As Long, ByVal ItemNum As Long)
    Dim Packet As String
    
    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).addHP & SEP_CHAR & Item(ItemNum).addMP & SEP_CHAR & Item(ItemNum).addSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
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

Sub SendUpdateQuestToAll(ByVal QuestNum As Long)
    Dim Packet As String
    Dim j As Long
    
    Packet = "UPDATEQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim$(Quest(QuestNum).Name) & SEP_CHAR & Val(Quest(QuestNum).Pictop) & SEP_CHAR & Val(Quest(QuestNum).Picleft)
    
    For j = 0 To MAX_QUEST_LENGHT
        Packet = Packet & SEP_CHAR & Quest(QuestNum).Map(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).x(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).y(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).Npc(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).Script(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake1num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake2num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake1val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake2val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive1num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive2num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive1val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive2val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ExpGiven(j)
    Next j
    
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateQuestTo(ByVal index As Long, ByVal QuestNum As Long)
    Dim Packet As String
    Dim j As Long
    
    Packet = "UPDATEQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim$(Quest(QuestNum).Name) & SEP_CHAR & Val(Quest(QuestNum).Pictop) & SEP_CHAR & Val(Quest(QuestNum).Picleft)
    
    For j = 0 To MAX_QUEST_LENGHT
        Packet = Packet & SEP_CHAR & Quest(QuestNum).Map(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).x(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).y(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).Npc(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).Script(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake1num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake2num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake1val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake2val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive1num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive2num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive1val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive2val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ExpGiven(j)
    Next j
    
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditQuestTo(ByVal index As Long, ByVal QuestNum As Long)
    Dim Packet As String
    Dim j As Long
    
    Packet = "EDITQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim$(Quest(QuestNum).Name) & SEP_CHAR & Val(Quest(QuestNum).Pictop) & SEP_CHAR & Val(Quest(QuestNum).Picleft)
    
    For j = 0 To MAX_QUEST_LENGHT
        Packet = Packet & SEP_CHAR & Quest(QuestNum).Map(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).x(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).y(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).Npc(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).Script(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake1num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake2num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake1val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemTake2val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive1num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive2num(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive1val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ItemGive2val(j)
        Packet = Packet & SEP_CHAR & Quest(QuestNum).ExpGiven(j)
    Next j
    
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateSkillToAll(ByVal skillNum As Long)
    Dim Packet As String
    Dim j As Long
    
    Packet = "UPDATESKILL" & SEP_CHAR & skillNum & SEP_CHAR & Trim$(skill(skillNum).Name) & SEP_CHAR & Trim$(skill(skillNum).Action) & SEP_CHAR & Trim$(skill(skillNum).Fail) & SEP_CHAR & Trim$(skill(skillNum).Succes) & SEP_CHAR & Trim$(skill(skillNum).AttemptName) & SEP_CHAR & Val(skill(skillNum).Pictop) & SEP_CHAR & Val(skill(skillNum).Picleft)
    
    For j = 1 To MAX_SKILLS_SHEETS
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake1num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake2num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive1num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive2num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).minlevel(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ExpGiven(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).base_chance(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake1val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake2val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive1val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive2val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).itemequiped(j)
    Next j
    
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSkillTo(ByVal index As Long, ByVal skillNum As Long)
    Dim Packet As String
    Dim j As Long
    
    Packet = "UPDATESKILL" & SEP_CHAR & skillNum & SEP_CHAR & Trim$(skill(skillNum).Name) & SEP_CHAR & Trim$(skill(skillNum).Action) & SEP_CHAR & Trim$(skill(skillNum).Fail) & SEP_CHAR & Trim$(skill(skillNum).Succes) & SEP_CHAR & Trim$(skill(skillNum).AttemptName) & SEP_CHAR & Val(skill(skillNum).Pictop) & SEP_CHAR & Val(skill(skillNum).Picleft)
    
    For j = 1 To MAX_SKILLS_SHEETS
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake1num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake2num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive1num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive2num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).minlevel(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ExpGiven(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).base_chance(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake1val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake2val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive1val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive2val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).itemequiped(j)
    Next j
    
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdatePlayerSkill(ByVal index As Integer, ByVal skillNum As Integer)
    Dim Packet As String
    
    Packet = "SKILLINFO" & SEP_CHAR & skillNum & SEP_CHAR & GetPlayerSkillExp(index, skillNum) & SEP_CHAR & GetPlayerSkillLvl(index, skillNum) & SEP_CHAR & END_CHAR
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

Sub SendEditSkillTo(ByVal index As Long, ByVal skillNum As Long)
    Dim Packet As String
    Dim j As Long
    
    Packet = "EDITSKILL" & SEP_CHAR & skillNum & SEP_CHAR & Trim$(skill(skillNum).Name) & SEP_CHAR & Trim$(skill(skillNum).Action) & SEP_CHAR & Trim$(skill(skillNum).Fail) & SEP_CHAR & Trim$(skill(skillNum).Succes) & SEP_CHAR & Trim$(skill(skillNum).AttemptName) & SEP_CHAR & Val(skill(skillNum).Pictop) & SEP_CHAR & Val(skill(skillNum).Picleft)
    
    For j = 1 To MAX_SKILLS_SHEETS
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake1num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake2num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive1num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive2num(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).minlevel(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ExpGiven(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).base_chance(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake1val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemTake2val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive1val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).ItemGive2val(j)
        Packet = Packet & SEP_CHAR & skill(skillNum).itemequiped(j)
    Next j
    
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
    Dim Packet As String
    
    Packet = "UPDATEArrow" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateArrowTo(ByVal index As Long, ByVal ItemNum As Long)
    Dim Packet As String
    
    Packet = "UPDATEArrow" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditArrowTo(ByVal index As Long, ByVal EmoNum As Long)
    Dim Packet As String
    
    Packet = "EDITArrow" & SEP_CHAR & EmoNum & SEP_CHAR & Trim(Arrows(EmoNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
    Dim Packet As String
    
    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Spritesize & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MAXHP & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NpcNum As Long)
    Dim Packet As String
    
    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Spritesize & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MAXHP & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditNpcTo(ByVal index As Long, ByVal NpcNum As Long)
    Dim Packet As String
    Dim i As Long
    
'    Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR
'    Packet = Packet & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & END_CHAR
    Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).Speed & SEP_CHAR & Npc(NpcNum).Magi & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MAXHP & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR & Npc(NpcNum).Element & SEP_CHAR & Npc(NpcNum).Spritesize & SEP_CHAR
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
    Dim i As Integer
    
    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).currencyItem & SEP_CHAR
    For i = 1 To MAX_SHOP_ITEMS
        Packet = Packet & Shop(ShopNum).ShopItem(i).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(i).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(i).Price & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum)
    Dim Packet As String
    Dim i As Integer
    
    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).currencyItem & SEP_CHAR
    For i = 1 To MAX_SHOP_ITEMS
        Packet = Packet & Shop(ShopNum).ShopItem(i).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(i).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(i).Price & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditShopTo(ByVal index As Long, ByVal ShopNum As Long)
    Dim Packet As String
    Dim z As Integer
    
    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).currencyItem & SEP_CHAR
    For z = 1 To MAX_SHOP_ITEMS
        Packet = Packet & Shop(ShopNum).ShopItem(z).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(z).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(z).Price & SEP_CHAR
    Next z
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
    
    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Spell(SpellNum).Element & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

' Updated with new shop system -Pickle
Sub SendTrade(ByVal index As Long, ByVal ShopNum As Long)
    Dim Packet As String
    
    Packet = "GOSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & END_CHAR
'    All we need are the shop identifier and num - we don't need to send the entire shop every time
    Call SendDataTo(index, Packet)
    
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

Sub SendWierdTo(ByVal index As Long)
    Dim Packet As String
    Packet = "WIERD" & SEP_CHAR & Wierd & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
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
Sub SendGameClockTo(ByVal index As Long)
    Dim Packet As String
    
    Packet = "GAMECLOCK" & SEP_CHAR & Seconds & SEP_CHAR & Minutes & SEP_CHAR & Hours & SEP_CHAR & Gamespeed & SEP_CHAR & END_CHAR
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
Sub SendNewsTo(ByVal index As Long)
    Dim Packet As String
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer
    
    On Error GoTo NewsError
    Red = Val(ReadINI("COLOR", "Red", App.Path & "\News.ini"))
    Green = Val(ReadINI("COLOR", "Green", App.Path & "\News.ini"))
    Blue = Val(ReadINI("COLOR", "Blue", App.Path & "\News.ini"))
    
    Packet = "NEWS" & SEP_CHAR & ReadINI("DATA", "ServerNews", App.Path & "\News.ini") & SEP_CHAR
    Packet = Packet & Red & SEP_CHAR & Green & SEP_CHAR & Blue & SEP_CHAR & ReadINI("DATA", "Desc", App.Path & "\News.ini") & END_CHAR
    
    Call SendDataTo(index, Packet)
    Exit Sub
    
NewsError:
'    Error reading the news, so just send white
    Red = 255
    Green = 255
    Blue = 255
    
    Packet = "NEWS" & SEP_CHAR & ReadINI("DATA", "ServerNews", App.Path & "\News.ini") & SEP_CHAR
    Packet = Packet & Red & SEP_CHAR & Green & SEP_CHAR & Blue & SEP_CHAR & END_CHAR
    
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
    Next i
    
    Call SpawnAllMapNpcs
End Sub

Sub MapMsg2(ByVal MapNum As Long, ByVal Msg As String, ByVal index As Long)
    Dim Packet As String
    
    Packet = "MAPMSG2" & SEP_CHAR & Msg & SEP_CHAR & index & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
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

Sub Sendsprite(ByVal index As Long, ByVal indexto As Long)
    Dim Packet As String
    
    Packet = "CusSprite" & SEP_CHAR & index & SEP_CHAR & player(index).Char(player(index).charnum).head & SEP_CHAR & player(index).Char(player(index).charnum).body & SEP_CHAR & player(index).Char(player(index).charnum).leg & SEP_CHAR & END_CHAR
    Call SendDataTo(indexto, Packet)
End Sub

Sub SendActionNames(ByVal index As Long)
    Dim Packet As String
    Dim i As Long
    
    Packet = "actionname" & SEP_CHAR & ReadINI("ACTION", "max", App.Path & "\Data.ini") & SEP_CHAR
    For i = 1 To ReadINI("ACTION", "max", App.Path & "\Data.ini")
        Packet = Packet & ReadINI("ACTION", "name" & i, App.Path & "\Data.ini") & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub GrapleHook(ByVal index As Long)
    Dim x As Long, y As Long, MapNum As Long
    MapNum = GetPlayerMap(index)
    
    If player(index).HookShotX <> 0 Or player(index).HookShotY <> 0 Then
        If player(index).locked = True Then
            Call PlayerMsg(index, "You can only fire one grappleshot at the time", 1)
            Exit Sub
        End If
    End If
    
    player(index).locked = True
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
    
    If GetPlayerDir(index) = DIR_DOWN Then
        x = GetPlayerX(index)
        y = GetPlayerY(index) + 1
        Do While y <= MAX_MAPY
            If Map(MapNum).tile(x, y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                player(index).HookShotX = x
                player(index).HookShotY = y
                Exit Sub
            Else
                If Map(MapNum).tile(x, y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                    player(index).HookShotX = x
                    player(index).HookShotY = y
                    Exit Sub
                End If
            End If
            y = y + 1
        Loop
        Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        player(index).HookShotX = x
        player(index).HookShotY = y
        Exit Sub
    End If
    If GetPlayerDir(index) = DIR_UP Then
        x = GetPlayerX(index)
        y = GetPlayerY(index) - 1
        Do While y >= 0
            If Map(MapNum).tile(x, y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                player(index).HookShotX = x
                player(index).HookShotY = y
                Exit Sub
            Else
                If Map(MapNum).tile(x, y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                    player(index).HookShotX = x
                    player(index).HookShotY = y
                    Exit Sub
                End If
            End If
            y = y - 1
        Loop
        Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        player(index).HookShotX = x
        player(index).HookShotY = y
        Exit Sub
    End If
    
    If GetPlayerDir(index) = DIR_RIGHT Then
        x = GetPlayerX(index) + 1
        y = GetPlayerY(index)
        Do While x <= MAX_MAPX
            If Map(MapNum).tile(x, y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                player(index).HookShotX = x
                player(index).HookShotY = y
                Exit Sub
            Else
                If Map(MapNum).tile(x, y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                    player(index).HookShotX = x
                    player(index).HookShotY = y
                    Exit Sub
                End If
            End If
            x = x + 1
        Loop
        Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        player(index).HookShotX = x
        player(index).HookShotY = y
        Exit Sub
    End If
    
    If GetPlayerDir(index) = DIR_LEFT Then
        x = GetPlayerX(index) - 1
        y = GetPlayerY(index)
        Do While x >= 0
            If Map(MapNum).tile(x, y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                player(index).HookShotX = x
                player(index).HookShotY = y
                Exit Sub
            Else
                If Map(MapNum).tile(x, y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                    player(index).HookShotX = x
                    player(index).HookShotY = y
                    Exit Sub
                End If
            End If
            x = x - 1
        Loop
        Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
        player(index).HookShotX = x
        player(index).HookShotY = y
        Exit Sub
    End If
End Sub

' Used for checking if a port is open, and opens if not
Sub CheckOpenPort(ByVal Port As Integer)
    On Error GoTo Err:
    
'    Check if ports.ini exists
    If VarExists(App.Path & "\Ports.ini", "PORTS", STR(Port)) Then
        If Val(GetVar(App.Path & "\Ports.ini", "PORTS", STR(Port))) = 1 Then
            Exit Sub
        End If
    End If
    
'    Prompy user
    If MsgBox("Would you like Eclipse to open a port in your firewall to run the server? People may not be able to connect if the port is closed.", vbYesNo) = vbYes Then
'        Use the shell to open the port
        Shell ("netsh firewall add portopening TCP " & Port & " EclipseEvolution-ServerPort")
        Call MsgBox("Port " & Port & " opened.", vbOKOnly, "Success!")
    Else
'        Keep going
        Call MsgBox("No action taken. You can open the port manually at a later date. This will not prompt you again.", vbOKOnly)
    End If
    
'    Write value to ports.ini
    Call PutVar(App.Path & "\Ports.ini", "PORTS", STR(Port), 1)
    Exit Sub
    
Err:
    Call MsgBox("Error occured - " & Err.Description & ". Running normally.", vbCritical)
    Err.Clear
    Call PutVar(App.Path & "\Ports.ini", "PORTS", STR(Port), 1)
End Sub



