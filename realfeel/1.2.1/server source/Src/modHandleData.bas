Attribute VB_Name = "modHandleData"
Option Explicit

Sub HandleData(ByVal Index As Long, ByVal Data As String)
'On Error GoTo errorhandler:

Dim Parse() As String
Dim Name As String
Dim Password As String
Dim Sex As Long
Dim ClassNum As Long
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
Dim Packet As String, FileName As String
Dim InvSlot As Byte
        
    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
        
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Requesting classes for making a character ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "getclasses" Then
        If Not IsPlaying(Index) Then
            Call SendStrings(Index)
            Call SendNewCharClasses(Index)
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::::
    ' :: New account packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "newaccount" Then
        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
        
            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = Asc(Mid(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next i
            
            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password)
                Call TextAdd(frmServer.txtText, "Account " & Name & " has been created.", True)
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                Call AlertMsg(Index, "Your account has been created!")
            Else
                Call AlertMsg(Index, "Sorry, that account name is already taken!")
            End If
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Delete account packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "delaccount" Then
        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
            
            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "The name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If
            
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If
                        
            ' Delete names from master name file
            Call LoadPlayer(Index, Name)
            For i = 1 To MAX_CHARS
                If Trim$(Player(Index).Char(i).Name) <> "" Then
                    Call DeleteName(Player(Index).Char(i).Name)
                End If
            Next i
            Call ClearPlayer(Index)
            
            ' Everything went ok
            Call Kill(App.Path & "\accounts\" & Trim$(Name) & ".ini")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(Index, "Your account has been deleted.")
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::
    ' :: Login packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "login" Then
        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
        
            ' Check versions
            If Val(Parse(3)) < CLIENT_MAJOR Or Val(Parse(4)) < CLIENT_MINOR Or Val(Parse(5)) < CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit " & GAME_WEBSITE)
                Exit Sub
            End If
            
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If
        
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If
        
            If IsMultiAccounts(Name) Then
                Call AlertMsg(Index, "Multiple account logins is not authorized.")
                Exit Sub
            End If
            
            ' Prevent Duping
            For i = 1 To Len(Name)
            n = Asc(Mid(Name, i, 1))
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(Index, "Account duping detected! Exiting client!")
                    Exit Sub
                End If
            Next i
                
            ' Everything went ok
            'Now send the max of every data
            Packet = "MAXDATA" & SEP_CHAR
            Packet = Packet & GAME_NAME & SEP_CHAR
            Packet = Packet & MAX_PLAYERS & SEP_CHAR
            Packet = Packet & MAX_MAPS & SEP_CHAR
            Packet = Packet & MAX_ITEMS & SEP_CHAR
            Packet = Packet & MAX_NPCS & SEP_CHAR
            Packet = Packet & MAX_SHOPS & SEP_CHAR
            Packet = Packet & MAX_SPELLS & SEP_CHAR
            Packet = Packet & END_CHAR
            Call SendDataTo(Index, Packet)
            
            ' Load the player
            Call LoadPlayer(Index, Name)
            Call SendChars(Index)
    
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", True)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Add character packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "addchar" Then
        If Not IsPlaying(Index) Then
            Name = Parse(1)
            Sex = Val(Parse(2))
            ClassNum = Val(Parse(3))
            CharNum = Val(Parse(4))
        
            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Then
                Call AlertMsg(Index, "Character name must be at least three characters in length.")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = Asc(Mid(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next i
                                    
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Prevent hacking
            If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
                Call HackingAttempt(Index, "Invalid Sex (dont laugh)")
                Exit Sub
            End If
            
            ' Prevent hacking
            If ClassNum < 0 Or ClassNum > Max_Classes Then
                Call HackingAttempt(Index, "Invalid Class")
                Exit Sub
            End If
        
            ' Check if char already exists in slot
            If CharExist(Index, CharNum) Then
                Call AlertMsg(Index, "Character already exists!")
                Exit Sub
            End If
            
            ' Check if name is already in use
            If FindChar(Name) Then
                Call AlertMsg(Index, "Sorry, but that name is in use!")
                Exit Sub
            End If
        
            ' Everything went ok, add the character
            Call AddChar(Index, Name, Sex, ClassNum, CharNum)
            Call SavePlayer(Index)
            Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
            Call AlertMsg(Index, "Character has been created!")
        End If
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::::::::
    ' :: Deleting character packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "delchar" Then
        If Not IsPlaying(Index) Then
            CharNum = Val(Parse(1))
        
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If
            
            Call DelChar(Index, CharNum)
            Call AddLog("Character deleted on " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
            Call AlertMsg(Index, "Character has been deleted!")
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Using character packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "usechar" Then
        If Not IsPlaying(Index) Then
            CharNum = Val(Parse(1))
        
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Check if the server is unlocked
            If ServerState = SERVER_LOCKED Then
                If Player(Index).Char(CharNum).Access < 2 Then
                    Call AlertMsg(Index, "Server Closed for Maintenance.")
                    Exit Sub
                End If
            End If
            
            ' Check to make sure the character exists and if so, set it as its current char
            If CharExist(Index, CharNum) Then
                Player(Index).CharNum = CharNum
                Call JoinGame(Index)
            
                CharNum = Player(Index).CharNum
                Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
                Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", True)
                Call UpdateCaption
                
            'Add character's name to list
            frmServer.lstPlayers.AddItem (GetPlayerLogin(Index) & "/" & GetPlayerName(Index))
            If ServerState = SERVER_UNLOCKED Then
                frmServer.txtTotal.Text = TotalOnlinePlayers
            End If
                
                ' Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
                If Not FindChar(GetPlayerName(Index)) Then
                    f = FreeFile
                    Open App.Path & "\accounts\charlist.txt" For Append As #f
                        Print #f, GetPlayerName(Index)
                    Close #f
                End If
            Else
                Call AlertMsg(Index, "Character does not exist!")
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
                Call HackingAttempt(Index, "Say Text Modification")
                Exit Sub
            End If
        Next i
        
        'Set player text
        'Player(Index).Char(Player(Index).CharNum).Text = Msg
        
        'Not needed
        'Call PutVar(App.Path & "\accounts\" & GetPlayerName(Index), "CHAR" & Player(Index).CharNum, "Text", Player(Index).Char(i).Text)
        
        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Msg, Grey)
        
        'Add message to the log
        Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & ": " & Msg, PLAYER_LOG)
        
        For n = 1 To MAX_TRACKERS
            If Player(Index).Char(Player(Index).CharNum).Trackers(n) <> "" Then Call TrackerMsg(FindPlayer(Player(Index).Char(Player(Index).CharNum).Trackers(n)), "TRACKER", "(Regular) " & GetPlayerName(Index) & ": " & Msg)
        Next n
        
        'Update tracker
        Call ATM("MAP", GetPlayerName(Index) & ": " & Msg)
        
        'Run commands script
        'MyScript.ExecuteStatement "main.txt", "Commands " & Index
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "emotemsg" Then
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Emote Text Modification")
                Exit Sub
            End If
        Next i
        
        Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Msg, EmoteColor)
        For n = 1 To MAX_TRACKERS
            If Player(Index).Char(Player(Index).CharNum).Trackers(n) <> "" Then Call TrackerMsg(FindPlayer(Player(Index).Char(Player(Index).CharNum).Trackers(n)), "TRACKER", "(Emote) " & GetPlayerName(Index) & ": " & Msg)
        Next n
        Call ATM("EMOTE", GetPlayerName(Index) & ": " & Msg)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "broadcastmsg" Then
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Broadcast Text Modification")
                Exit Sub
            End If
        Next i
        
        s = GetPlayerName(Index) & ": " & Msg
        Call AddLog(s, PLAYER_LOG)
        Call GlobalMsg(s, BroadcastColor)
        For n = 1 To MAX_TRACKERS
            If Player(Index).Char(Player(Index).CharNum).Trackers(n) <> "" Then Call TrackerMsg(FindPlayer(Player(Index).Char(Player(Index).CharNum).Trackers(n)), "TRACKER", "(Broadcast) " & GetPlayerName(Index) & ": " & Msg)
        Next n
        Call ATM("BROADCAST", GetPlayerName(Index) & ": " & Msg)
        Call TextAdd(frmServer.txtText, s, True)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "globalmsg" Then
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Global Text Modification")
                Exit Sub
            End If
        Next i
        
        If GetPlayerAccess(Index) > 0 Then
            s = "(global) " & GetPlayerName(Index) & ": " & Msg
            Call AddLog(s, ADMIN_LOG)
            Call GlobalMsg(s, GlobalColor)
            For n = 1 To MAX_TRACKERS
                If Player(Index).Char(Player(Index).CharNum).Trackers(n) <> "" Then Call TrackerMsg(FindPlayer(Player(Index).Char(Player(Index).CharNum).Trackers(n)), "TRACKER", "(Global) " & GetPlayerName(Index) & ": " & Msg)
            Next n
            Call ATM("GLOBAL", GetPlayerName(Index) & ": " & Msg)
            Call TextAdd(frmServer.txtText, s, True)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "adminmsg" Then
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Admin Text Modification")
                Exit Sub
            End If
        Next i
        
        If GetPlayerAccess(Index) > 0 Then
            Call AddLog("(admin " & GetPlayerName(Index) & ") " & Msg, ADMIN_LOG)
            Call AdminMsg("(admin " & GetPlayerName(Index) & ") " & Msg, AdminColor)
            For n = 1 To MAX_TRACKERS
                If Player(Index).Char(Player(Index).CharNum).Trackers(n) <> "" Then Call TrackerMsg(FindPlayer(Player(Index).Char(Player(Index).CharNum).Trackers(n)), "TRACKER", "(Admin) " & GetPlayerName(Index) & ": " & Msg)
            Next n
            Call ATM("ADMIN", GetPlayerName(Index) & ": " & Msg)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "playermsg" Then
        MsgTo = FindPlayer(Parse(1))
        Msg = Parse(2)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Player Msg Text Modification")
                Exit Sub
            End If
        Next i
        
        ' Check if they are trying to talk to themselves
        If MsgTo <> Index Then
            If MsgTo > 0 Then
                Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
                Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & Msg & "'", TellColor)
                Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
                For n = 1 To MAX_TRACKERS
                    If Player(Index).Char(Player(Index).CharNum).Trackers(n) <> "" Then Call TrackerMsg(FindPlayer(Player(Index).Char(Player(Index).CharNum).Trackers(n)), "TRACKER", "(Private) " & GetPlayerName(Index) & ": " & Msg)
                Next n
                Call ATM("PRIVATE", "(" & GetPlayerName(Index) & " to " & GetPlayerName(MsgTo) & "): " & Msg)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " begins to mumble to himself, what a wierdo...", PLAYER_LOG)
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " begins to mumble to himself, what a wierdo...", Green)
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playermove" And Player(Index).GettingMap = NO Then
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
        'Run Scripts
        For i = 0 To frmLibrary.lstLibrary.ListCount - 1
            FileName = App.Path & "\Library\" & frmLibrary.lstLibrary.List(i)
            If GetVar(FileName, "DATA", "Enabled") = "True" Then
                MyScript.ExecuteStatement frmLibrary.lstLibrary.List(i), "OnTile " & Index
            End If
        Next i
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerdir" And Player(Index).GettingMap = NO Then
        Dir = Val(Parse(1))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(Index, "Invalid Direction")
            Exit Sub
        End If
        
        Call SetPlayerDir(Index, Dir)
        Call SendDataToMapBut(Index, GetPlayerMap(Index), "PLAYERDIR" & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
        
    ' :::::::::::::::::::::
    ' :: Use item packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "useitem" Then
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
            
            ' Find out what kind of item it is
            Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum <> GetPlayerArmorSlot(Index) Then
                        If Int(GetPlayerDEF(Index)) < n Then
                            Call PlayerMsg(Index, "Your " & STRING_DEFENSE & " is to low to wear this armor!  Required DEF (" & n & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerArmorSlot(Index, InvNum)
                    Else
                        Call SetPlayerArmorSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum <> GetPlayerWeaponSlot(Index) Then
                        If Int(GetPlayerSTR(Index)) < n Then
                            Call PlayerMsg(Index, "Your " & STRING_STRENGTH & " is to low to hold this weapon!  Required STR (" & n & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerWeaponSlot(Index, InvNum)
                    Else
                        Call SetPlayerWeaponSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
                        
                Case ITEM_TYPE_HELMET
                    If InvNum <> GetPlayerHelmetSlot(Index) Then
                        If Int(GetPlayerSPEED(Index)) < n Then
                            Call PlayerMsg(Index, "Your " & STRING_SPEED & " is to low to wear this helmet!  Required SPEED (" & n & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerHelmetSlot(Index, InvNum)
                    Else
                        Call SetPlayerHelmetSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
            
                Case ITEM_TYPE_SHIELD
                    If InvNum <> GetPlayerShieldSlot(Index) Then
                        Call SetPlayerShieldSlot(Index, InvNum)
                    Else
                        Call SetPlayerShieldSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
            
                Case ITEM_TYPE_POTIONADDHP
                    Call SetPlayerHP(Index, GetPlayerHP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendHP(Index)
                    ' Send message to play sound
                    Call SendDataToMap(GetPlayerMap(Index), "PLAYSOUND" & SEP_CHAR & Trim$(Item(InvNum).Sound) & SEP_CHAR & END_CHAR)
                
                Case ITEM_TYPE_POTIONADDMP
                    Call SetPlayerMP(Index, GetPlayerMP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendMP(Index)
                    ' Send message to play sound
                    Call SendDataToMap(GetPlayerMap(Index), "PLAYSOUND" & SEP_CHAR & Trim$(Item(InvNum).Sound) & SEP_CHAR & END_CHAR)
        
                Case ITEM_TYPE_POTIONADDSP
                    Call SetPlayerSP(Index, GetPlayerSP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendSP(Index)
                    ' Send message to play sound
                    Call SendDataToMap(GetPlayerMap(Index), "PLAYSOUND" & SEP_CHAR & Trim$(Item(InvNum).Sound) & SEP_CHAR & END_CHAR)

                Case ITEM_TYPE_POTIONSUBHP
                    Call SetPlayerHP(Index, GetPlayerHP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendHP(Index)
                    ' Send message to play sound
                    Call SendDataToMap(GetPlayerMap(Index), "PLAYSOUND" & SEP_CHAR & Trim$(Item(InvNum).Sound) & SEP_CHAR & END_CHAR)
                    'Call SendDataTo(Index, "PLAYSOUND" & SEP_CHAR & Trim$(Item(InvNum).Sound) & SEP_CHAR & END_CHAR)
                
                Case ITEM_TYPE_POTIONSUBMP
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendMP(Index)
                    ' Send message to play sound
                    Call SendDataToMap(GetPlayerMap(Index), "PLAYSOUND" & SEP_CHAR & Trim$(Item(InvNum).Sound) & SEP_CHAR & END_CHAR)
        
                Case ITEM_TYPE_POTIONSUBSP
                    Call SetPlayerSP(Index, GetPlayerSP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendSP(Index)
                    ' Send message to play sound
                    Call SendDataToMap(GetPlayerMap(Index), "PLAYSOUND" & SEP_CHAR & Trim$(Item(InvNum).Sound) & SEP_CHAR & END_CHAR)
                    
                Case ITEM_TYPE_KEY
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            If GetPlayerY(Index) > 0 Then
                                x = GetPlayerX(Index)
                                y = GetPlayerY(Index) - 1
                            Else
                                Exit Sub
                            End If
                            
                        Case DIR_DOWN
                            If GetPlayerY(Index) < MAX_MAPY Then
                                x = GetPlayerX(Index)
                                y = GetPlayerY(Index) + 1
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_LEFT
                            If GetPlayerX(Index) > 0 Then
                                x = GetPlayerX(Index) - 1
                                y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_RIGHT
                            If GetPlayerX(Index) < MAX_MAPY Then
                                x = GetPlayerX(Index) + 1
                                y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                    End Select
                    
                    ' Check if a key exists
                    If Map(GetPlayerMap(Index)).Tile(x, y).Key = True Then
                        ' Check if the key they are using matches the map key
                        If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(x, y).KeyNum Then
                            TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            
                            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            ' Send message to play sound
                            Call SendDataToMap(GetPlayerMap(Index), "PLAYSOUND" & SEP_CHAR & Trim$(Item(InvNum).Sound) & SEP_CHAR & END_CHAR)
                            Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
                            
                            ' Check if we are supposed to take away the item
                            If Map(GetPlayerMap(Index)).Tile(x, y).KeyTake = 1 Then
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                Call PlayerMsg(Index, "The key breaks.", Yellow)
                            End If
                        End If
                    End If
                    
                Case ITEM_TYPE_SPELL
                    ' Get the spell num
                    n = Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                    
                    If n > 0 Then
                        ' Make sure they are the right class
                        If Spell(n).ClassReq - 1 = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                            ' Make sure they are the right level
                            i = GetSpellReqLevel(Index, n)
                            If i <= GetPlayerLevel(Index) Then
                                i = FindOpenSpellSlot(Index)
                                
                                ' Make sure they have an open spell slot
                                If i > 0 Then
                                    ' Make sure they dont already have the spell
                                    If Not HasSpell(Index, n) Then
                                        Call SetPlayerSpell(Index, i, n)
                                        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                        Call PlayerMsg(Index, "You study the spell carefully...", Yellow)
                                        Call PlayerMsg(Index, "You have learned a new spell!", White)
                                        ' Send message to play sound
                                        Call SendDataToMap(GetPlayerMap(Index), "PLAYSOUND" & SEP_CHAR & Trim$(Item(InvNum).Sound) & SEP_CHAR & END_CHAR)
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
                            Call PlayerMsg(Index, "This spell can only be learned by a " & GetClassName(Spell(n).ClassReq - 1) & ".", White)
                        End If
                    Else
                        Call PlayerMsg(Index, "This scroll is not connected to a spell, please inform an admin!", White)
                    End If
                    Call SendEquipDataTo(Index)
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
            If i <> Index Then
                If GetPlayerWeaponSlot(Index) > 0 And GetPlayerShieldSlot(Index) > 0 Then
                'See if the player is using a bow, if so, check range
                If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data4 = WEAPON_TYPE_BOW Then
                    'See if the player has any arrows
                    If GetPlayerShieldSlot(Index) < 1 Then
                        Call PlayerMsg(Index, "You have no arrows to fire!", 1)
                        Exit Sub
                    End If
                    
                    If Player(Index).Target < 1 Then
                        Call PlayerMsg(Index, "You have must select a target!", 1)
                        Exit Sub
                    End If
                    
                If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
                    'See if the arrow will break
                    If WillArrowSnap(Index) Then
                        Call PlayerMsg(Index, "The arrow snapped in your hands!", Yellow)
                        Call SetPlayerInvItemDur(Index, GetPlayerShieldSlot(Index), GetPlayerInvItemDur(Index, GetPlayerShieldSlot(Index)) - 1)
                            If GetPlayerInvItemDur(Index, GetPlayerShieldSlot(Index)) <= 0 Then
                                Call PlayerMsg(Index, "All your arrows have broken!", Yellow)
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index)), 0)
                            End If
                        Exit Sub
                    End If
                    
                    If Not CanPlayerBlockHit(i) Then
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(Index) Then
                            n = Int((((GetPlayerDamage(Index) / 5) * 4) * Rnd) + 1)
                            If n < (n * 4 / 5) Then
                                n = (n * 4 / 5)
                            End If
                            Damage = n - GetPlayerProtection(i)
                        Else
                            n = GetPlayerDamage(Index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                            Call PlayerMsg(Index, "You feel a strong vibration from your bow upon releasing your arrow!", BrightCyan)
                            Call PlayerMsg(i, GetPlayerName(Index) & " releases an arrow with tremendous force!", BrightCyan)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(Index, Player(Index).Target, Damage)
                        Else
                            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                        End If
                    Else
                        Call PlayerMsg(Index, GetPlayerName(i) & "'s " & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked your hit!", BrightCyan)
                        Call PlayerMsg(i, "Your " & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan)
                        ' Send message to play sound
                        Call SendDataToMap(GetPlayerMap(Index), "PLAYSOUND" & SEP_CHAR & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Sound) & SEP_CHAR & END_CHAR)
                    End If
                    
                    Exit Sub
                End If
                End If
                End If
            
                ' Can we attack the player?
                If GetPlayerWeaponSlot(Index) > 0 Then
                If CanAttackPlayer(Index, i) And Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data4 <> WEAPON_TYPE_BOW Then
                    If Not CanPlayerBlockHit(i) Then
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(Index) Then
                            n = Int((((GetPlayerDamage(Index) / 5) * 4) * Rnd) + 1)
                            If n < (n * 4 / 5) Then
                                n = (n * 4 / 5)
                            End If
                            Damage = n - GetPlayerProtection(i)
                        Else
                            n = GetPlayerDamage(Index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                            Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                            Call PlayerMsg(i, GetPlayerName(Index) & " swings with enormous might!", BrightCyan)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(Index, i, Damage)
                        Else
                            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                        End If
                    Else
                        Call PlayerMsg(Index, GetPlayerName(i) & "'s " & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked your hit!", BrightCyan)
                        Call PlayerMsg(i, "Your " & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan)
                        ' Send message to play sound
                        Call SendDataToMap(GetPlayerMap(Index), "PLAYSOUND" & SEP_CHAR & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Sound) & SEP_CHAR & END_CHAR)
                    End If
                    
                    Exit Sub
                End If
                Else
                ' Can we attack the player?
                If CanAttackPlayer(Index, i) Then
                    If Not CanPlayerBlockHit(i) Then
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(Index) Then
                            n = Int((((GetPlayerDamage(Index) / 5) * 4) * Rnd) + 1)
                            If n < (n * 4 / 5) Then
                                n = (n * 4 / 5)
                            End If
                            Damage = n - GetPlayerProtection(i)
                        Else
                            n = GetPlayerDamage(Index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                            Call PlayerMsg(Index, "You feel a surge of energy upon striking!", BrightCyan)
                            Call PlayerMsg(i, GetPlayerName(Index) & " attacks with enormous might!", BrightCyan)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(Index, i, Damage)
                        Else
                            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                        End If
                    Else
                        Call PlayerMsg(Index, GetPlayerName(i) & "'s " & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked your hit!", BrightCyan)
                        Call PlayerMsg(i, "Your " & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan)
                        ' Send message to play sound
                        Call SendDataToMap(GetPlayerMap(Index), "PLAYSOUND" & SEP_CHAR & Trim$(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Sound) & SEP_CHAR & END_CHAR)
                    End If
                    
                    Exit Sub
                End If
                End If
            End If
        Next i
        
        ' Try to attack a npc
        For i = 1 To MAX_MAP_NPCS
            If GetPlayerWeaponSlot(Index) < 1 Then
            ' Can we attack the npc?
            If CanAttackNpc(Index, i) Then
                ' Get the damage we can do
                If Not CanPlayerCriticalHit(Index) Then
                    n = Int((((GetPlayerDamage(Index) / 5) * 4) * Rnd) + 1)
                    If n < (n * 4 / 5) Then
                        n = (n * 4 / 5)
                    End If
                    Damage = n - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
                Else
                    n = GetPlayerDamage(Index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
                    Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(Index, i, Damage)
                    
                    'Send the NPC's HP
                    '-smchronos
                    
                Else
                    Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                End If
                Exit Sub
            End If
            Else
            If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data4 = WEAPON_TYPE_BOW Then
                'See if the player has any arrows
                    If GetPlayerShieldSlot(Index) < 1 Then
                        Call PlayerMsg(Index, "You have no arrows to fire!", 1)
                        Exit Sub
                    End If
                    
                    If Player(Index).Target < 1 Then
                        Call PlayerMsg(Index, "You have must select a target!", 1)
                        Exit Sub
                    End If
                    
                If Player(Index).TargetType = TARGET_TYPE_NPC Then
                    'See if the arrow will break
                    If WillArrowSnap(Index) Then
                        Call PlayerMsg(Index, "The arrow snapped in your hands!", Yellow)
                        Call SetPlayerInvItemDur(Index, GetPlayerShieldSlot(Index), GetPlayerInvItemDur(Index, GetPlayerShieldSlot(Index)) - 1)
                            If GetPlayerInvItemDur(Index, GetPlayerShieldSlot(Index)) <= 0 Then
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index)), 0)
                            End If
                        Exit Sub
                    End If
                    
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(Index) Then
                            n = Int((((GetPlayerDamage(Index) / 5) * 4) * Rnd) + 1)
                            If n < (n * 4 / 5) Then
                                n = (n * 4 / 5)
                            End If
                            Damage = n - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
                        Else
                            n = GetPlayerDamage(Index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), Player(Index).Target).Num).DEF / 2)
                            Call PlayerMsg(Index, "You feel a strong vibration from your bow upon releasing the arrow!", BrightCyan)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackNpc(Index, Player(Index).Target, Damage)
                        Else
                            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                        End If
                        Exit Sub
                End If
            Else
            ' Can we attack the npc?
            If CanAttackNpc(Index, i) Then
                ' Get the damage we can do
                If Not CanPlayerCriticalHit(Index) Then
                    n = Int((((GetPlayerDamage(Index) / 5) * 4) * Rnd) + 1)
                    If n < (n * 4 / 5) Then
                        n = (n * 4 / 5)
                    End If
                    Damage = n - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
                Else
                    n = GetPlayerDamage(Index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
                    Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(Index, i, Damage)
                    
                    'Send the NPC's HP
                    '-smchronos
                    
                Else
                    Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                End If
                Exit Sub
            End If
            End If
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
            Call HackingAttempt(Index, "Invalid Point Type")
            Exit Sub
        End If
                
        ' Make sure they have points
        If GetPlayerPOINTS(Index) > 0 Then
            ' Take away a stat point
            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)
            i = GetPlayerPOINTS(Index)
            ' Everything is ok
            Select Case PointType
                Case 0
                    Call SetPlayerSTR(Index, GetPlayerSTR(Index) + 1)
                    Call PlayerMsg(Index, "You have gained more " & STRING_STRENGTH & "!", White)
                Case 1
                    Call SetPlayerDEF(Index, GetPlayerDEF(Index) + 1)
                    Call PlayerMsg(Index, "You have gained more " & STRING_DEFENSE & "!", White)
                Case 2
                    Call SetPlayerMAGI(Index, GetPlayerMAGI(Index) + 1)
                    Call PlayerMsg(Index, "You have gained more " & STRING_MAGIC & "!", White)
                Case 3
                    Call SetPlayerSPEED(Index, GetPlayerSPEED(Index) + 1)
                    Call PlayerMsg(Index, "You have gained more " & STRING_SPEED & "!", White)
            End Select
        Else
            Call PlayerMsg(Index, "You have no skill points to train with!", BrightRed)
        End If
        
        ' Send the update
        Call SendStats(Index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Player info request packet ::
    ' ::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinforequest" Then
        Name = Parse(1)
        
        i = FindPlayer(Name)
        If i > 0 Then
            Call PlayerMsg(Index, "Account: " & Trim$(Player(i).Login) & ", Name: " & GetPlayerName(i), BrightGreen)
            If GetPlayerAccess(Index) > ADMIN_MONITER Then
                Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
                Call PlayerMsg(Index, "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), BrightGreen)
                Call PlayerMsg(Index, STRING_HP & ": " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  " & STRING_MP & ": " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  " & STRING_SP & ": " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
                Call PlayerMsg(Index, STRING_STRENGTH & ": " & GetPlayerSTR(i) & "  " & STRING_DEFENSE & ": " & GetPlayerDEF(i) & "  " & STRING_MAGIC & ": " & GetPlayerMAGI(i) & "  " & STRING_SPEED & ": " & GetPlayerSPEED(i), BrightGreen)
                n = Int(GetPlayerSTR(i) / 2) + Int(GetPlayerLevel(i) / 2)
                i = Int(GetPlayerDEF(i) / 2) + Int(GetPlayerLevel(i) / 2)
                If n > 100 Then n = 100
                If i > 100 Then i = 100
                Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", BrightGreen)
            End If
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Warp me to packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "warpmeto" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        n = FindPlayer(Parse(1))
        
        If n <> Index Then
            If n > 0 Then
                Call PlayerWarp(Index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
                Call PlayerMsg(n, GetPlayerName(Index) & " has warped to you.", BrightBlue)
                Call PlayerMsg(Index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
                Call AddLog(GetPlayerName(Index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot warp to yourself!", White)
        End If
                
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::
    ' :: Warp to me packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "warptome" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        n = FindPlayer(Parse(1))
        
        If n <> Index Then
            If n > 0 Then
                Call PlayerWarp(n, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(Index) & ".", BrightBlue)
                Call PlayerMsg(Index, GetPlayerName(n) & " has been summoned.", BrightBlue)
                Call AddLog(GetPlayerName(Index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(Index) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "You cannot warp yourself to yourself!", White)
        End If
        
        Exit Sub
    End If


    ' ::::::::::::::::::::::::
    ' :: Warp to map packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "warpto" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The map
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_MAPS Then
            Call HackingAttempt(Index, "Invalid map")
            Exit Sub
        End If
        
        Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
        Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
        Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Set sprite packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "setsprite" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The sprite
        n = Val(Parse(1))
        
        Call SetPlayerSprite(Index, n)
        Call SendPlayerData(Index)
        Exit Sub
    End If
                
    ' ::::::::::::::::::::::::::
    ' :: Stats request packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "getstats" Then
        'n is critical, i is blockage
        'If n > 100 Then n = 100
        'If i > 100 Then i = 100
        Packet = "pstats" & SEP_CHAR & GetPlayerSTR(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR
        Packet = Packet & GetPlayerName(Index) & SEP_CHAR & GetPlayerLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & END_CHAR
        Call SendDataTo(Index, Packet)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player request for a new map ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requestnewmap" Then
        Dir = Val(Parse(1))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(Index, "Invalid Direction")
            Exit Sub
        End If
                
        Call PlayerMove(Index, Dir, 1)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "mapdata" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        'Let's make sure no one will move
        Call SendDataToMap(GetPlayerMap(Index), "PAUSEMAP" & SEP_CHAR & "LOCK" & SEP_CHAR & "Updating map..." & SEP_CHAR & END_CHAR)
        DoEvents
        
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
        Map(MapNum).Shop = Val(Parse(n + 12))
        
        n = n + 13
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(MapNum).Tile(x, y).Ground = Val(Parse(n))
                Map(MapNum).Tile(x, y).Mask = Val(Parse(n + 1))
                Map(MapNum).Tile(x, y).Mask2 = Val(Parse(n + 2))
                Map(MapNum).Tile(x, y).Anim = Val(Parse(n + 3))
                Map(MapNum).Tile(x, y).Anim2 = Val(Parse(n + 4))
                Map(MapNum).Tile(x, y).Fringe = Val(Parse(n + 5))
                Map(MapNum).Tile(x, y).FringeAnim = Val(Parse(n + 6))
                Map(MapNum).Tile(x, y).Fringe2 = Val(Parse(n + 7))
                n = n + 8
            Next x
        Next y
        
        For x = 1 To MAX_MAP_NPCS
            Map(MapNum).Npc(x) = Val(Parse(n))
            n = n + 1
            'Call ClearMapNpc(x, MapNum)
        Next x
        Call SendMapNpcsToMap(MapNum)
        Call SpawnMapNpcs(MapNum)
        
        Call SaveMap(MapNum)
        
        ' Refresh map for everyone online
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
            End If
        Next i
        
        Exit Sub
    End If

    If LCase$(Parse(0)) = "mapattributes" Then
        n = 1
        MapNum = GetPlayerMap(Index)
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(MapNum).Tile(x, y).Walkable = Val(Parse(n))
                Map(MapNum).Tile(x, y).Blocked = Val(Parse(n + 1))
                Map(MapNum).Tile(x, y).Warp = Val(Parse(n + 2))
                Map(MapNum).Tile(x, y).WarpMap = Val(Parse(n + 3))
                Map(MapNum).Tile(x, y).WarpX = Val(Parse(n + 4))
                Map(MapNum).Tile(x, y).WarpY = Val(Parse(n + 5))
                Map(MapNum).Tile(x, y).Item = Val(Parse(n + 6))
                Map(MapNum).Tile(x, y).ItemNum = Val(Parse(n + 7))
                Map(MapNum).Tile(x, y).ItemValue = Val(Parse(n + 8))
                Map(MapNum).Tile(x, y).NpcAvoid = Val(Parse(n + 9))
                Map(MapNum).Tile(x, y).Key = Val(Parse(n + 10))
                Map(MapNum).Tile(x, y).KeyNum = Val(Parse(n + 11))
                Map(MapNum).Tile(x, y).KeyTake = Val(Parse(n + 12))
                Map(MapNum).Tile(x, y).KeyOpen = Val(Parse(n + 13))
                Map(MapNum).Tile(x, y).KeyOpenX = Val(Parse(n + 14))
                Map(MapNum).Tile(x, y).KeyOpenY = Val(Parse(n + 15))
                Map(MapNum).Tile(x, y).North = Val(Parse(n + 16))
                Map(MapNum).Tile(x, y).West = Val(Parse(n + 17))
                Map(MapNum).Tile(x, y).East = Val(Parse(n + 18))
                Map(MapNum).Tile(x, y).South = Val(Parse(n + 19))
                Map(MapNum).Tile(x, y).Shop = Val(Parse(n + 20))
                Map(MapNum).Tile(x, y).ShopNum = Val(Parse(n + 21))
                Map(MapNum).Tile(x, y).Bank = Val(Parse(n + 22))
                Map(MapNum).Tile(x, y).Heal = Val(Parse(n + 23))
                Map(MapNum).Tile(x, y).HealValue = Val(Parse(n + 24))
                Map(MapNum).Tile(x, y).Damage = Val(Parse(n + 25))
                Map(MapNum).Tile(x, y).DamageValue = Val(Parse(n + 26))
                Debug.Print "Warpmap: " & Map(MapNum).Tile(x, y).WarpMap
                n = n + 27
            Next x
        Next y
        Call SaveMap(MapNum)
        DoEvents
        
        'Free map from pause
        Call SendDataToMap(GetPlayerMap(Index), "PAUSEMAP" & SEP_CHAR & "UNLOCK" & SEP_CHAR & END_CHAR)
        DoEvents
    End If
    
    If LCase$(Parse(0)) = "nowsavemap" Then
        MapNum = Val(Parse(1))
        ' Save the map
        Call SaveMap(MapNum)
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Need map yes/no packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "needmap" Then
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
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to pick up something packet ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapgetitem" Then
        Call PlayerMapGetItem(Index)
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
            Call HackingAttempt(Index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Ammount > GetPlayerInvItemValue(Index, InvNum) Then
            Call HackingAttempt(Index, "Item ammount modification")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
            If Ammount <= 0 Then
                Call HackingAttempt(Index, "Trying to drop 0 ammount of currency")
                Exit Sub
            End If
        End If
            
        Call PlayerMapDropItem(Index, InvNum, Ammount)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Respawn map packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "maprespawn" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Clear out it all
        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
            Call ClearMapItem(i, GetPlayerMap(Index))
        Next i
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(Index))
        
        ' Respawn NPCS
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, GetPlayerMap(Index))
        Next i
        
        Call PlayerMsg(Index, "Map respawned.", Blue)
        Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Map report packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "mapreport" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If

        s = "Free Maps: "
        tMapStart = 1
        tMapEnd = 1
        
        For i = 1 To MAX_MAPS
            If Trim$(Map(i).Name) = "" Then
                tMapEnd = tMapEnd + 1
            Else
                If tMapEnd - tMapStart > 0 Then
                    s = s & Trim$(STR(tMapStart)) & "-" & Trim$(STR(tMapEnd - 1)) & ", "
                End If
                tMapStart = i + 1
                tMapEnd = i + 1
            End If
        Next i
        
        s = s & Trim$(STR(tMapStart)) & "-" & Trim$(STR(tMapEnd - 1)) & ", "
        s = Mid(s, 1, Len(s) - 2)
        s = s & "."
        
        Call PlayerMsg(Index, s, Brown)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Kick player packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "kickplayer" Then
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
    End If
        
    ' :::::::::::::::::::::
    ' :: Ban list packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "banlist" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = 1
        f = FreeFile
        Open App.Path & "\data\banlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            Input #f, Name
            
            Call PlayerMsg(Index, n & ": Banned IP " & s & " by " & Name, White)
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
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call Kill(App.Path & "\data\banlist.txt")
        Call PlayerMsg(Index, "Ban list destroyed.", White)
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::
    ' :: Ban player packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "banplayer" Then
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
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Request edit map packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditmap" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "EDITMAP" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit item packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requestedititem" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "ITEMEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit item packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "edititem" Then
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
    End If
    
    ' ::::::::::::::::::::::
    ' :: Save item packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "saveitem" Then
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
        Item(n).Description = Parse(3)
        Item(n).Pic = Val(Parse(4))
        Item(n).Type = Val(Parse(5))
        Item(n).Data1 = Val(Parse(6))
        Item(n).Data2 = Val(Parse(7))
        Item(n).Data3 = Val(Parse(8))
        Item(n).Data4 = Val(Parse(9))
        Item(n).Data5 = Val(Parse(10))
        Item(n).Sound = Trim$(Parse(11))
        
        ' Save it
        Call SendUpdateItemToAll(n)
        Call SaveItem(n)
        Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::
    ' :: Request edit npc packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditnpc" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "NPCEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Edit npc packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "editnpc" Then
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
    End If
    
    ' :::::::::::::::::::::
    ' :: Save npc packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "savenpc" Then
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
        Npc(n).DropChance = Val(Parse(8))
        Npc(n).DropItem = Val(Parse(9))
        Npc(n).DropItemValue = Val(Parse(10))
        Npc(n).HP = Val(Parse(11))
        Npc(n).STR = Val(Parse(12))
        Npc(n).DEF = Val(Parse(13))
        Npc(n).SPEED = Val(Parse(14))
        Npc(n).MAGI = Val(Parse(15))
        Npc(n).EXP = Val(Parse(16))
        Npc(n).Fear = CBool(Parse(17))
        Npc(n).TintR = Val(Parse(18))
        Npc(n).TintG = Val(Parse(19))
        Npc(n).TintB = Val(Parse(20))
        
        'Send a mapnpc behavior update
        Call SetAllMapNpcBehavior(n)
        
        ' Save it
        Call SendUpdateNpcToAll(n)
        Call SaveNpc(n)
        Call AddLog(GetPlayerName(Index) & " saved npc #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
            
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit shop packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditshop" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "SHOPEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "editshop" Then
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
    End If
    
    ' ::::::::::::::::::::::
    ' :: Save shop packet ::
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "saveshop") Then
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
        Shop(ShopNum).Restock = Val(Parse(6))
        
        n = 7
        For i = 1 To MAX_TRADES
            Shop(ShopNum).TradeItem(i).GiveItem = Val(Parse(n))
            Shop(ShopNum).TradeItem(i).GiveValue = Val(Parse(n + 1))
            Shop(ShopNum).TradeItem(i).GetItem = Val(Parse(n + 2))
            Shop(ShopNum).TradeItem(i).GetValue = Val(Parse(n + 3))
            Shop(ShopNum).TradeItem(i).Stock = Val(Parse(n + 4))
            Shop(ShopNum).TradeItem(i).MaxStock = Val(Parse(n + 5))
            n = n + 6
        Next i
        
        ' Save it
        Call SendUpdateShopToAll(ShopNum)
        Call SaveShop(ShopNum)
        Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Request edit spell packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditspell" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "SPELLEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "editspell" Then
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
    End If
    
    ' :::::::::::::::::::::::
    ' :: Save spell packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "savespell") Then
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
                
        ' Save it
        Call SendUpdateSpellToAll(n)
        Call SaveSpell(n)
        Call AddLog(GetPlayerName(Index) & " saving spell #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Request edit class ::
    ' ::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "requesteditclass") Then
    ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "CLASSEDITOR" & SEP_CHAR & END_CHAR)
    End If
    
    ' ::::::::::::::::::::::::
    ' ::     Edit class     ::
    ' ::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "editclass") Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The class #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > Max_Classes Then
            Call HackingAttempt(Index, "Invalid Class Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing class #" & n & ".", ADMIN_LOG)
        Call SendEditClassTo(Index, n)
    End If
    
    ' :::::::::::::::::::::::
    ' :: Save class packet ::
    ' :::::::::::::::::::::::
    If (LCase$(Parse(0)) = "saveclass") Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Class #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > Max_Classes Then
            Call HackingAttempt(Index, "Invalid Class Index")
            Exit Sub
        End If
        
        ' Update the class
        Class(n).Name = Parse(2)
        Class(n).Sprite = Parse(3)
        Class(n).HP = Parse(4)
        Class(n).MP = Parse(5)
        Class(n).SP = Parse(6)
        Class(n).STR = Parse(7)
        Class(n).DEF = Parse(8)
        Class(n).MAGI = Parse(9)
        Class(n).SPEED = Parse(10)
        Class(n).Map = Parse(11)
        Class(n).x = Parse(12)
        Class(n).y = Parse(13)
    
        'Debug.Print "Class save name"
        'Debug.Print Parse(2)
        'Debug.Print Class(n).Name
    
        ' Save it
        Call SendUpdateClassToAll(n)
        Call AddLog(GetPlayerName(Index) & " saving class #" & n & ".", ADMIN_LOG)
        Call SaveClass(n)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Delete class packet ::
    ' :::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "deleteclass") Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Delete old file
        Call Kill(App.Path & "\Data\classes.ini")
        
        ' Loop through, find all classes after that number
        ' Save the class data
        For n = 0 To Max_Classes
            If n > Val(Parse(1)) Then
                Class(n - 1).Name = Class(n).Name
                Class(n - 1).Sprite = Class(n).Sprite
                Class(n - 1).HP = Class(n).HP
                Class(n - 1).MP = Class(n).MP
                Class(n - 1).SP = Class(n).SP
                Class(n - 1).STR = Class(n).STR
                Class(n - 1).DEF = Class(n).DEF
                Class(n - 1).MAGI = Class(n).MAGI
                Class(n - 1).SPEED = Class(n).SPEED
                Class(n - 1).Map = Class(n).Map
                Class(n - 1).x = Class(n).x
                Class(n - 1).y = Class(n).y
            End If
        Next n
        
        ' Reduce the Max Class number by 1
        Max_Classes = Max_Classes - 1
        Debug.Print "MAX_CLASSES: " & Max_Classes
        ' If the Max Visible Classes happened to be set to all of them,
        ' reduce the total number by one
        If Max_Visible_Classes > Max_Classes Then Max_Visible_Classes = Max_Visible_Classes - 1
        
        'Redim the class array while keeping the old data
        ReDim Preserve Class(0 To Max_Classes) As ClassRec
        
        ' Print the full class INI
        Call PrintClassINI
        
        ' Resend all classes
        Call SendClassesToAll
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Create class packet ::
    ' :::::::::::::::::::::::::
    If (LCase$(Parse(0)) = "createclass") Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Increase the max classes by one
        Max_Classes = Max_Classes + 1
        
        ' Write the new max classes data
        Call PutVar(App.Path & "\data\classes.ini", "INIT", "MaxClasses", CStr(Max_Classes))
        
        ' Redim the class variable while keeping old data
        ReDim Preserve Class(0 To Max_Classes) As ClassRec
        
        'It's time to make the class basics
        Class(Max_Classes).Name = "Enter Name"
        Class(Max_Classes).Sprite = 0
        Class(Max_Classes).HP = 1
        Class(Max_Classes).MP = 0
        Class(Max_Classes).SP = 0
        Class(Max_Classes).STR = 1
        Class(Max_Classes).DEF = 1
        Class(Max_Classes).MAGI = 0
        Class(Max_Classes).SPEED = 1
        Class(Max_Classes).Map = 1
        Class(Max_Classes).x = 0
        Class(Max_Classes).y = 0
    
        ' Save it
        Call SendClassesToAll
        Call AddLog(GetPlayerName(Index) & " created class #" & n & ".", ADMIN_LOG)
        Call PrintClass(Max_Classes, True)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Set access packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "setaccess" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
            Call HackingAttempt(Index, "Trying to use powers not available")
            Exit Sub
        End If
        
        ' The index
        n = FindPlayer(Parse(1))
        ' The access
        i = Val(Parse(2))
        
        
        ' Check for invalid access level
        If i >= 0 And i <= 3 Then
            ' Check if player is on
            If n > 0 Then
                If GetPlayerAccess(n) <= 0 Then
                    Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
                End If
                
                Call SetPlayerAccess(n, i)
                Call SendPlayerData(n)
                Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            If i >= 4 And i <= 9 Then
                If n > 0 Then
                    If GetPlayerAccess(n) <= 0 Then
                        Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
                    End If
                    
                    Call SetPlayerAccess(n, i)
                    Call SendPlayerData(n)
                    Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
                Else
                    Call PlayerMsg(Index, "Player is not online.", White)
                End If
            Else
                Call PlayerMsg(Index, "Invalid access level.", Red)
            End If
        End If
                
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Who online packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "whosonline" Then
        Call SendWhosOnline(Index)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Set MOTD packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "setmotd" Then
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PutVar(App.Path & "\data\data.ini", "Strings", "MOTD", Parse(1))
        Call GlobalMsg("MOTD changed to: " & Parse(1), BrightCyan)
        Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Parse(1), ADMIN_LOG)
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "trade" Then
        If Map(Player(Index).Char(CharNum).Map).Shop > 0 Then
            Call SendTrade(Index, Map(GetPlayerMap(Index)).Shop)
        Else
            Call PlayerMsg(Index, "There is no shop here.", BrightRed)
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: Trade Get Item Data packet ::
    ' ::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "tradegetitem" Then
    ' Trade num
        n = Val(Parse(1))
        
        ' Prevent hacking
        If (n <= 0) Or (n > MAX_TRADES) Then
            Call HackingAttempt(Index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Index for shop
        i = Map(Player(Index).Char(Player(Index).CharNum).Map).Tile(GetPlayerX(Index), GetPlayerY(Index)).ShopNum
        
        'Note: I had to add in the player's item slots because they weren't set in the client
        '-smchronos
        '& SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data2) & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).Data2) & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).Data2) & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).Data2)
        'Call SendDataTo(Index, "TRADEGETITEM" & SEP_CHAR & CStr(Map(GetPlayerMap(Index)).Shop) & SEP_CHAR & Item(Shop(i).TradeItem(n).GetItem).Name & SEP_CHAR & Item(Shop(i).TradeItem(n).GetItem).Description & SEP_CHAR & Item(Shop(i).TradeItem(n).GiveItem).Name & SEP_CHAR & CStr(Shop(i).TradeItem(n).GiveItem) & SEP_CHAR & CStr(Shop(i).TradeItem(n).Stock) & SEP_CHAR & CStr(Item(Shop(i).TradeItem(n).GetItem).Type) & SEP_CHAR & CStr(Item(Shop(i).TradeItem(n).GetItem).Data2) & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data2) & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).Data2) & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).Data2) & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).Data2) & SEP_CHAR & END_CHAR)
        'Packet = "test" & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data2) & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).Data2) & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).Data2) & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).Data2) & SEP_CHAR & END_CHAR
        'CStr(Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).Data2) & SEP_CHAR & CStr(Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).Data2)
        Packet = "TRADEGETITEM" & SEP_CHAR & Trim$(CStr(Map(Player(Index).Char(Player(Index).CharNum).Map).Tile(GetPlayerX(Index), GetPlayerY(Index)).ShopNum)) & SEP_CHAR & Trim$(Item(Shop(i).TradeItem(n).GetItem).Name) & SEP_CHAR & Trim$(Item(Shop(i).TradeItem(n).GetItem).Description) & SEP_CHAR & Trim$(Item(Shop(i).TradeItem(n).GiveItem).Name) & SEP_CHAR & CStr(Shop(i).TradeItem(n).GiveValue) & SEP_CHAR & CStr(Shop(i).TradeItem(n).Stock) & SEP_CHAR & CStr(Item(Shop(i).TradeItem(n).GetItem).Type) & SEP_CHAR & CStr(Item(Shop(i).TradeItem(n).GetItem).Data2) & SEP_CHAR

        If GetPlayerWeaponSlot(Index) > 0 Then
            Packet = Packet & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data2 & SEP_CHAR
        Else
            Packet = Packet & 0 & SEP_CHAR
        End If
        
        If GetPlayerArmorSlot(Index) > 0 Then
            Packet = Packet & Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).Data2 & SEP_CHAR
        Else
            Packet = Packet & 0 & SEP_CHAR
        End If
        
        If GetPlayerHelmetSlot(Index) > 0 Then
            Packet = Packet & Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).Data2 & SEP_CHAR
        Else
            Packet = Packet & 0 & SEP_CHAR
        End If
        
        If GetPlayerShieldSlot(Index) > 0 Then
            Packet = Packet & Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).Data2 & SEP_CHAR
        Else
            Packet = Packet & 0 & SEP_CHAR
        End If
        
        Packet = Packet & END_CHAR
        
        Call SendDataTo(Index, Packet)
    End If

    ' ::::::::::::::::::::::::::
    ' :: Trade request packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "traderequest" Then
        ' Trade num
        n = Val(Parse(1))
        
        ' Prevent hacking
        If (n <= 0) Or (n > MAX_TRADES) Then
            Call HackingAttempt(Index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Index for shop
        i = Map(Player(Index).Char(Player(Index).CharNum).Map).Tile(GetPlayerX(Index), GetPlayerY(Index)).ShopNum
        
        ' Check if inv full
        x = FindOpenInvSlot(Index, Shop(i).TradeItem(n).GetItem)
        If x = 0 Then
            Call PlayerMsg(Index, "Trade unsuccessful, inventory full.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have the item
        If HasItem(Index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
            'Update stock (if 0, there is no item...if -1, it's considered infinite)
            If Shop(i).TradeItem(n).Stock > 0 Then
                Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
                Call GiveItem(Index, Shop(i).TradeItem(n).GetItem, Shop(i).TradeItem(n).GetValue)
            
                    Shop(i).TradeItem(n).Stock = Shop(i).TradeItem(n).Stock - 1
                '--------------------Sending packet update------------------- -smchronos
                    Packet = "TRADEGETITEM" & SEP_CHAR & Trim$(CStr(i)) & SEP_CHAR & Trim$(Item(Shop(i).TradeItem(n).GetItem).Name) & SEP_CHAR & Trim$(Item(Shop(i).TradeItem(n).GetItem).Description) & SEP_CHAR & Trim$(Item(Shop(i).TradeItem(n).GiveItem).Name) & SEP_CHAR & CStr(Shop(i).TradeItem(n).GiveItem) & SEP_CHAR & CStr(Shop(i).TradeItem(n).Stock) & SEP_CHAR & CStr(Item(Shop(i).TradeItem(n).GetItem).Type) & SEP_CHAR & CStr(Item(Shop(i).TradeItem(n).GetItem).Data2) & SEP_CHAR
        
                    If GetPlayerWeaponSlot(Index) > 0 Then
                        Packet = Packet & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data2 & SEP_CHAR
                    Else
                        Packet = Packet & 0 & SEP_CHAR
                    End If
        
                    If GetPlayerArmorSlot(Index) > 0 Then
                        Packet = Packet & Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).Data2 & SEP_CHAR
                    Else
                        Packet = Packet & 0 & SEP_CHAR
                    End If
        
                    If GetPlayerHelmetSlot(Index) > 0 Then
                        Packet = Packet & Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).Data2 & SEP_CHAR
                    Else
                        Packet = Packet & 0 & SEP_CHAR
                    End If
        
                    If GetPlayerShieldSlot(Index) > 0 Then
                        Packet = Packet & Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).Data2 & SEP_CHAR
                    Else
                        Packet = Packet & 0 & SEP_CHAR
                    End If
        
                    Packet = Packet & END_CHAR
        
                    Call SendDataTo(Index, Packet)
                '----------------------------------------------------------
            
                Call PlayerMsg(Index, "The trade was successful!", Yellow)
            ElseIf Shop(i).TradeItem(n).Stock = -1 Then
                Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
                Call GiveItem(Index, Shop(i).TradeItem(n).GetItem, Shop(i).TradeItem(n).GetValue)
                Call PlayerMsg(Index, "The trade was successful!", Yellow)
            Else
                Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
            End If
        Else
            Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
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
        i = Int(Item(GetPlayerInvItemNum(Index, n)).Data2 / 5)
        If i <= 0 Then i = 1
        
        DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(Index, n)
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
                Call SetPlayerInvItemDur(Index, n, Item(ItemNum).Data1)
                Call PlayerMsg(Index, "Item has been totally restored for " & GoldNeeded & " gold!", BrightBlue)
            Else
                ' They dont so restore as much as we can
                DurNeeded = (HasItem(Index, 1) / i)
                GoldNeeded = Int(DurNeeded * i / 2)
                If GoldNeeded <= 0 Then GoldNeeded = 1
                
                Call TakeItem(Index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(Index, n, GetPlayerInvItemDur(Index, n) + DurNeeded)
                Call PlayerMsg(Index, "Item has been partially fixed for " & GoldNeeded & " gold!", BrightBlue)
            End If
        Else
            Call PlayerMsg(Index, "Insufficient gold to fix this item!", BrightRed)
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
            If IsPlaying(i) And GetPlayerMap(Index) = GetPlayerMap(i) And GetPlayerX(i) = x And GetPlayerY(i) = y Then
                
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
        Next i
        
        ' Check for an item
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(GetPlayerMap(Index), i).Num > 0 Then
                If MapItem(GetPlayerMap(Index), i).x = x And MapItem(GetPlayerMap(Index), i).y = y Then
                    Call PlayerMsg(Index, "You see a " & Trim$(Item(MapItem(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next i
        
        ' Check for an npc
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(Index), i).Num > 0 Then
                If MapNpc(GetPlayerMap(Index), i).x = x And MapNpc(GetPlayerMap(Index), i).y = y Then
                    ' Change target
                    Player(Index).Target = i
                    Player(Index).TargetType = TARGET_TYPE_NPC
                    Call PlayerMsg(Index, "Your target is now a " & Trim$(Npc(MapNpc(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
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
        If n = Index Then
            Exit Sub
        End If
                
        ' Check for a previous party and if so drop it
        If Player(Index).InParty = YES Then
            Call PlayerMsg(Index, "You are already in a party!", Pink)
            Exit Sub
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
            
            ' Make sure they are in right level range
            If GetPlayerLevel(Index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(Index) - 5 > GetPlayerLevel(n) Then
                Call PlayerMsg(Index, "There is more then a 5 level gap between you two, party failed.", Pink)
                Exit Sub
            End If
            
            ' Check to see if player is already in a party
            If Player(n).InParty = NO Then
                Call PlayerMsg(Index, "Party request has been sent to " & GetPlayerName(n) & ".", Pink)
                Call PlayerMsg(n, GetPlayerName(Index) & " wants you to join their party.  Type /join to join, or /leave to decline.", Pink)
            
                Player(Index).PartyStarter = YES
                Player(Index).PartyPlayer = n
                Player(n).PartyPlayer = Index
            Else
                Call PlayerMsg(Index, "Player is already in a party!", Pink)
            End If
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Join party packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "joinparty" Then
        n = Player(Index).PartyPlayer
        
        If n > 0 Then
            ' Check to make sure they aren't the starter
            If Player(Index).PartyStarter = NO Then
                ' Check to make sure that each of there party players match
                If Player(n).PartyPlayer = Index Then
                    Call PlayerMsg(Index, "You have joined " & GetPlayerName(n) & "'s party!", Pink)
                    Call PlayerMsg(n, GetPlayerName(Index) & " has joined your party!", Pink)
                    
                    Player(Index).InParty = YES
                    Player(n).InParty = YES
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
    End If

    ' ::::::::::::::::::::::::
    ' :: Leave party packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "leaveparty" Then
        n = Player(Index).PartyPlayer
        
        If n > 0 Then
            If Player(Index).InParty = YES Then
                Call PlayerMsg(Index, "You have left the party.", Pink)
                Call PlayerMsg(n, GetPlayerName(Index) & " has left the party.", Pink)
                
                Player(Index).PartyPlayer = 0
                Player(Index).PartyStarter = NO
                Player(Index).InParty = NO
                Player(n).PartyPlayer = 0
                Player(n).PartyStarter = NO
                Player(n).InParty = NO
            Else
                Call PlayerMsg(Index, "Declined party request.", Pink)
                Call PlayerMsg(n, GetPlayerName(Index) & " declined your request.", Pink)
                
                Player(Index).PartyPlayer = 0
                Player(Index).PartyStarter = NO
                Player(Index).InParty = NO
                Player(n).PartyPlayer = 0
                Player(n).PartyStarter = NO
                Player(n).InParty = NO
            End If
        Else
            Call PlayerMsg(Index, "You are not in a party!", Pink)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "spells" Then
        Call SendPlayerSpells(Index)
        Exit Sub
    End If
    
    ' :::::::::::::::::
    ' :: Cast packet ::
    ' :::::::::::::::::
    If LCase(Parse(0)) = "cast" Then
        ' Spell slot
        n = Val(Parse(1))
        '-smchronos
        If Player(Index).Target > 0 Then
            If IsPlaying(Player(Index).Target) Then
                'Run Scripts
                For i = 0 To frmLibrary.lstLibrary.ListCount - 1
                    FileName = App.Path & "\Library\" & frmLibrary.lstLibrary.List(i)
                    If GetVar(FileName, "DATA", "Enabled") = "True" Then
                        MyScript.ExecuteStatement frmLibrary.lstLibrary.List(i), "SpellCast " & Index & "," & Player(Index).Target & "," & GetPlayerSpell(Index, n)
                    End If
                Next i
            End If
        End If
        Call CastSpell(Index, n)
        
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Location packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "requestlocation" Then
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PlayerMsg(Index, "Map: " & GetPlayerMap(Index) & ", X: " & GetPlayerX(Index) & ", Y: " & GetPlayerY(Index), Pink)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Remove Friend packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "removefriend" Then
        For n = 1 To MAX_FRIENDS
            If Player(Index).Char(Player(Index).CharNum).Friends(n) = Parse(1) Then
                Player(Index).Char(Player(Index).CharNum).Friends(n) = vbNullString
                Call SendFriends(Index)
                Exit Sub
            End If
        Next n
    End If
    
    ' :::::::::::::::::::::::
    ' :: Add Friend packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "addfriend" Then
        If GetPlayerName(Index) = Parse(1) Then
            Call PlayerMsg(Index, "You can't add yourself as a friend!", 4)
            Exit Sub
        End If
    
        For n = 1 To MAX_FRIENDS
            If Player(Index).Char(Player(Index).CharNum).Friends(n) = "" Then
                Player(Index).Char(Player(Index).CharNum).Friends(n) = Parse(1)
                Call SendFriends(Index)
                Exit Sub
            ElseIf Player(Index).Char(Player(Index).CharNum).Friends(n) = Parse(1) Then
                Call PlayerMsg(Index, Parse(1) & " is already your friend!", Red)
                Exit Sub
            ElseIf n = MAX_FRIENDS Then
                Call PlayerMsg(Index, "Too many friends!", Red)
                Exit Sub
            End If
        Next n
    End If
    
    'Get Info Packet
    If LCase(Parse(0)) = "getinfo" Then
        Packet = "SERVERINFO" & SEP_CHAR & GetVar(App.Path & "\Data\data.ini", "Strings", "Msg") & SEP_CHAR & END_CHAR
        Call SendDataTo(Index, Packet)
        
        Call TextAdd(frmServer.txtText, "Connection from " & GetPlayerIP(Index) & " has been terminated.", True)
        
        frmServer.Socket(Index).Close
    End If
    
    'send equip data
    If LCase(Parse(0)) = "getequipdata" Then
        Call SendEquipDataTo(Index)
        Exit Sub
    End If
    
    'send bank item data
    If LCase(Parse(0)) = "updatebankitem" Then
        n = CByte(Parse(1))
        Call SetPlayerBankItemNum(Index, n, Val(Parse(n + 1)))
        Call SetPlayerBankItemValue(Index, n, Val(Parse(n + 2)))
        Call SetPlayerBankItemDur(Index, n, Val(Parse(n + 3)))
        Call SendUpdateBankItemTo(Index, Val(Parse(n + 4)))
        Exit Sub
    End If
    
    'send whole bank
    If LCase(Parse(0)) = "getbank" Then
        Call SendBankInv(Index)
    End If
    
    'deposit from bank
    If LCase(Parse(0)) = "bankdeposit" Then
        InvSlot = FindOpenBankSlot(Index, Val(Parse(1)))
        Debug.Print "InvSlot: " & InvSlot
        If InvSlot <> 0 Then
            If Item(GetPlayerInvItemNum(Index, Val(Parse(3)))).Type = ITEM_TYPE_CURRENCY Then
                If GetPlayerInvItemValue(Index, Val(Parse(3))) >= Val(Parse(2)) Then
                    Call TakeItem(Index, Val(Parse(1)), Val(Parse(2)))
                    Call GiveBankItem(Index, Val(Parse(1)), Val(Parse(2)))
                    Exit Sub
                Else
                    Call PlayerMsg(Index, "You don't have that many " & Trim$(Item(Val(Parse(1))).Name) & " in your inventory!", Red)
                    Exit Sub
                End If
            Else
                Call TakeItem(Index, Val(Parse(1)), Val(Parse(2)))
                Call GiveBankItem(Index, Val(Parse(1)), Val(Parse(2)))
                Exit Sub
            End If
        Else
            Call PlayerMsg(Index, "Your bank is full!", Red)
            Exit Sub
        End If
    End If
    
    'withdraw from bank
    If LCase(Parse(0)) = "bankwithdraw" Then
        InvSlot = FindOpenInvSlot(Index, Val(Parse(1)))
        If InvSlot <> 0 Then
            If GetPlayerBankItemValue(Index, Val(Parse(3))) >= Val(Parse(2)) Then
                Call TakeBankItem(Index, Val(Parse(1)), Val(Parse(2)))
                Call GiveItem(Index, Val(Parse(1)), Val(Parse(2)))
                Exit Sub
            Else
                Call PlayerMsg(Index, "You don't have that many " & Trim$(Item(Val(Parse(1))).Name) & " in your bank!", Red)
                Exit Sub
            End If
        Else
            Call PlayerMsg(Index, "Your inventory is full!", Red)
            Exit Sub
        End If
    End If
    
    'Add tracker
    If LCase(Parse(0)) = "addtracker" Then
        If GetPlayerAccess(Index) < 1 Then
            Call AlertMsg(Index, "Tracking attempt through illegal means!")
            Exit Sub
        End If
    
        For n = 1 To MAX_TRACKERS
            If Player(FindPlayer(Parse(1))).Char(Player(FindPlayer(Parse(1))).CharNum).Trackers(n) = "" Then
                Player(FindPlayer(Parse(1))).Char(Player(FindPlayer(Parse(1))).CharNum).Trackers(n) = GetPlayerName(Index)
                Exit Sub
            End If
            If Err.Number <> 0 Then Exit Sub
        Next n
        Call PlayerMsg(Index, "There are too many people tracking this player!", Red)
    End If
    
    'remove tracker
    If LCase(Parse(0)) = "removetracker" Then
        If GetPlayerAccess(Index) < 1 Then
            Call AlertMsg(Index, "Tracking attempt through illegal means!")
            Exit Sub
        End If
    
        For n = 1 To MAX_TRACKERS
            If Player(FindPlayer(Parse(1))).Char(Player(FindPlayer(Parse(1))).CharNum).Trackers(n) = GetPlayerName(Index) Then
                Player(FindPlayer(Parse(1))).Char(Player(FindPlayer(Parse(1))).CharNum).Trackers(n) = ""
                Exit Sub
            End If
        Next n
        Call PlayerMsg(Index, "Error removing name!", Red)
    End If
    
    'check if the person is a mod or above and send their access data
    If LCase(Parse(0)) = "adminpanel" Then
        If GetPlayerAccess(Index) < 1 Then
            Call AlertMsg(Index, "Administration panel illegal access!")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "ADMINPANEL" & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & END_CHAR)
    End If
    
    'check if the person is a mapper or higher, if so, allow it
    If LCase$(Parse(0)) = "pausemap" Then
        If GetPlayerAccess(Index) < 2 Then
            Call AlertMsg(Index, "Illegal map command!")
            Exit Sub
        End If
        
        ' Check to see if we need to add a message
        If Trim$(Parse(2)) = END_CHAR Then
            Call SendDataToMap(GetPlayerMap(Index), "PAUSEMAP" & SEP_CHAR & Parse(1) & SEP_CHAR & END_CHAR)
        Else
            If Trim$(Parse(2)) <> "" Then
                Call SendDataToMap(GetPlayerMap(Index), "PAUSEMAP" & SEP_CHAR & Parse(1) & SEP_CHAR & Parse(2) & SEP_CHAR & END_CHAR)
            End If
        End If
    End If
    
    'used to change movement setting for everyone playing
    If LCase$(Parse(0)) = "mapextra" Then
        If GetPlayerAccess(Index) < 2 Then
            Call AlertMsg(Index, "Illegal map command!")
            Exit Sub
        End If
    
        Select Case LCase$(Trim$(Parse(1)))
        
        Case "allowmovement":
        Call SendDataToAllBut(Index, "MAPEXTRA" & SEP_CHAR & "ALLOWMOVEMENT" & SEP_CHAR & END_CHAR)
        Exit Sub
        
        Case "disallowmovement":
        Call SendDataToAllBut(Index, "MAPEXTRA" & SEP_CHAR & "DISALLOWMOVEMENT" & SEP_CHAR & END_CHAR)
        Exit Sub
        
        End Select
    End If
    
    'Update all inventory items
    If LCase$(Trim$(Parse(0))) = "updateallinv" Then
        Call SendInventory(Index)
    End If
    
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modHandleData.bas", "HandleData", Err.Number, Err.Description)
End Sub
