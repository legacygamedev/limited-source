Attribute VB_Name = "modHandleData"
Option Explicit

Sub HandleData(ByVal Index As Long, ByVal Data As String)
On Error Resume Next

Dim Parse() As String
Dim sType As String
Dim HDSerial As String
Dim HDModel As String
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
Dim I As Long, N As Long, X As Long, Y As Long, F As Long
Dim lTO As Long
Dim lFROM As Long
Dim MapNum As Long
Dim s As String
Dim tMapStart As Long, tMapEnd As Long
Dim ShopNum As Long, ItemNum As Long
Dim DurNeeded As Long, GoldNeeded As Long
        
Debug.Print "R: " & Data
        
    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
        
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Requesting classes for making a character ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "getclasses" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If
    
        If Not IsPlaying(Index) Then
            Call SendNewCharClasses(Index)
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::::
    ' :: New account packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "newaccount" Then
            If UBound(Parse) <> 5 Then
                Call HackingAttempt(Index, ".")
            End If

        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
            HDModel = Decrypt(DBKEY, Trim(Parse(3)))
            HDSerial = Decrypt(DBKEY, Trim(Parse(4)))
        
            'Check serials, if they are blank then they are not allowed to perform action.
            If HDModel = ENC_ERR Or HDSerial = ENC_ERR Then
                Call AlertMsg(Index, "Access Denied")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            ' Prevent hacking
            For I = 1 To Len(Name)
                N = Asc(Mid(Name, I, 1))
                
                If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                Else
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next I
            
            If SerialsExist(HDModel, HDSerial) Then
                Call AlertMsg(Index, "Computer already bound.")
                Exit Sub
            End If
            
            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password, HDModel, HDSerial)
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
            If UBound(Parse) <> 5 Then
                Call HackingAttempt(Index, ".")
            End If

        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
            HDModel = Decrypt(DBKEY, Trim(Parse(3)))
            HDSerial = Decrypt(DBKEY, Trim(Parse(4)))
            
            'Check serials, if they are blank then they are not allowed to perform action.
            If HDModel = ENC_ERR Or HDSerial = ENC_ERR Then
                Call AlertMsg(Index, "Access Denied")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
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
                        
            If Not SerialsOK(Name, HDModel, HDSerial) Then
                Call AlertMsg(Index, "Access Denied")
                Exit Sub
            End If
                        
            ' Delete names from master name file
            Call LoadPlayer(Index, Name)
            For I = 1 To MAX_CHARS
                If Trim(Player(Index).Char(I).Name) <> "" Then
                    Call DeleteChar(Player(Index).Char(I).FKey)
                End If
            Next I
            Call DeleteAccount(Player(Index).FKey)
            Call ClearPlayer(Index)
            
            ' Everything went ok
            Call AddLog("Account " & Trim(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(Index, "Your account has been deleted.")
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::
    ' :: Login packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "login" Then
        If UBound(Parse) <> 8 Then
            Call HackingAttempt(Index, ".")
        End If

        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Decrypt(DBKEY, Parse(1))
            Password = Decrypt(DBKEY, Parse(2))
            HDModel = Decrypt(DBKEY, Parse(6))
            HDSerial = Decrypt(DBKEY, Parse(7))
            
            'Check serials, if they are blank then they are not allowed to perform action.
            If HDModel = ENC_ERR Or HDSerial = ENC_ERR Then
                Call AlertMsg(Index, "Access Denied")
                Exit Sub
            End If
            
            ' Check versions
            If Val(Parse(3)) <> CLIENT_MAJOR Or Val(Parse(4)) <> CLIENT_MINOR Or Val(Parse(5)) <> CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit http://www.mirageuniverse.com")
                Exit Sub
            End If
            
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            'Check to see if login or password have the / or \, if so, it's a hack and call hacking attempt.
            If InStr(Name, "/") > 0 Or InStr(Name, "\") > 0 Or InStr(Password, "/") > 0 Or InStr(Password, "\") > 0 Then
                Call AlertMsg(Index, "Invalid login/password")
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
        
            'If logging on two different accounts from the same computer, should not happen as the client will
            'soon not allow it
            'SHANN NOTE: Add client to make sure it will not launch twice at the same time.
            If IsMultiComputers(HDModel, HDSerial) Then
                Call AlertMsg(Index, "Multiple computer logins is not authorized.")
                Exit Sub
            End If
        
            If IsMultiAccounts(Name) Then
                Call AlertMsg(Index, "Multiple account logins is not authorized.")
                Exit Sub
            End If
            
            'For first time loginers (after bind update), their db HD fields would be purely blank. so we check for blank
            'serials, if both fields are blank (a must), then we populate them before any checkings :)
            If SerialsBlank(Name) Then
                'Since there are no serials, we then need to use current serials for replacements..
                Call SaveSerials(Name, HDModel, HDSerial)
            Else
                'Serials are currently only used for del account, add/del characters. and primary character setting
                'and use.
                'Will give option to bind an account to a computer, if person wants, via primary character.
            End If

            'Check Cache against local copy :)
            
            ' Everything went ok
    
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
        If UBound(Parse) <> 7 Then
            Call HackingAttempt(Index, ".")
        End If

        If Not IsPlaying(Index) Then
            Name = Parse(1)
            Sex = Val(Parse(2))
            Class = Val(Parse(3))
            CharNum = Val(Parse(4))
            HDModel = Decrypt(DBKEY, Trim(Parse(5)))
            HDSerial = Decrypt(DBKEY, Trim(Parse(6)))
        
            'Check serials, if they are blank then they are not allowed to perform action.
            If HDModel = ENC_ERR Or HDSerial = ENC_ERR Then
                Call AlertMsg(Index, "Access Denied")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Then
                Call AlertMsg(Index, "Character name must be at least three characters in length.")
                Exit Sub
            End If
            
            ' Prevent being me
            'If LCase(Trim(Name)) = "consty" Then
            '    Call AlertMsg(Index, "Lets get one thing straight, you are not me, ok? :)")
            '    Exit Sub
            'End If
            
            ' Prevent hacking
            For I = 1 To Len(Name)
                N = Asc(Mid(Name, I, 1))
                
                If (N >= 65 And N <= 90) Or (N >= 97 And N <= 122) Or (N = 95) Or (N = 32) Or (N >= 48 And N <= 57) Then
                Else
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next I
                                    
                                    
            'Check to see if serials ok, remember
            'Remember that you can only create/del chars on bound computers
            If Not SerialsOK(Name, HDModel, HDSerial) Then
                Call AlertMsg(Index, "Access Denied")
                Exit Sub
            End If
            
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Prevent hacking
            If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
                Call HackingAttempt(Index, "Invalid Gender")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Class < 0 Or Class > Max_Classes Then
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
            Call AddChar(Index, Name, Sex, Class, CharNum)
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
        If UBound(Parse) <> 4 Then
            Call HackingAttempt(Index, ".")
        End If
            
        If Not IsPlaying(Index) Then
            CharNum = Val(Parse(1))
            HDModel = Decrypt(DBKEY, Trim(Parse(2)))
            HDSerial = Decrypt(DBKEY, Trim(Parse(3)))
            
            'Check serials, if they are blank then they are not allowed to perform action.
            If HDModel = ENC_ERR Or HDSerial = ENC_ERR Then
                Call AlertMsg(Index, "Access Denied")
                Exit Sub
            End If
            
            'Check to see if serials ok, remember
            'Remember that you can only create/del chars on bound computers
            If Not SerialsOK(Name, HDModel, HDSerial) Then
                Call AlertMsg(Index, "Access Denied")
                Exit Sub
            End If
            
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
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

        If Not IsPlaying(Index) Then
            CharNum = Val(Parse(1))
        
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Check to make sure the character exists and if so, set it as its current char
            If CharExist(Index, CharNum) Then
                Player(Index).CharNum = CharNum
                Call JoinGame(Index)
            
                CharNum = Player(Index).CharNum
                Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " (" & GetPlayerIP(Index) & ") has began playing " & GAME_NAME & ".", PLAYER_LOG)
                Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " (" & GetPlayerIP(Index) & ") has began playing " & GAME_NAME & ".", True)
                Call UpdateCaption

            Else
                Call AlertMsg(Index, "Character does not exist!")
            End If
        End If
        Exit Sub
    End If

'Simple check before saying stuff or anything else, if they do this, they are hacking so ban them, because we are nice :)
If Trim(Player(Index).Login) = "" Then
    Call BanIndex(Index, 0)
    Exit Sub
End If

    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "saymsg" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

        Msg = Trim(Parse(1))
        If Msg = "" Then Exit Sub
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Say Text Modification")
                Exit Sub
            End If
        Next I
        
        Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " says, '" & Msg & "'", SayColor)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "emotemsg" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

        Msg = Trim(Parse(1))
        If Msg = "" Then Exit Sub
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Emote Text Modification")
                Exit Sub
            End If
        Next I
        
        Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Msg, EmoteColor)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "broadcastmsg" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
        Msg = Trim(Parse(1))
        If Msg = "" Then Exit Sub
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Broadcast Text Modification")
                Exit Sub
            End If
        Next I
        
        s = GetPlayerName(Index) & ": " & Msg
        Call AddLog(s, PLAYER_LOG)
        Call GlobalMsg(s, BroadcastColor)
        Call TextAdd(frmServer.txtText, s, True)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "globalmsg" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

        Msg = Trim(Parse(1))
        If Msg = "" Then Exit Sub
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Global Text Modification")
                Exit Sub
            End If
        Next I
        
        If GetPlayerAccess(Index) > 0 Then
            s = "(global) " & GetPlayerName(Index) & ": " & Msg
            Call AddLog(s, ADMIN_LOG)
            Call GlobalMsg(s, GlobalColor)
            Call TextAdd(frmServer.txtText, s, True)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "canassign" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If

        'Debug.Print "CANASSIGN PACKET IN!"

        If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
            'Create Staff list / string
            Name = GetStaffListString()
            'Debug.Print "STAFF LISTING: " & Name
            'Send confirmation w/ list :)
            Call SendDataTo(Index, "ASSIGNOK" & SEP_CHAR & Name)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "assign" Then
        If UBound(Parse) <> 6 Then
            Call HackingAttempt(Index, ".")
        End If
        
        'Debug.Print Data
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
            Exit Sub
        End If
        sType = Trim(Parse(1))
        lFROM = CLng(Parse(2))
        lTO = CLng(Parse(3))
        Name = CStr(Parse(4))
        F = CLng(Parse(5))
        If F = 0 Then
            N = GetCharFKeyByName(Name)
        Else
            N = 0
        End If
        Debug.Print "acct char name: " & Name
        Debug.Print "acct char fkey: " & N
        Select Case LCase(sType)
            Case "items"
                If N > 0 Then
                    Call PlayerMsg(Index, "*** Assigning items #" & lFROM & " thru #" & lTO & " to " & Name & ".", Yellow)
                Else
                    Call PlayerMsg(Index, "*** Unassigning items #" & lFROM & " thru #" & lTO & ".", Yellow)
                End If
                For I = lFROM To lTO
                    Item(I).Assigned = N
                Next I
                Call PlayerMsg(Index, "*** Saving changed items.", Yellow)
                Call SaveItemArray(lFROM, lTO)
                Call PlayerMsg(Index, "*** Items Saved.", Yellow)
            
            Case "maps"
            
            Case "npcs"
            
            Case "spells"
            
            Case "shops"
            
        End Select
            
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "adminmsg" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
        
        Msg = Trim(Parse(1))
        If Msg = "" Then Exit Sub
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Admin Text Modification")
                Exit Sub
            End If
        Next I
        
        If GetPlayerAccess(Index) > 0 Then
            Call AddLog("(admin " & GetPlayerName(Index) & ") " & Msg, ADMIN_LOG)
            Call AdminMsg("(admin " & GetPlayerName(Index) & ") " & Msg, AdminColor)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "playermsg" Then
        If UBound(Parse) <> 3 Then
            Call HackingAttempt(Index, ".")
        End If
        
        MsgTo = FindPlayer(Parse(1))
        Msg = Trim(Parse(2))
        If Msg = "" Then Exit Sub
        
        ' Prevent hacking
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call HackingAttempt(Index, "Player Msg Text Modification")
                Exit Sub
            End If
        Next I
        
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
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playermove" And Player(Index).GettingMap = NO Then
        If UBound(Parse) <> 3 Then
            Call HackingAttempt(Index, ".")
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
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerdir" And Player(Index).GettingMap = NO Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
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
    End If
        
    ' :::::::::::::::::::::
    ' :: Use item packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "useitem" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
        
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
                
        ' If item is disabled, we cannot use.. but we should be able to drop :D
        If Item(GetPlayerInvItemNum(Index, InvNum)).Disabled = 1 Then
            Call PlayerMsg(Index, "*** Unable to use item, item disabled.", Yellow)
            Exit Sub
        End If
        
        If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
            N = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
            
            ' Find out what kind of item it is
            Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum <> GetPlayerArmorSlot(Index) Then
                        If Int(GetPlayerDEF(Index)) < N Then
                            Call PlayerMsg(Index, "Your defense is to low to wear this armor!  Required DEF (" & N & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerArmorSlot(Index, InvNum)
                    Else
                        Call SetPlayerArmorSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum <> GetPlayerWeaponSlot(Index) Then
                        If Int(GetPlayerSTR(Index)) < N Then
                            Call PlayerMsg(Index, "Your strength is to low to hold this weapon!  Required STR (" & N & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerWeaponSlot(Index, InvNum)
                    Else
                        Call SetPlayerWeaponSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
                        
                Case ITEM_TYPE_HELMET
                    If InvNum <> GetPlayerHelmetSlot(Index) Then
                        If Int(GetPlayerSPEED(Index)) < N Then
                            Call PlayerMsg(Index, "Your speed coordination is to low to wear this helmet!  Required SPEED (" & N & ")", BrightRed)
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
                
                Case ITEM_TYPE_POTIONADDMP
                    Call SetPlayerMP(Index, GetPlayerMP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendMP(Index)
        
                Case ITEM_TYPE_POTIONADDSP
                    Call SetPlayerSP(Index, GetPlayerSP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendSP(Index)

                Case ITEM_TYPE_POTIONSUBHP
                    Call SetPlayerHP(Index, GetPlayerHP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendHP(Index)
                
                Case ITEM_TYPE_POTIONSUBMP
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendMP(Index)
        
                Case ITEM_TYPE_POTIONSUBSP
                    Call SetPlayerSP(Index, GetPlayerSP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendSP(Index)
                    
                Case ITEM_TYPE_KEY
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            If GetPlayerY(Index) > 0 Then
                                X = GetPlayerX(Index)
                                Y = GetPlayerY(Index) - 1
                            Else
                                Exit Sub
                            End If
                            
                        Case DIR_DOWN
                            If GetPlayerY(Index) < MAX_MAPY Then
                                X = GetPlayerX(Index)
                                Y = GetPlayerY(Index) + 1
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_LEFT
                            If GetPlayerX(Index) > 0 Then
                                X = GetPlayerX(Index) - 1
                                Y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_RIGHT
                            If GetPlayerX(Index) < MAX_MAPY Then
                                X = GetPlayerX(Index) + 1
                                Y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                    End Select
                    
                    ' Check if a key exists
                    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY Then
                        ' Check if the key they are using matches the map key
                        If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(X, Y).Data1 Then
                            TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            
                            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
                            
                            ' Check if we are supposed to take away the item
                            If Map(GetPlayerMap(Index)).Tile(X, Y).Data2 = 1 Then
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                Call PlayerMsg(Index, "The key disolves.", Yellow)
                            End If
                        End If
                    End If
                    
                Case ITEM_TYPE_SPELL
                    ' Get the spell num
                    N = Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                    
                    If N > 0 Then
                        ' Make sure they are the right class
                        If Spell(N).ClassReq - 1 = GetPlayerClass(Index) Or Spell(N).ClassReq = 0 Then
                            ' Make sure they are the right level
                            I = GetSpellReqLevel(Index, N)
                            If I <= GetPlayerLevel(Index) Then
                                I = FindOpenSpellSlot(Index)
                                
                                ' Make sure they have an open spell slot
                                If I > 0 Then
                                    ' Make sure they dont already have the spell
                                    If Not HasSpell(Index, N) Then
                                        Call SetPlayerSpell(Index, I, N)
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
                                Call PlayerMsg(Index, "You must be level " & I & " to learn this spell.", White)
                            End If
                        Else
                            Call PlayerMsg(Index, "This spell can only be learned by a " & GetClassName(Spell(N).ClassReq - 1) & ".", White)
                        End If
                    Else
                        Call PlayerMsg(Index, "This scroll is not connected to a spell, please inform an admin!", White)
                    End If
                    
            End Select
        End If
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "attack" Then
        'If UBound(Parse) <> 1 Then
        '    Call HackingAttempt(Index, ".")
        'End If
        
        ' Try to attack a player
        For I = 1 To MAX_PLAYERS
            ' Make sure we dont try to attack ourselves
            If I <> Index Then
                ' Can we attack the player?
                If CanAttackPlayer(Index, I) Then
                    If Not CanPlayerBlockHit(I) Then
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(Index) Then
                            Damage = GetPlayerDamage(Index) - GetPlayerProtection(I)
                        Else
                            N = GetPlayerDamage(Index)
                            Damage = N + Int(Rnd * Int(N / 2)) + 1 - GetPlayerProtection(I)
                            Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                            Call PlayerMsg(I, GetPlayerName(Index) & " swings with enormous might!", BrightCyan)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(Index, I, Damage)
                        Else
                            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                        End If
                    Else
                        Call PlayerMsg(Index, GetPlayerName(I) & "'s " & Trim(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked your hit!", BrightCyan)
                        Call PlayerMsg(I, "Your " & Trim(Item(GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))).Name) & " has blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan)
                    End If
                    
                    Exit Sub
                End If
            End If
        Next I
        
        ' Try to attack a npc
        For I = 1 To MAX_MAP_NPCS
            ' Can we attack the npc?
            If CanAttackNpc(Index, I) Then
                ' Get the damage we can do
                If Not CanPlayerCriticalHit(Index) Then
                    Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), I).Num).DEF / 2)
                Else
                    N = GetPlayerDamage(Index)
                    Damage = N + Int(Rnd * Int(N / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), I).Num).DEF / 2)
                    Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(Index, I, Damage)
                Else
                    Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                End If
                Exit Sub
            End If
        Next I
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Use stats packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "usestatpoint" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
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
            
            ' Everything is ok
            Select Case PointType
                Case 0
                    Call SetPlayerSTR(Index, GetPlayerSTR(Index) + 1)
                    Call PlayerMsg(Index, "You have gained more strength!", White)
                Case 1
                    Call SetPlayerDEF(Index, GetPlayerDEF(Index) + 1)
                    Call PlayerMsg(Index, "You have gained more defense!", White)
                Case 2
                    Call SetPlayerMAGI(Index, GetPlayerMAGI(Index) + 1)
                    Call PlayerMsg(Index, "You have gained more magic abilities!", White)
                Case 3
                    Call SetPlayerSPEED(Index, GetPlayerSPEED(Index) + 1)
                    Call PlayerMsg(Index, "You have gained more speed!", White)
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
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
        Name = Parse(1)
        
        I = FindPlayer(Name)
        If I > 0 Then
            Call PlayerMsg(Index, "Account: " & Trim(Player(I).Login) & ", Name: " & GetPlayerName(I), BrightGreen)
            If GetPlayerAccess(Index) > ADMIN_MONITER Then
                Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(I) & " -=-", BrightGreen)
                Call PlayerMsg(Index, "Level: " & GetPlayerLevel(I) & "  Exp: " & GetPlayerExp(I) & "/" & GetPlayerNextLevel(I), BrightGreen)
                Call PlayerMsg(Index, "HP: " & GetPlayerHP(I) & "/" & GetPlayerMaxHP(I) & "  MP: " & GetPlayerMP(I) & "/" & GetPlayerMaxMP(I) & "  SP: " & GetPlayerSP(I) & "/" & GetPlayerMaxSP(I), BrightGreen)
                Call PlayerMsg(Index, "STR: " & GetPlayerSTR(I) & "  DEF: " & GetPlayerDEF(I) & "  MAGI: " & GetPlayerMAGI(I) & "  SPEED: " & GetPlayerSPEED(I), BrightGreen)
                N = Int(GetPlayerSTR(I) / 2) + Int(GetPlayerLevel(I) / 2)
                I = Int(GetPlayerDEF(I) / 2) + Int(GetPlayerLevel(I) / 2)
                If N > 100 Then N = 100
                If I > 100 Then I = 100
                Call PlayerMsg(Index, "Critical Hit Chance: " & N & "%, Block Chance: " & I & "%", BrightGreen)
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
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        N = FindPlayer(Parse(1))
        
        If N <> Index Then
            If N > 0 Then
                Call PlayerWarp(Index, GetPlayerMap(N), GetPlayerX(N), GetPlayerY(N))
                Call PlayerMsg(N, GetPlayerName(Index) & " has warped to you.", BrightBlue)
                Call PlayerMsg(Index, "You have been warped to " & GetPlayerName(N) & ".", BrightBlue)
                Call AddLog(GetPlayerName(Index) & " has warped to " & GetPlayerName(N) & ", map #" & GetPlayerMap(N) & ".", ADMIN_LOG)
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
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The player
        N = FindPlayer(Parse(1))
        
        If N <> Index Then
            If N > 0 Then
                Call PlayerWarp(N, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                Call PlayerMsg(N, "You have been summoned by " & GetPlayerName(Index) & ".", BrightBlue)
                Call PlayerMsg(Index, GetPlayerName(N) & " has been summoned.", BrightBlue)
                Call AddLog(GetPlayerName(Index) & " has warped " & GetPlayerName(N) & " to self, map #" & GetPlayerMap(Index) & ".", ADMIN_LOG)
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
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The map
        N = Val(Parse(1))
        
        ' Prevent hacking
        If N < 0 Or N > MAX_MAPS Then
            Call HackingAttempt(Index, "Invalid map")
            Exit Sub
        End If
        
        Call PlayerWarp(Index, N, GetPlayerX(Index), GetPlayerY(Index))
        Call PlayerMsg(Index, "You have been warped to map #" & N, BrightBlue)
        Call AddLog(GetPlayerName(Index) & " warped to map #" & N & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Set sprite packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "setsprite" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If
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
    End If
                
    ' ::::::::::::::::::::::::::
    ' :: Stats request packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "getstats" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If
        Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(Index) & " -=-", White)
        Call PlayerMsg(Index, "Level: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), White)
        Call PlayerMsg(Index, "HP: " & GetPlayerHP(Index) & "/" & GetPlayerMaxHP(Index) & "  MP: " & GetPlayerMP(Index) & "/" & GetPlayerMaxMP(Index) & "  SP: " & GetPlayerSP(Index) & "/" & GetPlayerMaxSP(Index), White)
        Call PlayerMsg(Index, "STR: " & GetPlayerSTR(Index) & "  DEF: " & GetPlayerDEF(Index) & "  MAGI: " & GetPlayerMAGI(Index) & "  SPEED: " & GetPlayerSPEED(Index), White)
        N = Int(GetPlayerSTR(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
        I = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
        If N > 100 Then N = 100
        If I > 100 Then I = 100
        Call PlayerMsg(Index, "Critical Hit Chance: " & N & "%, Block Chance: " & I & "%", White)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player request for a new map ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requestnewmap" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
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
        
        N = 1
        
        MapNum = GetPlayerMap(Index)
        Map(MapNum).Name = Parse(N + 1)
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        Map(MapNum).Moral = Val(Parse(N + 3))
        Map(MapNum).Up = Val(Parse(N + 4))
        Map(MapNum).Down = Val(Parse(N + 5))
        Map(MapNum).Left = Val(Parse(N + 6))
        Map(MapNum).Right = Val(Parse(N + 7))
        Map(MapNum).Music = Val(Parse(N + 8))
        Map(MapNum).BootMap = Val(Parse(N + 9))
        Map(MapNum).BootX = Val(Parse(N + 10))
        Map(MapNum).BootY = Val(Parse(N + 11))
        Map(MapNum).Shop = Val(Parse(N + 12))
        
        N = N + 13
        
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(MapNum).Tile(X, Y).Ground = Val(Parse(N))
                Map(MapNum).Tile(X, Y).Mask = Val(Parse(N + 1))
                Map(MapNum).Tile(X, Y).Anim = Val(Parse(N + 2))
                Map(MapNum).Tile(X, Y).Fringe = Val(Parse(N + 3))
                Map(MapNum).Tile(X, Y).Type = Val(Parse(N + 4))
                Map(MapNum).Tile(X, Y).Data1 = Val(Parse(N + 5))
                Map(MapNum).Tile(X, Y).Data2 = Val(Parse(N + 6))
                Map(MapNum).Tile(X, Y).Data3 = Val(Parse(N + 7))
                
                N = N + 8
            Next X
        Next Y
        
        For X = 1 To MAX_MAP_NPCS
            Map(MapNum).Npc(X) = Val(Parse(N))
            N = N + 1
            Call ClearMapNpc(X, MapNum)
        Next X
        Call SendMapNpcsToMap(MapNum)
        Call SpawnMapNpcs(MapNum)
        
        ' Save the map
        Call SaveMap(MapNum)
        
        ' Refresh map for everyone online
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
                Call PlayerWarp(I, MapNum, GetPlayerX(I), GetPlayerY(I))
            End If
        Next I
        
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Need map yes/no packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "needmap" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
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
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If

        Call PlayerMapGetItem(Index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to drop something packet ::
    ' ::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapdropitem" Then
        If UBound(Parse) <> 3 Then
            Call HackingAttempt(Index, ".")
        End If
        InvNum = Val(Parse(1))
        Ammount = Val(Parse(2))
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_INV Then
            Call HackingAttempt(Index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Ammount > GetPlayerInvItemValue(Index, InvNum) Then
            Call HackingAttempt(Index, "Item amount modification")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
            If Ammount <= 0 Then
                Call HackingAttempt(Index, "Trying to drop 0 amount of currency")
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
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If

        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Clear out it all
        For I = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), I).X, MapItem(GetPlayerMap(Index), I).Y)
            Call ClearMapItem(I, GetPlayerMap(Index))
        Next I
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(Index))
        
        ' Respawn NPCS
        For I = 1 To MAX_MAP_NPCS
            Call SpawnNpc(I, GetPlayerMap(Index))
        Next I
        
        Call PlayerMsg(Index, "Map respawned.", Blue)
        Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Map report packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "mapreport" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If

        s = "Free Maps: "
        tMapStart = 1
        tMapEnd = 1
        
        For I = 1 To MAX_MAPS
            If Trim(Map(I).Name) = "" Then
                tMapEnd = tMapEnd + 1
            Else
                If tMapEnd - tMapStart > 0 Then
                    s = s & Trim(str(tMapStart)) & "-" & Trim(str(tMapEnd - 1)) & ", "
                End If
                tMapStart = I + 1
                tMapEnd = I + 1
            End If
        Next I
        
        s = s & Trim(str(tMapStart)) & "-" & Trim(str(tMapEnd - 1)) & ", "
        s = Mid(s, 1, Len(s) - 2)
        s = s & "."
        
        Call PlayerMsg(Index, s, Brown)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Kick player packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "kickplayer" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
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
    End If
        
    ' :::::::::::::::::::::
    ' :: Ban list packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "banlist" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendBanList(Index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Ban destroy packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "bandestroy" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call Kill(App.Path & "\banlist.txt")
        Call PlayerMsg(Index, "Ban list destroyed.", White)
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::
    ' :: Ban player packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "banplayer" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MONITER Then
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
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Request edit map packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditmap" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If
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
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        'If item is locked, cannot edit, unless an owner
        
        
        Call SendDataTo(Index, "ITEMEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit item packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "edititem" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If
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
    End If
    
    ' ::::::::::::::::::::::
    ' :: Save item packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "saveitem" Then
        If UBound(Parse) <> 11 Then
            Call HackingAttempt(Index, ".")
        End If
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
        
        ' Check to see if item is locked, cannot edit locked item :)
        If Item(N).Locked = 1 And GetPlayerAccess(Index) < ADMIN_CREATOR Then
            Call PlayerMsg(Index, "*** Cannot edit locked item.", Yellow)
            Exit Sub
        End If
        
        'If they are not assigned to item, then do not let them edit item..
        If GetPlayerAccess(Index) < ADMIN_CREATOR And Item(N).Assigned <> GetPlayerFKey(Index) Then
            Call PlayerMsg(Index, "*** Access Denied. You are not assigned to this item.", Yellow)
            Exit Sub
        End If
        
        ' Update the item
        Item(N).Name = Parse(2)
        Item(N).Pic = Val(Parse(3))
        Item(N).Type = Val(Parse(4))
        Item(N).Data1 = Val(Parse(5))
        Item(N).Data2 = Val(Parse(6))
        Item(N).Data3 = Val(Parse(7))
        
        Item(N).Unbreakable = Val(Parse(8))
        
        If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
            Item(N).Locked = Val(Parse(9))
            Item(N).Disabled = Val(Parse(10))
        End If
        
        ' Save it
        Call SendUpdateItemToAll(N)
        Call SaveItem(N)
        Call PlayerMsg(Index, "*** Item #" & N & " " & Item(N).Name & " have been saved.", Yellow)
        Call AddLog(GetPlayerName(Index) & " saved item #" & N & ".", ADMIN_LOG)
        'Recreate item cache
        Call CreateCache(1)
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::
    ' :: Request edit npc packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditnpc" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If
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
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

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
    End If
    
    ' :::::::::::::::::::::
    ' :: Save npc packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "savenpc" Then
        If UBound(Parse) <> 15 Then
            Call HackingAttempt(Index, ".")
        End If

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
        Npc(N).DropChance = Val(Parse(8))
        Npc(N).DropItem = Val(Parse(9))
        Npc(N).DropItemValue = Val(Parse(10))
        Npc(N).str = Val(Parse(11))
        Npc(N).DEF = Val(Parse(12))
        Npc(N).SPEED = Val(Parse(13))
        Npc(N).MAGI = Val(Parse(14))
        
        ' Save it
        Call SendUpdateNpcToAll(N)
        Call SaveNpc(N)
        Call AddLog(GetPlayerName(Index) & " saved npc #" & N & ".", ADMIN_LOG)
        'Recreate npc cache
        Call CreateCache(2)
        Exit Sub
    End If
            
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit shop packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditshop" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If

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
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

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
        
        N = 6
        For I = 1 To MAX_TRADES
            Shop(ShopNum).TradeItem(I).GiveItem = Val(Parse(N))
            Shop(ShopNum).TradeItem(I).GiveValue = Val(Parse(N + 1))
            Shop(ShopNum).TradeItem(I).GetItem = Val(Parse(N + 2))
            Shop(ShopNum).TradeItem(I).GetValue = Val(Parse(N + 3))
            N = N + 4
        Next I
        
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
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If

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
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

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
    End If
    
    ' :::::::::::::::::::::::
    ' :: Save spell packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "savespell") Then
        If UBound(Parse) <> 9 Then
            Call HackingAttempt(Index, ".")
        End If

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
                
        ' Save it
        Call SendUpdateSpellToAll(N)
        Call SaveSpell(N)
        Call AddLog(GetPlayerName(Index) & " saving spell #" & N & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Set access packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "setaccess" Then
        If UBound(Parse) <> 3 Then
            Call HackingAttempt(Index, ".")
        End If

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
            Call PlayerMsg(Index, "Invalid access level.", Red)
        End If
                
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Who online packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "whosonline" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If

        Call SendWhosOnline(Index)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Set MOTD packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "setmotd" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SetMOTD(Trim(Parse(1)))
        Call GlobalMsg("MOTD changed to: " & Trim(Parse(1)), BrightCyan)
        Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Trim(Parse(1)), ADMIN_LOG)
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "trade" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If

        If Map(GetPlayerMap(Index)).Shop > 0 Then
            Call SendTrade(Index, Map(GetPlayerMap(Index)).Shop)
        Else
            Call PlayerMsg(Index, "There is no shop here.", BrightRed)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Trade request packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "traderequest" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

        ' Trade num
        N = Val(Parse(1))
        
        ' Prevent hacking
        If (N <= 0) Or (N > MAX_TRADES) Then
            Call HackingAttempt(Index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Index for shop
        I = Map(GetPlayerMap(Index)).Shop
        
        ' Check if inv full
        X = FindOpenInvSlot(Index, Shop(I).TradeItem(N).GetItem)
        If X = 0 Then
            Call PlayerMsg(Index, "Trade unsuccessful, backpack full.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have the item
        If HasItem(Index, Shop(I).TradeItem(N).GiveItem) >= Shop(I).TradeItem(N).GiveValue Then
            Call TakeItem(Index, Shop(I).TradeItem(N).GiveItem, Shop(I).TradeItem(N).GiveValue)
            Call GiveItem(Index, Shop(I).TradeItem(N).GetItem, Shop(I).TradeItem(N).GetValue)
            Call PlayerMsg(Index, "The trade was successful!", Yellow)
        Else
            Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Fix item packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "fixitem" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

        ' Inv num
        N = Val(Parse(1))
        
        'Make sure they are in a map with a shop
        If Map(GetPlayerMap(Index)).Shop = 0 Then
            Exit Sub
        End If
        
        'if there is a map with a shop, make sure it can fix items
        If Shop(Map(GetPlayerMap(Index)).Shop).FixesItems = False Then
            Exit Sub
        End If

        
        ' Make sure its a equipable item
        If Item(GetPlayerInvItemNum(Index, N)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(Index, N)).Type > ITEM_TYPE_SHIELD Then
            Call PlayerMsg(Index, "You can only fix weapons, armors, helmets, and shields.", BrightRed)
            Exit Sub
        End If
        
        ' Check if they have a full inventory
        If FindOpenInvSlot(Index, GetPlayerInvItemNum(Index, N)) <= 0 Then
            Call PlayerMsg(Index, "You have no backpack space left!", BrightRed)
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
                Call SetPlayerInvItemDur(Index, N, Item(ItemNum).Data1)
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
    End If

    ' :::::::::::::::::::
    ' :: Search packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "search" Then
        If UBound(Parse) <> 3 Then
            Call HackingAttempt(Index, ".")
        End If

        X = Val(Parse(1))
        Y = Val(Parse(2))
        
        ' Prevent subscript out of range
        If X < 0 Or X > MAX_MAPX Or Y < 0 Or Y > MAX_MAPY Then
            Exit Sub
        End If
        
        ' Check for a player
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) And GetPlayerMap(Index) = GetPlayerMap(I) And GetPlayerX(I) = X And GetPlayerY(I) = Y Then
                
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
        Next I
        
        ' Check for an item
        For I = 1 To MAX_MAP_ITEMS
            If MapItem(GetPlayerMap(Index), I).Num > 0 Then
                If MapItem(GetPlayerMap(Index), I).X = X And MapItem(GetPlayerMap(Index), I).Y = Y Then
                    Call PlayerMsg(Index, "You see a " & Trim(Item(MapItem(GetPlayerMap(Index), I).Num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next I
        
        ' Check for an npc
        For I = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(Index), I).Num > 0 Then
                If MapNpc(GetPlayerMap(Index), I).X = X And MapNpc(GetPlayerMap(Index), I).Y = Y Then
                    ' Change target
                    Player(Index).Target = I
                    Player(Index).TargetType = TARGET_TYPE_NPC
                    Call PlayerMsg(Index, "Your target is now a " & Trim(Npc(MapNpc(GetPlayerMap(Index), I).Num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        Next I
        
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Party packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "party" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

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
    End If

    ' :::::::::::::::::::::::
    ' :: Join party packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "joinparty" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If

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
    End If

    ' ::::::::::::::::::::::::
    ' :: Leave party packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "leaveparty" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If

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
    End If
    
    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "spells" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If

        Call SendPlayerSpells(Index)
        Exit Sub
    End If
    
    ' :::::::::::::::::
    ' :: Cast packet ::
    ' :::::::::::::::::
    If LCase(Parse(0)) = "cast" Then
        If UBound(Parse) <> 2 Then
            Call HackingAttempt(Index, ".")
        End If

        ' Spell slot
        N = Val(Parse(1))
        
        Call CastSpell(Index, N)
        
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Location packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "requestlocation" Then
        If UBound(Parse) <> 1 Then
            Call HackingAttempt(Index, ".")
        End If

        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call PlayerMsg(Index, "Map: " & GetPlayerMap(Index) & ", X: " & GetPlayerX(Index) & ", Y: " & GetPlayerY(Index), Pink)
        Exit Sub
    End If
    
    'It must be something else, so ban them
    'Call HackingAttempt(Index, ".")

End Sub

