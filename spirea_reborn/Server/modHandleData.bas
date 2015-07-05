Attribute VB_Name = "modHandleData"
Option Explicit

Sub HandleData(ByVal index As Long, ByVal Data As String)
On Error Resume Next

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
        
            ' Check versions
            If Val(Parse(3)) < CLIENT_MAJOR Or Val(Parse(4)) < CLIENT_MINOR Or Val(Parse(5)) < CLIENT_REVISION Then
                Call AlertMsg(index, "Version outdated, please visit http://mirage.katami.com")
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
                
            ' Everything went ok
    
            ' Load the player
            Call LoadPlayer(index, Name)
            Call SendChars(index)
    
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
            Call TextAdd(frmServer.txtText, GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", True)
        End If
        Exit Sub
    End If

' :::::::::::::::::::::::::::
    ' :: Npc hp request packet ::
    ' :::::::::::::::::::::::::::
    Dim Pack As String
    If LCase(Parse(0)) = "getnpchp" Then

    Pack = "npchp" & SEP_CHAR
        For i = 1 To MAX_MAP_NPCS
            n = MapNpc(GetPlayerMap(index), i).Num
            Pack = Pack & MapNpc(GetPlayerMap(index), i).HP & SEP_CHAR & GetNpcMaxHP(n) & SEP_CHAR
        Next i
    Pack = Pack & END_CHAR

        Call SendDataTo(index, Pack)
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
            If LCase(Trim(Name)) = "consty" Then
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
        
        Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " says, '" & Msg & "'", PLAYER_LOG)
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " says, '" & Msg & "'", SayColor)
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
        Call SendDataToMapBut(index, GetPlayerMap(index), "PLAYERDIR" & SEP_CHAR & index & SEP_CHAR & GetPlayerDir(index) & END_CHAR)
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
            
            ' Find out what kind of item it is
            Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum <> GetPlayerArmorSlot(index) Then
                        If Int(GetPlayerDEF(index)) < n Then
                            Call PlayerMsg(index, "Your defense is to low to wear this armor!  Required DEF (" & n * 2 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerArmorSlot(index, InvNum)
                    Else
                        Call SetPlayerArmorSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum <> GetPlayerWeaponSlot(index) Then
                        If Int(GetPlayerSTR(index)) < n Then
                            Call PlayerMsg(index, "Your strength is to low to hold this weapon!  Required STR (" & n * 2 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerWeaponSlot(index, InvNum)
                    Else
                        Call SetPlayerWeaponSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
                        
                Case ITEM_TYPE_HELMET
                    If InvNum <> GetPlayerHelmetSlot(index) Then
                        If Int(GetPlayerSPEED(index)) < n Then
                            Call PlayerMsg(index, "Your speed coordination is to low to wear this helmet!  Required SPEED (" & n * 2 & ")", BrightRed)
                            Exit Sub
                        End If
                        Call SetPlayerHelmetSlot(index, InvNum)
                    Else
                        Call SetPlayerHelmetSlot(index, 0)
                    End If
                    Call SendWornEquipment(index)
            
                Case ITEM_TYPE_SHIELD
                    If InvNum <> GetPlayerShieldSlot(index) Then
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
                            
                            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                            Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
                            
                            ' Check if we are supposed to take away the item
                            If Map(GetPlayerMap(index)).Tile(x, y).Data2 = 1 Then
                                Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                Call PlayerMsg(index, "The key disolves.", Yellow)
                            End If
                        End If
                    End If
                    
                Case ITEM_TYPE_SPELL
                    ' Get the spell num
                    n = Item(GetPlayerInvItemNum(index, InvNum)).Data1
                    
                    If n > 0 Then
                        ' Make sure they are the right class
                        If Spell(n).ClassReq - 1 = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
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
                        Else
                            n = GetPlayerDamage(index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                            Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                            Call PlayerMsg(i, GetPlayerName(index) & " swings with enormous might!", BrightCyan)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(index, i, Damage)
                        Else
                            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
                        End If
                    Else
                        Call PlayerMsg(index, GetPlayerName(i) & "'s " & Trim(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked your hit!", BrightCyan)
                        Call PlayerMsg(i, "Your " & Trim(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).Name) & " has blocked " & GetPlayerName(index) & "'s hit!", BrightCyan)
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
                Else
                    n = GetPlayerDamage(index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(index), i).Num).DEF / 2)
                    Call PlayerMsg(index, "You feel a surge of energy upon swinging!", BrightCyan)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(index, i, Damage)
                Else
                    Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
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
            
            ' Everything is ok
            Select Case PointType
                Case 0
                    Call SetPlayerSTR(index, GetPlayerSTR(index) + 1)
                    Call PlayerMsg(index, "You have gained more strength!", White)
                Case 1
                    Call SetPlayerDEF(index, GetPlayerDEF(index) + 1)
                    Call PlayerMsg(index, "You have gained more defense!", White)
                Case 2
                    Call SetPlayerMAGI(index, GetPlayerMAGI(index) + 1)
                    Call PlayerMsg(index, "You have gained more magic abilities!", White)
                Case 3
                    Call SetPlayerSPEED(index, GetPlayerSPEED(index) + 1)
                    Call PlayerMsg(index, "You have gained more speed!", White)
            End Select
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
        Map(MapNum).Music = Val(Parse(n + 8))
        Map(MapNum).BootMap = Val(Parse(n + 9))
        Map(MapNum).BootX = Val(Parse(n + 10))
        Map(MapNum).BootY = Val(Parse(n + 11))
        Map(MapNum).Shop = Val(Parse(n + 12))
        
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

n = n + 13
Next x
Next y

For x = 1 To MAX_MAP_NPCS
            Map(MapNum).Npc(x) = Val(Parse(n))
            n = n + 1
            Call ClearMapNpc(x, MapNum)
        Next x
        Call SendMapNpcsToMap(MapNum)
        Call SpawnMapNpcs(MapNum)
        
        ' Save the map
        Call SaveMap(MapNum)
        
        ' Refresh map for everyone online
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
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
            Call SendDataTo(index, "MAPDONE" & END_CHAR)
        Else
            Call SendMapItemsTo(index, GetPlayerMap(index))
            Call SendMapNpcsTo(index, GetPlayerMap(index))
            Call SendJoinMap(index)
            Player(index).GettingMap = NO
            Call SendDataTo(index, "MAPDONE" & END_CHAR)
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
    If LCase(Parse(0)) = "mapreport" Then
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
        
        Call SendDataTo(index, "EDITMAP" & END_CHAR)
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
        
        Call SendDataTo(index, "ITEMEDITOR" & END_CHAR)
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
        
        Call SendDataTo(index, "NPCEDITOR" & END_CHAR)
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
        Npc(n).DropChance = Val(Parse(8))
        Npc(n).DropItem = Val(Parse(9))
        Npc(n).DropItemValue = Val(Parse(10))
        Npc(n).STR = Val(Parse(11))
        Npc(n).DEF = Val(Parse(12))
        Npc(n).SPEED = Val(Parse(13))
        Npc(n).MAGI = Val(Parse(14))

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
        
        Call SendDataTo(index, "SHOPEDITOR" & END_CHAR)
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
        
        Call SendDataTo(index, "SPELLEDITOR" & END_CHAR)
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
            Call PlayerMsg(index, "Invalid access level.", Red)
        End If
                
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Who online packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "whosonline" Then
        Call SendWhosOnline(index)
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
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "trade" Then
        If Map(GetPlayerMap(index)).Shop > 0 Then
            Call SendTrade(index, Map(GetPlayerMap(index)).Shop)
        Else
            Call PlayerMsg(index, "There is no shop here.", BrightRed)
        End If
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
        i = Map(GetPlayerMap(index)).Shop
        
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
End Sub

