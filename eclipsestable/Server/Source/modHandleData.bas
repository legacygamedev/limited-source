Attribute VB_Name = "modHandleData"
Option Explicit

Sub HandleData(ByVal index As Long, ByVal Data As String)
    Dim Parse() As String

    On Error Resume Next

    Parse = Split(Data, SEP_CHAR)

    Select Case LCase$(Parse(0))
        Case "requesteditmain"
            Call Packet_RequestEditMain(index, Parse(1))
            Exit Sub
            
        Case "newmain"
            Call Packet_NewMain(index, Parse(1), Parse(2))
            Exit Sub
            
        Case "guildmsg"
            Call Packet_GuildMsg(index, Parse(1))
            Exit Sub

        Case "getclasses"
            Call Packet_GetClasses(index)
            Exit Sub
    
        Case "newaccount"
            Call Packet_NewAccount(index, Parse(1), Parse(2), Parse(3))
            Exit Sub
            
        Case "delaccount"
            Call Packet_DeleteAccount(index, Parse(1), Parse(2))
            Exit Sub
    
        Case "acclogin"
            Call Packet_AccountLogin(index, Parse(1), Parse(2), Val(Parse(3)), Val(Parse(4)), Val(Parse(5)), Parse(6))
            Exit Sub
    
        Case "givemethemax"
            Call Packet_GiveMeTheMax(index)
            Exit Sub
    
        Case "addchar"
            Call Packet_AddCharacter(index, Parse(1), Val(Parse(2)), Val(Parse(3)), Val(Parse(4)), Val(Parse(5)), Val(Parse(6)), Val(Parse(7)))
            Exit Sub
    
        Case "delchar"
            Call Packet_DeleteCharacter(index, Val(Parse(1)))
            Exit Sub
    
        Case "usechar"
            Call Packet_UseCharacter(index, Val(Parse(1)))
            Exit Sub

        Case "guildchangeaccess"
            Call Packet_GuildChangeAccess(index, Parse(1), Val(Parse(2)))
            Exit Sub

        Case "guilddisown"
            Call Packet_GuildDisown(index, Parse(1))
            Exit Sub

        Case "guildleave"
            Call Packet_GuildLeave(index)
            Exit Sub

        Case "guildmake"
            Call Packet_GuildMake(index, Parse(1), Parse(2))
            Exit Sub

        Case "guildmember"
            Call Packet_GuildMember(index, Parse(1))
            Exit Sub

        Case "guildtrainee"
            Call Packet_GuildTrainee(index, Parse(1))
            Exit Sub

        Case "saymsg"
            Call Packet_SayMessage(index, Parse(1))
            Exit Sub

        Case "emotemsg"
            Call Packet_EmoteMessage(index, Parse(1))
            Exit Sub

        Case "broadcastmsg"
            Call Packet_BroadcastMessage(index, Parse(1))
            Exit Sub

        Case "globalmsg"
            Call Packet_GlobalMessage(index, Parse(1))
            Exit Sub

        Case "adminmsg"
            Call Packet_AdminMessage(index, Parse(1))
            Exit Sub

        Case "playermsg"
            Call Packet_PlayerMessage(index, Parse(1), Parse(2))
            Exit Sub

        Case "playermove"
            Call Packet_PlayerMove(index, Val(Parse(1)), Val(Parse(2)), Val(Parse(3)), Val(Parse(4)))
            Exit Sub

        Case "playerdir"
            Call Packet_PlayerDirection(index, Val(Parse(1)))
            Exit Sub

        Case "useitem"
            Call Packet_UseItem(index, Val(Parse(1)))
            Exit Sub

        Case "playermovemouse"
            Call Packet_PlayerMoveMouse(index, Val(Parse(1)))
            Exit Sub

        Case "warp"
            Call Packet_Warp(index, Val(Parse(1)))
            Exit Sub

        Case "endshot"
            Call Packet_EndShot(index, Val(Parse(1)))
            Exit Sub

        Case "attack"
            Call Packet_Attack(index)
            Exit Sub

        Case "usestatpoint"
            Call Packet_UseStatPoint(index, Val(Parse(1)))
            Exit Sub

        Case "setplayersprite"
            Call Packet_SetPlayerSprite(index, Parse(1), Val(Parse(2)))
            Exit Sub

        Case "getstats"
            Call Packet_GetStats(index, Parse(1))
            Exit Sub

        Case "requestnewmap"
            Call Packet_RequestNewMap(index, Val(Parse(1)))
            Exit Sub

        Case "warpmeto"
            Call Packet_WarpMeTo(index, Parse(1))
            Exit Sub

        Case "warptome"
            Call Packet_WarpToMe(index, Parse(1))
            Exit Sub

        Case "mapdata"
            Call Packet_MapData(index, Parse)
            Exit Sub

        Case "needmap"
            Call Packet_NeedMap(index, Parse(1))
            Exit Sub

        Case "mapgetitem"
            Call Packet_MapGetItem(index)
            Exit Sub
            
        Case "mapdropitem"
            Call Packet_MapDropItem(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "maprespawn"
            Call Packet_MapRespawn(index)
            Exit Sub

        Case "kickplayer"
            Call Packet_KickPlayer(index, Parse(1))
            Exit Sub

        Case "banlist"
            Call Packet_BanList(index)
            Exit Sub

        Case "bandestroy"
            Call Packet_BanListDestroy(index)
            Exit Sub

        Case "banplayer"
            Call Packet_BanPlayer(index, Parse(1))
            Exit Sub

        Case "requesteditmap"
            Call Packet_RequestEditMap(index)
            Exit Sub

        Case "requestedititem"
            Call Packet_RequestEditItem(index)
            Exit Sub

        Case "edititem"
            Call Packet_EditItem(index, Val(Parse(1)))
            Exit Sub

        Case "saveitem"
            Call Packet_SaveItem(index, Parse)
            Exit Sub

        Case "enabledaynight"
            Call Packet_EnableDayNight(index)
            Exit Sub

        Case "daynight"
            Call Packet_DayNight(index)
            Exit Sub

        Case "requesteditnpc"
            Call Packet_RequestEditNPC(index)
            Exit Sub

        Case "editnpc"
            Call Packet_EditNPC(index, Val(Parse(1)))
            Exit Sub

        Case "savenpc"
            Call Packet_SaveNPC(index, Parse)
            Exit Sub

        Case "requesteditshop"
            Call Packet_RequestEditShop(index)
            Exit Sub

        Case "editshop"
            Call Packet_EditShop(index, Val(Parse(1)))
            Exit Sub

        Case "saveshop"
            Call Packet_SaveShop(index, Parse)
            Exit Sub

        Case "requesteditspell"
            Call Packet_RequestEditSpell(index)
            Exit Sub

        Case "editspell"
            Call Packet_EditSpell(index, Val(Parse(1)))
            Exit Sub

        Case "savespell"
            Call Packet_SaveSpell(index, Parse)
            Exit Sub

        Case "forgetspell"
            Call Packet_ForgetSpell(index, Val(Parse(1)))
            Exit Sub

        Case "setaccess"
            Call Packet_SetAccess(index, Parse(1), Val(Parse(2)))
            Exit Sub

        Case "whosonline"
            Call Packet_WhoIsOnline(index)
            Exit Sub

        Case "onlinelist"
            Call Packet_OnlineList(index)
            Exit Sub

        Case "setmotd"
            Call Packet_SetMOTD(index, Parse(1))
            Exit Sub

        Case "buy"
            Call Packet_BuyItem(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "sellitem"
            Call Packet_SellItem(index, Val(Parse(1)), Val(Parse(2)), Val(Parse(3)), Val(Parse(4)))
            Exit Sub

        Case "fixitem"
            Call Packet_FixItem(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "search"
            Call Packet_Search(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub
        
        Case "search2"
            Call Packet_Search2(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "playerchat"
            Call Packet_PlayerChat(index, Parse(1))
            Exit Sub

        Case "achat"
            Call Packet_AcceptChat(index)
            Exit Sub

        Case "dchat"
            Call Packet_DenyChat(index)
            Exit Sub

        Case "qchat"
            Call Packet_QuitChat(index)
            Exit Sub

        Case "sendchat"
            Call Packet_SendChat(index, Parse(1))
            Exit Sub

        Case "pptrade"
            Call Packet_PrepareTrade(index, Parse(1))
            Exit Sub

        Case "atrade"
            Call Packet_AcceptTrade(index)
            Exit Sub

        Case "qtrade"
            Call Packet_QuitTrade(index)
            Exit Sub

        Case "dtrade"
            Call Packet_DenyTrade(index)
            Exit Sub

        Case "updatetradeinv"
            Call Packet_UpdateTradeInventory(index, Val(Parse(1)), Val(Parse(2)), Parse(3), Val(Parse(4)))
            Exit Sub

        Case "swapitems"
            Call Packet_SwapItems(index)
            Exit Sub

        Case "party"
            Call Packet_Party(index, Parse(1))
            Exit Sub

        Case "joinparty"
            Call Packet_JoinParty(index)
            Exit Sub

        Case "leaveparty"
            Call Packet_LeaveParty(index)
            Exit Sub

        Case "partychat"
            Call Packet_PartyChat(index, Parse(1))
            Exit Sub

        Case "spells"
            Call Packet_Spells(index)
            Exit Sub

        Case "hotscript"
            Call Packet_HotScript(index, Val(Parse(1)))
            Exit Sub

        Case "scripttile"
            Call Packet_ScriptTile(index, Val(Parse(1)))
            Exit Sub

        Case "cast"
            Call Packet_Cast(index, Val(Parse(1)))
            Exit Sub

        Case "refresh"
            Call Packet_Refresh(index)
            Exit Sub

        Case "buysprite"
            Call Packet_BuySprite(index)
            Exit Sub

        Case "clearowner"
            Call Packet_ClearOwner(index)
            Exit Sub

        Case "requestedithouse"
            Call Packet_RequestEditHouse(index)
            Exit Sub

        Case "buyhouse"
            Call Packet_BuyHouse(index)
            Exit Sub
            
        Case "sellhouse"
            Call Packet_SellHouse(index)
            Exit Sub

        Case "checkcommands"
            Call Packet_CheckCommands(index, Parse(1))
            Exit Sub

        Case "prompt"
            Call Packet_Prompt(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "querybox"
            Call Packet_QueryBox(index, Parse(1), Val(Parse(2)))
            Exit Sub

        Case "requesteditarrow"
            Call Packet_RequestEditArrow(index)
            Exit Sub

        Case "editarrow"
            Call Packet_EditArrow(index, Val(Parse(1)))
            Exit Sub

        Case "savearrow"
            Call Packet_SaveArrow(index, Val(Parse(1)), Parse(2), Val(Parse(3)), Val(Parse(4)), Val(Parse(5)))
            Exit Sub

        Case "checkarrows"
            Call Packet_CheckArrows(index, Val(Parse(1)))
            Exit Sub

        Case "requesteditemoticon"
            Call Packet_RequestEditEmoticon(index)
            Exit Sub

        Case "requesteditelement"
            Call Packet_RequestEditElement(index)
            Exit Sub

        Case "requesteditquest"
            Call Packet_RequestEditQuest(index)
            Exit Sub

        Case "editemoticon"
            Call Packet_EditEmoticon(index, Val(Parse(1)))
            Exit Sub

        Case "editelement"
            Call Packet_EditElement(index, Val(Parse(1)))
            Exit Sub

        Case "saveemoticon"
            Call Packet_SaveEmoticon(index, Val(Parse(1)), Parse(2), Val(Parse(3)))
            Exit Sub

        Case "saveelement"
            Call Packet_SaveElement(index, Val(Parse(1)), Parse(2), Val(Parse(3)), Val(Parse(4)))
            Exit Sub

        Case "checkemoticons"
            Call Packet_CheckEmoticon(index, Val(Parse(1)))
            Exit Sub

        Case "mapreport"
            Call Packet_MapReport(index)
            Exit Sub

        Case "gmtime"
            Call Packet_GMTime(index, Val(Parse(1)))
            Exit Sub

        Case "weather"
            Call Packet_Weather(index, Val(Parse(1)))
            Exit Sub

        Case "warpto"
            Call Packet_WarpTo(index, Val(Parse(1)), Val(Parse(2)), Val(Parse(3)))
            Exit Sub

        Case "localwarp"
            Call Packet_LocalWarp(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "arrowhit"
            Call Packet_ArrowHit(index, Val(Parse(1)), Val(Parse(2)), Val(Parse(3)), Val(Parse(4)))
            Exit Sub

        Case "bankdeposit"
            Call Packet_BankDeposit(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "bankwithdraw"
            Call Packet_BankWithdraw(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "reloadscripts"
            Call Packet_ReloadScripts(index)
            Exit Sub

        Case "custommenuclick"
            Call Packet_CustomMenuClick(index, Val(Parse(1)), Val(Parse(2)), Parse(3), Val(Parse(4)), Parse(5))
            Exit Sub

        Case "returningcustomboxmsg"
            Call Packet_CustomBoxReturnMsg(index, Val(Parse(1)))
            Exit Sub
            
        Case "getonline"
            Call Packet_GetWhosOnline(index)
            Exit Sub
        
                Case "sendemail"
            Call Packet_SendEmail(index, Parse(1), Parse(2), Parse(3), Parse(4))
            Exit Sub
            
        Case "getmsgbody"
            Call Packet_EmailBody(index, Parse(1), Parse(2), Parse(3))
            Exit Sub
            
        Case "dltmsg"
            Call Packet_RemoveMail(index, Parse(1), Parse(2), Parse(3))
            Exit Sub
            
        Case "updateinbox"
            Call Packet_MyInbox(index, Parse(1), Parse(2))
            Exit Sub
            
        Case "checkchar"
            Call Packet_CheckChar(index, Parse(1), Parse(2), Parse(3), Parse(4))
            Exit Sub

    End Select

    Call HackingAttempt(index, "Received invalid packet: " & Parse(0))
End Sub

Public Sub Packet_GetWhosOnline(ByVal index As String)
    Call SendDataTo(index, "totalonline" & SEP_CHAR & TotalOnlinePlayers & SEP_CHAR & END_CHAR)
End Sub

Public Sub Packet_GetClasses(ByVal index As Long)
    Call SendNewCharClasses(index)
End Sub

Public Sub Packet_NewAccount(ByVal index As Long, ByVal Username As String, ByVal Password As String, ByVal Email As String)
    If Not IsLoggedIn(index) Then
        If LenB(Username) < 6 Then
            Call PlainMsg(index, "Your username must be at least three characters in length.", 1)
            Exit Sub
        End If

        If LenB(Password) < 6 Then
            Call PlainMsg(index, "Your password must be at least three characters in length.", 1)
            Exit Sub
        End If

        If EMAIL_AUTH = 1 Then
            If LenB(Email) = 0 Then
                Call PlainMsg(index, "Your email address cannot be blank.", 1)
                Exit Sub
            End If
        End If

        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(index, "Your username must consist of alpha-numeric characters!", 1)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(index, "Your password must consist of alpha-numeric characters!", 1)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call AddAccount(index, Username, Password, Email)
            Call PlainMsg(index, "Your account has been created!", 0)
        Else
            Call PlainMsg(index, "Sorry, that account name is already taken!", 1)
        End If
    End If
End Sub

Public Sub Packet_DeleteAccount(ByVal index As Long, ByVal Username As String, ByVal Password As String)
    Dim I As Long
    
    If Not IsLoggedIn(index) Then
        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(index, "Your username must consist of alpha-numeric characters!", 2)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(index, "Your password must consist of alpha-numeric characters!", 2)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call PlainMsg(index, "That account name does not exist.", 2)
            Exit Sub
        End If

        If Not PasswordOK(Username, Password) Then
            Call PlainMsg(index, "You've entered an incorrect password.", 2)
            Exit Sub
        End If
    
        Call LoadPlayer(index, Username)
        For I = 1 To MAX_CHARS
            If LenB(Trim$(Player(index).Char(I).Name)) <> 0 Then
                Call DeleteName(Player(index).Char(I).Name)
            End If
        Next I
        Call ClearPlayer(index)

        ' Remove the users main player profile.
        Kill App.Path & "\Accounts\" & Username & "_Info.ini"
        Kill App.Path & "\Accounts\" & Username & "\*.*"

        ' Delete the users account directory.
        RmDir App.Path & "\Accounts\" & Username & "\"
    
        Call PlainMsg(index, "Your account has been deleted.", 0)
    End If
End Sub

Public Sub Packet_AccountLogin(ByVal index As Long, ByVal Username As String, ByVal Password As String, ByVal Major As Long, ByVal Minor As Long, ByVal Revision As Long, ByVal Code As String)
    If Not IsLoggedIn(index) Then
        ' I'll re-add this when I change it to the new DAT method. [Mellowz]
        'If ACC_VERIFY = 1 Then
        '    If Val(ReadINI("GENERAL", "verified", App.Path & "\Accounts\" & Trim$(Player(Index).Login) & ".ini")) = 0 Then
        '        Call PlainMsg(Index, "Your account hasn't been verified yet!", 3)
        '        Exit Sub
        '    End If
        'End If

        If Major < CLIENT_MAJOR Or Minor < CLIENT_MINOR Or Revision < CLIENT_REVISION Then
            'Call PlainMsg(Index, "Version out-dated. Please visit " & Trim$(GetVar(App.Path & "\Data.ini", "CONFIG", "WebSite")), 3)
            'Exit Sub
        End If

        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(index, "Your username must consist of alpha-numeric characters!", 3)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(index, "Your password must consist of alpha-numeric characters!", 3)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call PlainMsg(index, "That account name does not exist.", 3)
            Exit Sub
        End If
    
        If Not PasswordOK(Username, Password) Then
            Call PlainMsg(index, "You've entered an incorrect password.", 3)
            Exit Sub
        End If
    
        If IsMultiAccounts(Username) Then
            Call PlainMsg(index, "Multiple account logins is not authorized.", 3)
            Exit Sub
        End If
    
        If frmServer.Closed.Value = Checked Then
            Call PlainMsg(index, "The server is closed at the moment!", 3)
            Exit Sub
        End If
    
        If Code <> SEC_CODE Then
            Call AlertMsg(index, "The client password does not match the server password.")
            Exit Sub
        End If
    
        Call LoadPlayer(index, Username)
        Call SendChars(index)
    
        Call TextAdd(frmServer.txtText(0), GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", True)
    End If
End Sub

Public Sub Packet_GiveMeTheMax(ByVal index As Long)
    Dim packet As String
    Dim WalkFix As Long
    Dim GameVersion As String
    
    WalkFix = GetVar(App.Path & "\Data.ini", "Config", "WalkFix")
    GameVersion = GetVar(App.Path & "\Data.ini", "Config", "GameVersion")

    packet = "MAXINFO" & SEP_CHAR
    packet = packet & GAME_NAME & SEP_CHAR
    packet = packet & MAX_PLAYERS & SEP_CHAR
    packet = packet & MAX_ITEMS & SEP_CHAR
    packet = packet & MAX_NPCS & SEP_CHAR
    packet = packet & MAX_SHOPS & SEP_CHAR
    packet = packet & MAX_SPELLS & SEP_CHAR
    packet = packet & MAX_MAPS & SEP_CHAR
    packet = packet & MAX_MAP_ITEMS & SEP_CHAR
    packet = packet & MAX_MAPX & SEP_CHAR
    packet = packet & MAX_MAPY & SEP_CHAR
    packet = packet & MAX_EMOTICONS & SEP_CHAR
    packet = packet & MAX_ELEMENTS & SEP_CHAR
    packet = packet & PAPERDOLL & SEP_CHAR
    packet = packet & SPRITESIZE & SEP_CHAR
    packet = packet & MAX_SCRIPTSPELLS & SEP_CHAR
    packet = packet & CUSTOM_SPRITE & SEP_CHAR
    packet = packet & LEVEL & SEP_CHAR
    packet = packet & MAX_PARTY_MEMBERS & SEP_CHAR
    packet = packet & STAT1 & SEP_CHAR
    packet = packet & STAT2 & SEP_CHAR
    packet = packet & STAT3 & SEP_CHAR
    packet = packet & STAT4 & SEP_CHAR
    packet = packet & WalkFix & SEP_CHAR
    packet = packet & GameVersion & SEP_CHAR
    packet = packet & END_CHAR

    Call SendDataTo(index, packet)
    Call SendNewsTo(index)
End Sub

Public Sub Packet_AddCharacter(ByVal index As Long, ByVal Name As String, ByVal Sex As Long, ByVal Class As Long, ByVal CharNum As Long, ByVal Head As Long, ByVal Body As Long, ByVal Leg As Long)
    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(index, "Invalid CharNum")
        Exit Sub
    End If
    
    If LenB(Name) < 6 Then
        Call HackingAttempt(index, "Invalid Name Length")
        Exit Sub
    End If
    
    If Sex <> SEX_MALE And Sex <> SEX_FEMALE Then
        Call HackingAttempt(index, "Invalid Sex")
        Exit Sub
    End If
    
    If Class < 0 Or Class > MAX_CLASSES Then
        Call HackingAttempt(index, "Invalid Class")
        Exit Sub
    End If

    If Not IsAlphaNumeric(Name) Then
        Call PlainMsg(index, "Your username must consist of alpha-numeric characters!", 4)
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

    Call AddChar(index, Name, Sex, Class, CharNum, Head, Body, Leg)

    Call SendChars(index)

    Call PlainMsg(index, "Character has been created!", 5)

    If SCRIPTING = 1 Then
        Call MyScript.ExecuteStatement("Scripts\main.ess", "OnNewChar " & index & "," & CharNum)
    End If
End Sub

Public Sub Packet_DeleteCharacter(ByVal index As Long, ByVal CharNum As Long)
    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(index, "Invalid CharNum")
        Exit Sub
    End If
    
    If CharExist(index, CharNum) Then
        Call DelChar(index, CharNum)
        Call SendChars(index)
    
        Call PlainMsg(index, "Character has been deleted!", 5)
    Else
        Call PlainMsg(index, "Character does not exist!", 5)
    End If
End Sub

Public Sub Packet_UseCharacter(ByVal index As Long, ByVal CharNum As Long)
    Dim FileID As Integer

    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(index, "Invalid CharNum")
        Exit Sub
    End If
    
    If CharExist(index, CharNum) Then
        Player(index).CharNum = CharNum
    
        If frmServer.GMOnly.Value = Checked Then
            If GetPlayerAccess(index) = 0 Then
                Call PlainMsg(index, "The server is only available to GMs at the moment!", 5)
                Exit Sub
            End If
        End If
    
        Call JoinGame(index)

        Call TextAdd(frmServer.txtText(0), GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & GAME_NAME & ".", True)
        Call UpdateTOP
    
        If Not FindChar(GetPlayerName(index)) Then
            FileID = FreeFile
            Open App.Path & "\Accounts\CharList.txt" For Append As #FileID
                Print #FileID, GetPlayerName(index)
            Close #FileID
        End If
    Else
        Call PlainMsg(index, "Character does not exist!", 5)
    End If
End Sub

Public Sub Packet_GuildChangeAccess(ByVal index As Long, ByVal Name As String, ByVal Rank As Long)
    Dim NameIndex As Long
    
    If LenB(Name) = 0 Then
        Call PlayerMsg(index, "You must enter a player name to proceed.", WHITE)
        Exit Sub
    End If

    If Rank < 0 Or Rank > 4 Then
        Call PlayerMsg(index, "You must provide a valid rank to proceed.", RED)
        Exit Sub
    End If

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If GetPlayerGuild(NameIndex) <> GetPlayerGuild(index) Then
        Call PlayerMsg(index, Name & " is not in your guild.", RED)
        Exit Sub
    End If

    If GetPlayerGuildAccess(index) < 4 Then
        Call PlayerMsg(index, "You are not the owner of this guild.", RED)
        Exit Sub
    End If

    Call SetPlayerGuildAccess(NameIndex, Rank)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_GuildDisown(ByVal index As Long, ByVal Name As String)
    Dim NameIndex As Long

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If GetPlayerGuild(NameIndex) <> GetPlayerGuild(index) Then
        Call PlayerMsg(index, Name & " is not in your guild.", RED)
        Exit Sub
    End If

    If GetPlayerGuildAccess(NameIndex) > GetPlayerGuildAccess(index) Then
        Call PlayerMsg(index, Name & " has a higher guild level than you.", RED)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, vbNullString)
    Call SetPlayerGuildAccess(NameIndex, 0)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_GuildLeave(ByVal index As Long)
    If LenB(GetPlayerGuild(index)) = 0 Then
        Call PlayerMsg(index, "You are not in a guild.", RED)
        Exit Sub
    End If

    Call SetPlayerGuild(index, vbNullString)
    Call SetPlayerGuildAccess(index, 0)
    Call SendPlayerData(index)
End Sub

Public Sub Packet_GuildMake(ByVal index As Long, ByVal Name As String, ByVal Guild As String)
    Dim NameIndex As Long

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If LenB(GetPlayerGuild(NameIndex)) <> 0 Then
        Call PlayerMsg(index, Name & " is already in a guild.", RED)
        Exit Sub
    End If

    If LenB(Guild) = 0 Then
        Call PlayerMsg(index, "Please enter a valid guild name.", RED)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, Guild)
    Call SetPlayerGuildAccess(NameIndex, 4)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_GuildMember(ByVal index As Long, ByVal Name As String)
    Dim NameIndex As Long

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If GetPlayerGuild(NameIndex) <> GetPlayerGuild(index) Then
        Call PlayerMsg(index, Name & " is not in your guild.", RED)
        Exit Sub
    End If

    If GetPlayerGuildAccess(NameIndex) > 1 Then
        Call PlayerMsg(index, Name & " has already been admitted.", WHITE)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, GetPlayerGuild(index))
    Call SetPlayerGuildAccess(NameIndex, 1)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_GuildTrainee(ByVal index As Long, ByVal Name As String)
    Dim NameIndex As Long

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If LenB(GetPlayerGuild(NameIndex)) <> 0 Then
        Call PlayerMsg(index, Name & " is already in a guild.", RED)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, GetPlayerGuild(index))
    Call SetPlayerGuildAccess(NameIndex, 0)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_SayMessage(ByVal index As Long, ByVal Message As String)
    If frmServer.chkLogMap.Value = Unchecked Then
        If GetPlayerAccess(index) = 0 Then
            Call PlayerMsg(index, "Map messages have been disabled by the server!", BRIGHTRED)
            Exit Sub
        End If
    End If

    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & ": " & Message, SayColor)
    Call MapMsg2(GetPlayerMap(index), Message, index)

    Call TextAdd(frmServer.txtText(3), GetPlayerName(index) & " On Map " & GetPlayerMap(index) & ": " & Message, True)
    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " : " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_EmoteMessage(ByVal index As Long, ByVal Message As String)
    If frmServer.chkLogEmote.Value = Unchecked Then
        If GetPlayerAccess(index) = 0 Then
            Call PlayerMsg(index, "Emote messages have been disabled by the server!", BRIGHTRED)
            Exit Sub
        End If
    End If

    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & ": " & Message, EmoteColor)

    Call TextAdd(frmServer.txtText(6), GetPlayerName(index) & " " & Message, True)
    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_BroadcastMessage(ByVal index As Long, ByVal Message As String)
    If frmServer.chkLogBC.Value = Unchecked Then
        If GetPlayerAccess(index) = 0 Then
            Call PlayerMsg(index, "Broadcast messages have been disabled by the server!", BRIGHTRED)
            Exit Sub
        End If
    End If

    If Player(index).Mute Then
        Call PlayerMsg(index, "You are muted. You cannot broadcast messages.", BRIGHTRED)
        Exit Sub
    End If

    Call GlobalMsg(GetPlayerName(index) & ": " & Message, BroadcastColor)

    Call TextAdd(frmServer.txtText(0), GetPlayerName(index) & ": " & Message, True)
    Call TextAdd(frmServer.txtText(1), GetPlayerName(index) & ": " & Message, True)
    Call AddLog(GetPlayerName(index) & ": " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_GlobalMessage(ByVal index As Long, ByVal Message As String)
    If frmServer.chkLogGlobal.Value = Unchecked Then
        If GetPlayerAccess(index) = 0 Then
            Call PlayerMsg(index, "Global messages have been disabled by the server!", BRIGHTRED)
            Exit Sub
        End If
    End If

    If Player(index).Mute Then
        Call PlayerMsg(index, "You are muted. You cannot broadcast messages.", BRIGHTRED)
        Exit Sub
    End If

    If GetPlayerAccess(index) > 0 Then
        Call GlobalMsg("(Global) " & GetPlayerName(index) & ": " & Message, GlobalColor)

        Call TextAdd(frmServer.txtText(0), "(Global) " & GetPlayerName(index) & ": " & Message, True)
        Call TextAdd(frmServer.txtText(2), GetPlayerName(index) & ": " & Message, True)
        Call AddLog("(Global) " & GetPlayerName(index) & ": " & Message, ADMIN_LOG)
    End If
End Sub

Public Sub Packet_AdminMessage(ByVal index As Long, ByVal Message As String)
    If frmServer.chkLogAdmin.Value = Unchecked Then
        Call PlayerMsg(index, "Admin messages have been disabled by the server!", BRIGHTRED)
        Exit Sub
    End If

    If GetPlayerAccess(index) > 0 Then
        Call AdminMsg("(Admin " & GetPlayerName(index) & ") " & Message, AdminColor)

        Call TextAdd(frmServer.txtText(5), GetPlayerName(index) & ": " & Message, True)
        Call AddLog("(Admin " & GetPlayerName(index) & ") " & Message, ADMIN_LOG)
    End If
End Sub

Public Sub Packet_PlayerMessage(ByVal index As Long, ByVal Name As String, ByVal Message As String)
    Dim MsgTo As Long
    
    If frmServer.chkLogPM.Value = Unchecked Then
        If GetPlayerAccess(index) = 0 Then
            Call PlayerMsg(index, "Personal messages have been disabled by the server!", BRIGHTRED)
            Exit Sub
        End If
    End If

    If LenB(Name) = 0 Then
        Call PlayerMsg(index, "You must select a player name to private message.", BRIGHTRED)
        Exit Sub
    End If

    If LenB(Message) = 0 Then
        Call PlayerMsg(index, "You must send a message to private message another player.", BRIGHTRED)
        Exit Sub
    End If

    MsgTo = FindPlayer(Name)

    If MsgTo = 0 Then
        Call PlayerMsg(index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    Call PlayerMsg(index, "You tell " & GetPlayerName(MsgTo) & ", '" & Message & "'", TellColor)
    Call PlayerMsg(MsgTo, GetPlayerName(index) & " tells you, '" & Message & "'", TellColor)

    Call TextAdd(frmServer.txtText(4), "To " & GetPlayerName(MsgTo) & " From " & GetPlayerName(index) & ": " & Message, True)
    Call AddLog(GetPlayerName(index) & " tells " & GetPlayerName(MsgTo) & ", " & Message & "'", PLAYER_LOG)
End Sub

Public Sub Packet_PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long, Xpos As Integer, Ypos As Integer)
    If Player(index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(index, "Invalid Direction")
        Exit Sub
    End If

    If Movement <> 1 And Movement <> 2 Then
        Call HackingAttempt(index, "Invalid Movement")
        Exit Sub
    End If

    If Player(index).CastedSpell = YES Then
        If GetTickCount > Player(index).AttackTimer + 1000 Then
            Player(index).CastedSpell = NO
        Else
            Call SendPlayerXY(index)
            Exit Sub
        End If
    End If

    If Player(index).Locked = True Then
        Call SendPlayerXY(index)
        Exit Sub
    End If

    Call PlayerMove(index, Dir, Movement, Xpos, Ypos)
End Sub

Public Sub Packet_PlayerDirection(ByVal index As Long, ByVal Dir As Long)
    If Player(index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(index, "Invalid Direction")
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)

    Call SendDataToMapBut(index, GetPlayerMap(index), "PLAYERDIR" & SEP_CHAR & index & SEP_CHAR & GetPlayerDir(index) & END_CHAR)
End Sub

Public Sub Packet_UseItem(ByVal index As Long, ByVal InvNum As Long)
    Dim CharNum As Long
    Dim SpellID As Long
    Dim MinLvl As Long
    Dim X As Long
    Dim Y As Long

    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        Call HackingAttempt(index, "Invalid InvNum")
        Exit Sub
    End If

    If Player(index).LockedItems Then
        Call PlayerMsg(index, "You currently cannot use any items.", BRIGHTRED)
        Exit Sub
    End If

    CharNum = Player(index).CharNum

    Dim n As Long

    ' Find out what kind of item it is
    Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
        Case ITEM_TYPE_ARMOR
            If InvNum <> GetPlayerArmorSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerArmorSlot(index, InvNum)
                End If
            Else
                Call SetPlayerArmorSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_WEAPON
            If InvNum <> GetPlayerWeaponSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerWeaponSlot(index, InvNum)
                End If
            Else
                Call SetPlayerWeaponSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_TWO_HAND
            If InvNum <> GetPlayerWeaponSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    If GetPlayerShieldSlot(index) <> 0 Then
                        Call SetPlayerShieldSlot(index, 0)
                    End If

                    Call SetPlayerWeaponSlot(index, InvNum)
                End If
            Else
                Call SetPlayerWeaponSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_HELMET
            If InvNum <> GetPlayerHelmetSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerHelmetSlot(index, InvNum)
                End If
            Else
                Call SetPlayerHelmetSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_SHIELD
            If InvNum <> GetPlayerShieldSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    If GetPlayerWeaponSlot(index) <> 0 Then
                        If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Type = ITEM_TYPE_TWO_HAND Then
                            Call SetPlayerWeaponSlot(index, 0)
                        End If
                    End If

                    Call SetPlayerShieldSlot(index, InvNum)
                End If
            Else
                Call SetPlayerShieldSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_LEGS
            If InvNum <> GetPlayerLegsSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerLegsSlot(index, InvNum)
                End If
            Else
                Call SetPlayerLegsSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_RING
            If InvNum <> GetPlayerRingSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerRingSlot(index, InvNum)
                End If
            Else
                Call SetPlayerRingSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_NECKLACE
            If InvNum <> GetPlayerNecklaceSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerNecklaceSlot(index, InvNum)
                End If
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
                    If GetPlayerX(index) < MAX_MAPX Then
                        X = GetPlayerX(index) + 1
                        Y = GetPlayerY(index)
                    Else
                        Exit Sub
                    End If
            End Select
    
            ' Check if a key exists.
            If Map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_KEY Then
                ' Check if the key they are using matches the map key.
                If GetPlayerInvItemNum(index, InvNum) = Map(GetPlayerMap(index)).Tile(X, Y).Data1 Then
                    TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                    TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
    
                    Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)

                    If Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) = vbNullString Then
                        Call MapMsg(GetPlayerMap(index), "A door has been unlocked!", WHITE)
                    Else
                        Call MapMsg(GetPlayerMap(index), Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), WHITE)
                    End If

                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & END_CHAR)
    
                    ' Check if we are supposed to take away the item.
                    If Map(GetPlayerMap(index)).Tile(X, Y).Data2 = 1 Then
                        Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                        Call PlayerMsg(index, "The key disolves.", YELLOW)
                    End If
                End If
            End If
    
            If Map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_DOOR Then
                TempTile(GetPlayerMap(index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
    
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & END_CHAR)
            End If
    
        Case ITEM_TYPE_SPELL
            SpellID = Item(GetPlayerInvItemNum(index, InvNum)).Data1
    
            If SpellID > 0 Then
                If Spell(SpellID).ClassReq - 1 = GetPlayerClass(index) Or Spell(SpellID).ClassReq = 0 Then
                    If Spell(SpellID).LevelReq = 0 And Player(index).Char(Player(index).CharNum).Access < 1 Then
                        Call PlayerMsg(index, "This spell can only be used by admins!", BRIGHTRED)
                        Exit Sub
                    End If

                    MinLvl = GetSpellReqLevel(SpellID)

                    If MinLvl <= GetPlayerLevel(index) Then
                        MinLvl = FindOpenSpellSlot(index)
    
                        If MinLvl > 0 Then
                            If Not HasSpell(index, SpellID) Then
                                Call SetPlayerSpell(index, MinLvl, SpellID)
                                Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                Call PlayerMsg(index, "You have learned a new spell!", WHITE)
                            Else
                                Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                Call PlayerMsg(index, "You have already learned this spell!  The spells crumbles into dust.", BRIGHTRED)
                            End If
                        Else
                            Call PlayerMsg(index, "You have learned all that you can learn!", BRIGHTRED)
                        End If
                    Else
                        Call PlayerMsg(index, "You must be level " & MinLvl & " to learn this spell.", WHITE)
                    End If
                Else
                    Call PlayerMsg(index, "This spell can only be learned by a " & GetClassName(Spell(SpellID).ClassReq - 1) & ".", WHITE)
                End If
            End If
    
        Case ITEM_TYPE_SCRIPTED
            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\main.ess", "ScriptedItem " & index & "," & Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1 & "," & InvNum
            End If
    End Select
    
    Call SendStats(index)
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)

    Call SendIndexWornEquipment(index)
End Sub

' This packet seems to me like it's incomplete. [Mellowz]
Public Sub Packet_PlayerMoveMouse(ByVal index As Long, ByVal Dir As Long)
    If Player(index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(index, "Invalid Direction")
        Exit Sub
    End If

    If Player(index).Locked = True Then
        Call SendPlayerXY(index)
        Exit Sub
    End If

    If Player(index).CastedSpell = YES Then
        If GetTickCount > Player(index).AttackTimer + 1000 Then
            Player(index).CastedSpell = NO
        Else
            Call SendPlayerXY(index)
            Exit Sub
        End If
    End If

    If Val(ReadINI("CONFIG", "mouse", App.Path & "\Data.ini", "0")) = 1 Then
        Call SendDataTo(index, "mouse" & END_CHAR)
    End If
End Sub

Public Sub Packet_Warp(ByVal index As Long, ByVal Dir As Long)
    Select Case Dir
        Case DIR_UP
            If Map(GetPlayerMap(index)).Up > 0 Then
                If GetPlayerY(index) = 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), MAX_MAPY)
                    Exit Sub
                End If
            End If

        Case DIR_DOWN
            If Map(GetPlayerMap(index)).Down > 0 Then
                If GetPlayerY(index) = MAX_MAPY Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Exit Sub
                End If
            End If

        Case DIR_LEFT
            If Map(GetPlayerMap(index)).Left > 0 Then
                If GetPlayerX(index) = 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, MAX_MAPX, GetPlayerY(index))
                    Exit Sub
                End If
            End If

        Case DIR_RIGHT
            If Map(GetPlayerMap(index)).Right > 0 Then
                If GetPlayerX(index) = MAX_MAPX Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Exit Sub
                End If
            End If
    End Select
End Sub

Public Sub Packet_EndShot(ByVal index As Long, ByVal Unknown As Long)
    If Unknown = 0 Then
        Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)
        Player(index).Locked = False
        Player(index).HookShotX = 0
        Player(index).HookShotY = 0
        Exit Sub
    End If

    Call PlayerMsg(index, "You carefully cross the wire.", 1)

    Player(index).Locked = False

    Call SetPlayerX(index, Player(index).HookShotX)
    Call SetPlayerY(index, Player(index).HookShotY)

    Player(index).HookShotX = 0
    Player(index).HookShotY = 0

    Call SendPlayerXY(index)
End Sub

Public Sub Packet_Attack(ByVal index As Long)
    Dim I As Long
    Dim Damage As Long

    If Player(index).LockedAttack Then
        Exit Sub
    End If

    If GetPlayerWeaponSlot(index) > 0 Then
        If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 > 0 Then
            If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Stackable = 0 Then
                Call SendDataToMap(GetPlayerMap(index), "checkarrows" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & END_CHAR)
            Else
                Call GrapleHook(index)
            End If

            Exit Sub
        End If
    End If

    ' Try to attack another player.
    For I = 1 To MAX_PLAYERS
        If I <> index Then
            If CanAttackPlayer(index, I) Then
            
                Player(index).Target = I
                Player(index).TargetType = TARGET_TYPE_PLAYER
            
                If Not CanPlayerBlockHit(I) Then
                    If Not CanPlayerCriticalHit(index) Then
                        Damage = GetPlayerDamage(index) - GetPlayerProtection(I)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & END_CHAR)
                    Else
                        Damage = GetPlayerDamage(index) + Int(Rnd * Int(GetPlayerDamage(index) / 2)) + 1 - GetPlayerProtection(I)

                        Call BattleMsg(index, "You feel a surge of energy upon swinging!", BRIGHTCYAN, 0)
                        Call BattleMsg(I, GetPlayerName(index) & " swings with enormous might!", BRIGHTCYAN, 1)

                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & END_CHAR)
                    End If

                    If Damage > 0 Then
                    If SCRIPTING = 1 Then
                        MyScript.ExecuteStatement "Scripts\main.ess", "OnAttack " & index & "," & Damage
                    Else
                        Call AttackPlayer(index, I, Damage)
                    End If
                    Else
                        If SCRIPTING = 1 Then
                            MyScript.ExecuteStatement "Scripts\main.ess", "OnAttack " & index & "," & Damage
                        End If
                        Call PlayerMsg(index, "Your attack does nothing.", BRIGHTRED)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                    End If
                Else
                    If SCRIPTING = 1 Then
                        MyScript.ExecuteStatement "Scripts\main.ess", "OnAttack " & index & "," & 0
                    End If

                    Call BattleMsg(index, GetPlayerName(I) & " blocked your hit!", BRIGHTCYAN, 0)
                    Call BattleMsg(I, "You blocked " & GetPlayerName(index) & "'s hit!", BRIGHTCYAN, 1)

                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                End If

                Exit Sub
            End If
        End If
    Next I

    ' Try to attack an NPC.
    For I = 1 To MAX_MAP_NPCS
        If CanAttackNpc(index, I) Then
            ' Get the damage we can do
            Player(index).TargetNPC = I
            Player(index).TargetType = TARGET_TYPE_NPC
            If Not CanPlayerCriticalHit(index) Then
                Damage = GetPlayerDamage(index) - Int(NPC(MapNPC(GetPlayerMap(index), I).num).DEF / 2)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & END_CHAR)
            Else
                Damage = GetPlayerDamage(index) + Int(Rnd * Int(GetPlayerDamage(index) / 2)) + 1 - Int(NPC(MapNPC(GetPlayerMap(index), I).num).DEF / 2)
                Call BattleMsg(index, "You feel a surge of energy upon swinging!", BRIGHTCYAN, 0)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & END_CHAR)
            End If
            
            

            If Damage > 0 Then
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\main.ess", "OnAttack " & index & "," & Damage
                Else
                    Call AttackNpc(index, I, Damage)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & I & END_CHAR)
                End If
            Else
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\main.ess", "OnAttack " & index & "," & Damage
                End If
                
                Call BattleMsg(index, "Your attack does nothing.", BRIGHTRED, 0)

                Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & I & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
            End If

            Exit Sub
        End If
    Next I
End Sub

Public Sub Packet_UseStatPoint(ByVal index As Long, ByVal PointType As Long)
    If PointType < 0 Or PointType > 3 Then
        Call HackingAttempt(index, "Invalid Point Type")
        Exit Sub
    End If

    If GetPlayerPOINTS(index) > 0 Then
        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\main.ess", "UsingStatPoints " & index & "," & PointType
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
                    Call BattleMsg(index, "You have gained more magic!", 15, 0)

                Case 3
                    Call SetPlayerSPEED(index, GetPlayerSPEED(index) + 1)
                    Call BattleMsg(index, "You have gained more speed!", 15, 0)
            End Select

            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)
        End If
    Else
        Call BattleMsg(index, "You have no stat points to train with!", BRIGHTRED, 0)
    End If

    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)

    Player(index).Char(Player(index).CharNum).MAXHP = GetPlayerMaxHP(index)
    Player(index).Char(Player(index).CharNum).MAXMP = GetPlayerMaxMP(index)
    Player(index).Char(Player(index).CharNum).MAXSP = GetPlayerMaxSP(index)

    Call SendStats(index)

    Call SendDataTo(index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(index) & END_CHAR)
End Sub

Public Sub Packet_GetStats(ByVal index As Long, ByVal Name As String)
    Dim PlayerID As Long
    Dim BlockChance As Long
    Dim CritChance As Long

    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerMsg(index, "Account: " & Trim$(Player(PlayerID).Login) & "; Name: " & GetPlayerName(PlayerID), BRIGHTGREEN)

        If GetPlayerAccess(index) > ADMIN_MONITER Then
            Call PlayerMsg(index, "Stats for " & GetPlayerName(PlayerID) & ":", BRIGHTGREEN)
            Call PlayerMsg(index, "Level: " & GetPlayerLevel(PlayerID) & "; EXP: " & GetPlayerExp(PlayerID) & "/" & GetPlayerNextLevel(PlayerID), BRIGHTGREEN)
            Call PlayerMsg(index, "HP: " & GetPlayerHP(PlayerID) & "/" & GetPlayerMaxHP(PlayerID) & "; MP: " & GetPlayerMP(PlayerID) & "/" & GetPlayerMaxMP(PlayerID) & "; SP: " & GetPlayerSP(PlayerID) & "/" & GetPlayerMaxSP(PlayerID), BRIGHTGREEN)
            Call PlayerMsg(index, "STR: " & GetPlayerSTR(PlayerID) & "; DEF: " & GetPlayerDEF(PlayerID) & "; MGC: " & GetPlayerMAGI(PlayerID) & "; SPD: " & GetPlayerSPEED(PlayerID), BRIGHTGREEN)
            
            CritChance = Int(GetPlayerSTR(PlayerID) / 2) + Int(GetPlayerLevel(PlayerID) / 2)
            If CritChance < 0 Then
                CritChance = 0
            End If
            If CritChance > 100 Then
                CritChance = 100
            End If

            BlockChance = Int(GetPlayerDEF(PlayerID) / 2) + Int(GetPlayerLevel(PlayerID) / 2)
            If BlockChance < 0 Then
                BlockChance = 0
            End If
            If BlockChance > 100 Then
                BlockChance = 100
            End If

            Call PlayerMsg(index, "Critical Chance: " & CritChance & "%; Block Chance: " & BlockChance & "%", BRIGHTGREEN)
        End If
    Else
        Call PlayerMsg(index, Name & " is currently not online.", WHITE)
    End If
End Sub

Public Sub Packet_SetPlayerSprite(ByVal index As Long, ByVal Name As String, ByVal SpriteID As Long)
    Dim PlayerID As Long

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call SetPlayerSprite(PlayerID, SpriteID)
        Call SendPlayerData(PlayerID)
    Else
        Call PlayerMsg(index, Name & " is currently not online.", WHITE)
    End If
End Sub

Public Sub Packet_RequestNewMap(ByVal index As Long, ByVal Dir As Long)
    Dim X As Integer
    Dim Y As Integer
    
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(index, "Invalid Direction")
        Exit Sub
    End If
    
     Y = GetPlayerY(index)
     X = GetPlayerX(index)
       
     Select Case Dir
         Case DIR_UP
             Y = Y - 1
         Case DIR_DOWN
             Y = Y + 1
         Case DIR_LEFT
             X = X - 1
         Case DIR_RIGHT
             X = X + 1
    End Select

    Call PlayerMove(index, Dir, 1, X, Y)
    Call SendPlayerNewXY(index)
End Sub

Public Sub Packet_WarpMeTo(ByVal index As Long, ByVal Name As String)
    Dim PlayerID As Long

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If
    
    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerWarp(index, GetPlayerMap(PlayerID), GetPlayerX(PlayerID), GetPlayerY(PlayerID))
    Else
        Call PlayerMsg(index, Name & " is currently not online.", WHITE)
    End If
End Sub

Public Sub Packet_WarpToMe(ByVal index As Long, ByVal Name As String)
    Dim PlayerID As Long

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If
    
    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerWarp(PlayerID, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    Else
        Call PlayerMsg(index, Name & " is currently not online.", WHITE)
    End If
End Sub


Public Sub Packet_MapData(ByVal index As Long, ByRef MapData() As String)
    Dim MapIndex As Long
    Dim MapNum As Long
    Dim MapRevision As Long
    Dim X As Long
    Dim Y As Long
    Dim I As Long
    
    ' Check to see if the user is at least a mapper.
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If
            
    MapNum = GetPlayerMap(index)
            
    ' Get revision number before it clears
    MapRevision = Map(MapNum).Revision + 1
            
    MapIndex = 1

    Call ClearMap(MapNum)

    MapNum = Val(MapData(MapIndex))
    Map(MapNum).Name = MapData(MapIndex + 1)
    Map(MapNum).Revision = MapRevision
    Map(MapNum).Moral = Val(MapData(MapIndex + 3))
    Map(MapNum).Up = Val(MapData(MapIndex + 4))
    Map(MapNum).Down = Val(MapData(MapIndex + 5))
    Map(MapNum).Left = Val(MapData(MapIndex + 6))
    Map(MapNum).Right = Val(MapData(MapIndex + 7))
    Map(MapNum).music = MapData(MapIndex + 8)
    Map(MapNum).BootMap = Val(MapData(MapIndex + 9))
    Map(MapNum).BootX = Val(MapData(MapIndex + 10))
    Map(MapNum).BootY = Val(MapData(MapIndex + 11))
    Map(MapNum).Indoors = Val(MapData(MapIndex + 12))
    Map(MapNum).Weather = Val(MapData(MapIndex + 13))

    MapIndex = MapIndex + 14

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map(MapNum).Tile(X, Y).Ground = Val(MapData(MapIndex))
            Map(MapNum).Tile(X, Y).Mask = Val(MapData(MapIndex + 1))
            Map(MapNum).Tile(X, Y).Anim = Val(MapData(MapIndex + 2))
            Map(MapNum).Tile(X, Y).Mask2 = Val(MapData(MapIndex + 3))
            Map(MapNum).Tile(X, Y).M2Anim = Val(MapData(MapIndex + 4))
            Map(MapNum).Tile(X, Y).Fringe = Val(MapData(MapIndex + 5))
            Map(MapNum).Tile(X, Y).FAnim = Val(MapData(MapIndex + 6))
            Map(MapNum).Tile(X, Y).Fringe2 = Val(MapData(MapIndex + 7))
            Map(MapNum).Tile(X, Y).F2Anim = Val(MapData(MapIndex + 8))
            Map(MapNum).Tile(X, Y).Type = Val(MapData(MapIndex + 9))
            Map(MapNum).Tile(X, Y).Data1 = Val(MapData(MapIndex + 10))
            Map(MapNum).Tile(X, Y).Data2 = Val(MapData(MapIndex + 11))
            Map(MapNum).Tile(X, Y).Data3 = Val(MapData(MapIndex + 12))
            Map(MapNum).Tile(X, Y).String1 = MapData(MapIndex + 13)
            Map(MapNum).Tile(X, Y).String2 = MapData(MapIndex + 14)
            Map(MapNum).Tile(X, Y).String3 = MapData(MapIndex + 15)
            Map(MapNum).Tile(X, Y).Light = Val(MapData(MapIndex + 16))
            Map(MapNum).Tile(X, Y).GroundSet = Val(MapData(MapIndex + 17))
            Map(MapNum).Tile(X, Y).MaskSet = Val(MapData(MapIndex + 18))
            Map(MapNum).Tile(X, Y).AnimSet = Val(MapData(MapIndex + 19))
            Map(MapNum).Tile(X, Y).Mask2Set = Val(MapData(MapIndex + 20))
            Map(MapNum).Tile(X, Y).M2AnimSet = Val(MapData(MapIndex + 21))
            Map(MapNum).Tile(X, Y).FringeSet = Val(MapData(MapIndex + 22))
            Map(MapNum).Tile(X, Y).FAnimSet = Val(MapData(MapIndex + 23))
            Map(MapNum).Tile(X, Y).Fringe2Set = Val(MapData(MapIndex + 24))
            Map(MapNum).Tile(X, Y).F2AnimSet = Val(MapData(MapIndex + 25))

            MapIndex = MapIndex + 26
        Next X
    Next Y

    For X = 1 To MAX_MAP_NPCS
        Map(MapNum).NPC(X) = Val(MapData(MapIndex))
        Map(MapNum).SpawnX(X) = Val(MapData(MapIndex + 1))
        Map(MapNum).SpawnY(X) = Val(MapData(MapIndex + 2))
        MapIndex = MapIndex + 3
        Call ClearMapNpc(X, MapNum)
    Next X

    ' Clear out it all
    For I = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), I).X, MapItem(GetPlayerMap(index), I).Y)
        Call ClearMapItem(I, GetPlayerMap(index))
    Next I

    ' Save the map
    Call SaveMap(MapNum)
            
    ' Mapper is on the map
    PlayersOnMap(MapNum) = YES

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))

    ' Respawn NPCS
    For I = 1 To MAX_MAP_NPCS
        Call SpawnNpc(I, GetPlayerMap(index))
    Next I

    ' Refresh map for everyone online
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                Call SendDataTo(I, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(I) & SEP_CHAR & Map(GetPlayerMap(I)).Revision & END_CHAR)
            End If
        End If
    Next I
End Sub

Public Sub Packet_NeedMap(ByVal index As Long, ByVal NeedMap As String)
    Dim I As Long

    NeedMap = UCase$(NeedMap)

    If NeedMap = "YES" Then
        Call SendMap(index, GetPlayerMap(index))
    End If

    Call SendMapItemsTo(index, GetPlayerMap(index))
    Call SendMapNpcsTo(index, GetPlayerMap(index))
    Call SendJoinMap(index)
    Call SendDataTo(index, "MAPDONE" & END_CHAR)

    Player(index).GettingMap = NO

    Call SendPlayerData(index)

    For I = 1 To MAX_PLAYERS
       If IsPlaying(I) Then
           Call SendHP(I)
           Call SendIndexWornEquipment(I)
           Call SendWornEquipment(I)
       End If
   Next
End Sub

Public Sub Packet_MapGetItem(ByVal index As Long)
    Call PlayerMapGetItem(index)
End Sub

Public Sub Packet_MapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    If InvNum < 1 Or InvNum > MAX_INV Then
        Call HackingAttempt(index, "Invalid InvNum")
        Exit Sub
    End If

    ' Prevent hacking
    If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
        If Amount <= 0 Then
            Call PlayerMsg(index, "You must at least drop 1 of that item!", BRIGHTRED)
            Exit Sub
        End If

        If Amount > GetPlayerInvItemValue(index, InvNum) Then
            Call PlayerMsg(index, "You don't have that much to drop!", BRIGHTRED)
            Exit Sub
        End If
    End If

    ' Prevent hacking
    If Item(GetPlayerInvItemNum(index, InvNum)).Type <> ITEM_TYPE_CURRENCY Then
        If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
            If Amount > GetPlayerInvItemValue(index, InvNum) Then
                Call HackingAttempt(index, "Item amount modification")
                Exit Sub
            End If
        End If
    End If

    Call PlayerMapDropItem(index, InvNum, Amount)

    Call SendStats(index)
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
End Sub

Public Sub Packet_MapRespawn(ByVal index As Long)
    Dim I As Long

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    ' Clear out all of the floor items.
    For I = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), I).X, MapItem(GetPlayerMap(index), I).Y)
        Call ClearMapItem(I, GetPlayerMap(index))
    Next I

    ' Respawn all of the floor items.
    Call SpawnMapItems(GetPlayerMap(index))

    ' Respawn NPCS
    For I = 1 To MAX_MAP_NPCS
        Call SpawnNpc(I, GetPlayerMap(index))
    Next I

    Call PlayerMsg(index, "Map respawned.", BLUE)
End Sub

Public Sub Packet_KickPlayer(ByVal index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    If GetPlayerAccess(index) < 1 Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If PlayerIndex <> index Then
            If GetPlayerAccess(PlayerIndex) <= GetPlayerAccess(index) Then
                Call GlobalMsg(GetPlayerName(PlayerIndex) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(index) & "!", WHITE)
                Call AddLog(GetPlayerName(index) & " has kicked " & GetPlayerName(PlayerIndex) & ".", ADMIN_LOG)
                Call AlertMsg(PlayerIndex, "You have been kicked by " & GetPlayerName(index) & "!")
            Else
                Call PlayerMsg(index, "That admin has a higher access then you!", WHITE)
            End If
        Else
            Call PlayerMsg(index, "You cannot kick yourself!", WHITE)
        End If
    Else
        Call PlayerMsg(index, "Player is not online.", WHITE)
    End If
End Sub

Public Sub Packet_BanList(ByVal index As Long)
    Dim FileID As Integer
    Dim PlayerName As String

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If
            
    If Not FileExists("BanList.txt") Then
        Call PlayerMsg(index, "The ban list cannot be found!", BRIGHTRED)
        Exit Sub
    End If

    FileID = FreeFile

    Open App.Path & "\BanList.txt" For Input As #FileID
    Do While Not EOF(FileID)
        Line Input #FileID, PlayerName
        Call PlayerMsg(index, PlayerName, WHITE)
    Loop
    Close #FileID
End Sub

Public Sub Packet_BanListDestroy(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If FileExists("BanList.txt") Then
        Call Kill(App.Path & "\BanList.txt")
    End If

    Call PlayerMsg(index, "Ban list destroyed.", WHITE)
End Sub

Public Sub Packet_BanPlayer(ByVal index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If PlayerIndex <> index Then
            If GetPlayerAccess(PlayerIndex) <= GetPlayerAccess(index) Then
                Call BanIndex(PlayerIndex, index)
            Else
                Call PlayerMsg(index, "That admin has a higher access then you!", WHITE)
            End If
        Else
            Call PlayerMsg(index, "You cannot ban yourself!", WHITE)
        End If
    Else
        Call PlayerMsg(index, "Player is not online.", WHITE)
    End If
End Sub

Public Sub Packet_RequestEditMap(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "EDITMAP" & END_CHAR)
End Sub

Public Sub Packet_RequestEditItem(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "ITEMEDITOR" & END_CHAR)
End Sub

Public Sub Packet_EditItem(ByVal index As Long, ByVal ItemNum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Call HackingAttempt(index, "Invalid Item Index")
        Exit Sub
    End If

    Call SendEditItemTo(index, ItemNum)

    Call AddLog(GetPlayerName(index) & " editing item #" & ItemNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveItem(ByVal index As Long, ByRef ItemData() As String)
    Dim ItemNum As Long

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    ItemNum = Val(ItemData(1))

    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Call HackingAttempt(index, "Invalid Item Index")
        Exit Sub
    End If

    Item(ItemNum).Name = ItemData(2)
    Item(ItemNum).Pic = Val(ItemData(3))
    Item(ItemNum).Type = Val(ItemData(4))
    Item(ItemNum).Data1 = Val(ItemData(5))
    Item(ItemNum).Data2 = Val(ItemData(6))
    Item(ItemNum).Data3 = Val(ItemData(7))
    Item(ItemNum).StrReq = Val(ItemData(8))
    Item(ItemNum).DefReq = Val(ItemData(9))
    Item(ItemNum).SpeedReq = Val(ItemData(10))
    Item(ItemNum).MagicReq = Val(ItemData(11))
    Item(ItemNum).ClassReq = Val(ItemData(12))
    Item(ItemNum).AccessReq = Val(ItemData(13))

    Item(ItemNum).addHP = Val(ItemData(14))
    Item(ItemNum).addMP = Val(ItemData(15))
    Item(ItemNum).addSP = Val(ItemData(16))
    Item(ItemNum).AddStr = Val(ItemData(17))
    Item(ItemNum).AddDef = Val(ItemData(18))
    Item(ItemNum).AddMagi = Val(ItemData(19))
    Item(ItemNum).AddSpeed = Val(ItemData(20))
    Item(ItemNum).AddEXP = Val(ItemData(21))
    Item(ItemNum).Desc = ItemData(22)
    Item(ItemNum).AttackSpeed = Val(ItemData(23))
    Item(ItemNum).Price = Val(ItemData(24))
    Item(ItemNum).Stackable = Val(ItemData(25))
    Item(ItemNum).Bound = Val(ItemData(26))

    Call SendUpdateItemToAll(ItemNum)
    Call SaveItem(ItemNum)

    Call AddLog(GetPlayerName(index) & " saved item #" & ItemNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_EnableDayNight(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If Not TimeDisable Then
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
End Sub

Public Sub Packet_DayNight(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If Hours > 12 Then
        Hours = Hours - 12
    Else
        Hours = Hours + 12
    End If
End Sub

Public Sub Packet_RequestEditNPC(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "NPCEDITOR" & END_CHAR)
End Sub

Public Sub Packet_EditNPC(ByVal index As Long, ByVal NPCnum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If NPCnum < 0 Or NPCnum > MAX_NPCS Then
        Call HackingAttempt(index, "Invalid NPC Index")
        Exit Sub
    End If

    Call SendEditNpcTo(index, NPCnum)

    Call AddLog(GetPlayerName(index) & " editing npc #" & NPCnum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveNPC(ByVal index As Long, ByRef NPCData() As String)
    Dim NPCnum As Long
    Dim NPCIndex As Long
    Dim I As Long

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    NPCnum = Val(NPCData(1))

    If NPCnum < 0 Or NPCnum > MAX_NPCS Then
        Call HackingAttempt(index, "Invalid NPC Index")
        Exit Sub
    End If

    NPC(NPCnum).Name = NPCData(2)
    NPC(NPCnum).AttackSay = NPCData(3)
    NPC(NPCnum).Sprite = Val(NPCData(4))
    NPC(NPCnum).SpawnSecs = Val(NPCData(5))
    NPC(NPCnum).Behavior = Val(NPCData(6))
    NPC(NPCnum).Range = Val(NPCData(7))
    NPC(NPCnum).STR = Val(NPCData(8))
    NPC(NPCnum).DEF = Val(NPCData(9))
    NPC(NPCnum).Speed = Val(NPCData(10))
    NPC(NPCnum).Magi = Val(NPCData(11))
    NPC(NPCnum).Big = Val(NPCData(12))
    NPC(NPCnum).MAXHP = Val(NPCData(13))
    NPC(NPCnum).Exp = Val(NPCData(14))
    NPC(NPCnum).SpawnTime = Val(NPCData(15))
    NPC(NPCnum).Element = Val(NPCData(16))
    NPC(NPCnum).SPRITESIZE = Val(NPCData(17))

    NPCIndex = 18

    For I = 1 To MAX_NPC_DROPS
        NPC(NPCnum).ItemNPC(I).Chance = Val(NPCData(NPCIndex))
        NPC(NPCnum).ItemNPC(I).ItemNum = Val(NPCData(NPCIndex + 1))
        NPC(NPCnum).ItemNPC(I).ItemValue = Val(NPCData(NPCIndex + 2))
        NPCIndex = NPCIndex + 3
    Next I

    Call SendUpdateNpcToAll(NPCnum)
    Call SaveNpc(NPCnum)

    Call AddLog(GetPlayerName(index) & " saved npc #" & NPCnum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_RequestEditShop(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "SHOPEDITOR" & END_CHAR)
End Sub

Public Sub Packet_EditShop(ByVal index As Long, ByVal ShopNum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
        Call HackingAttempt(index, "Invalid Shop Index")
        Exit Sub
    End If

    Call SendEditShopTo(index, ShopNum)

    Call AddLog(GetPlayerName(index) & " editing shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveShop(ByVal index As Long, ByRef ShopData() As String)
    Dim ShopNum As Long
    Dim ShopIndex As Long
    Dim I As Long

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    ShopNum = Val(ShopData(1))

    If ShopNum < 1 Or ShopNum > MAX_SHOPS Then
        Call HackingAttempt(index, "Invalid Shop Index")
        Exit Sub
    End If

    Shop(ShopNum).Name = ShopData(2)
    Shop(ShopNum).FixesItems = Val(ShopData(3))
    Shop(ShopNum).BuysItems = Val(ShopData(4))
    Shop(ShopNum).ShowInfo = Val(ShopData(5))
    Shop(ShopNum).CurrencyItem = Val(ShopData(6))

    ShopIndex = 7

    For I = 1 To MAX_SHOP_ITEMS
        Shop(ShopNum).ShopItem(I).ItemNum = Val(ShopData(ShopIndex))
        Shop(ShopNum).ShopItem(I).Amount = Val(ShopData(ShopIndex + 1))
        Shop(ShopNum).ShopItem(I).Price = Val(ShopData(ShopIndex + 2))
        ShopIndex = ShopIndex + 3
    Next I

    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)

    Call AddLog(GetPlayerName(index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_RequestEditSpell(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "SPELLEDITOR" & END_CHAR)
End Sub

Public Sub Packet_EditSpell(ByVal index As Long, ByVal SpellNum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Call HackingAttempt(index, "Invalid Spell Index")
        Exit Sub
    End If

    Call SendEditSpellTo(index, SpellNum)

    Call AddLog(GetPlayerName(index) & " editing spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveSpell(ByVal index As Long, ByRef SpellData() As String)
    Dim SpellNum As Long
    
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    SpellNum = Val(SpellData(1))

    If SpellNum < 1 Or SpellNum > MAX_SPELLS Then
        Call HackingAttempt(index, "Invalid Spell Index")
        Exit Sub
    End If

    Spell(SpellNum).Name = SpellData(2)
    Spell(SpellNum).ClassReq = Val(SpellData(3))
    Spell(SpellNum).LevelReq = Val(SpellData(4))
    Spell(SpellNum).Type = Val(SpellData(5))
    Spell(SpellNum).Data1 = Val(SpellData(6))
    Spell(SpellNum).Data2 = Val(SpellData(7))
    Spell(SpellNum).Data3 = Val(SpellData(8))
    Spell(SpellNum).MPCost = Val(SpellData(9))
    Spell(SpellNum).Sound = Val(SpellData(10))
    Spell(SpellNum).Range = Val(SpellData(11))
    Spell(SpellNum).SpellAnim = Val(SpellData(12))
    Spell(SpellNum).SpellTime = Val(SpellData(13))
    Spell(SpellNum).SpellDone = Val(SpellData(14))
    Spell(SpellNum).AE = Val(SpellData(15))
    Spell(SpellNum).Big = Val(SpellData(16))
    Spell(SpellNum).Element = Val(SpellData(17))

    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)

    Call AddLog(GetPlayerName(index) & " saving spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_ForgetSpell(ByVal index As Long, ByVal SpellNum As Long)
    If SpellNum < 1 Or SpellNum > MAX_PLAYER_SPELLS Then
        Call HackingAttempt(index, "Invalid Spell Slot")
        Exit Sub
    End If

    With Player(index).Char(Player(index).CharNum)
        If .Spell(SpellNum) = 0 Then
            Call PlayerMsg(index, "No spell here.", RED)
        Else
            Call PlayerMsg(index, "You have forgotten the spell " & Trim$(Spell(.Spell(SpellNum)).Name) & ".", GREEN)

            .Spell(SpellNum) = 0

            Call SendSpells(index)
        End If
    End With
End Sub

Public Sub Packet_SetAccess(ByVal index As Long, ByVal Name As String, ByVal AccessLvl As Long)
    Dim PlayerIndex As Long
    
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Call HackingAttempt(index, "Invalid Access")
        Exit Sub
    End If
    
    If AccessLvl < 0 Or AccessLvl > 5 Then
        Call PlayerMsg(index, "You have entered an invalid access level.", BRIGHTRED)
        Exit Sub
    End If

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If GetPlayerName(index) <> GetPlayerName(PlayerIndex) Then
            If GetPlayerAccess(index) > GetPlayerAccess(PlayerIndex) Then
                Call SetPlayerAccess(PlayerIndex, AccessLvl)
                Call SendPlayerData(PlayerIndex)
    
                If GetPlayerAccess(PlayerIndex) = 0 Then
                    Call GlobalMsg(GetPlayerName(PlayerIndex) & " has been blessed with administrative access.", BRIGHTBLUE)
                End If
    
                Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(PlayerIndex) & "'s access.", ADMIN_LOG)
            Else
                Call PlayerMsg(index, "Your access level is lower than " & GetPlayerName(PlayerIndex) & ".", RED)
            End If
        Else
            Call PlayerMsg(index, "You cant change your access.", RED)
        End If
    Else
        Call PlayerMsg(index, "Player is not online.", WHITE)
    End If
End Sub

Public Sub Packet_WhoIsOnline(ByVal index As Long)
    Call SendWhosOnline(index)
End Sub

Public Sub Packet_OnlineList(ByVal index As Long)
    Call SendOnlineList
End Sub

Public Sub Packet_SetMOTD(ByVal index As Long, ByVal MOTD As String)
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call PutVar(App.Path & "\MOTD.ini", "MOTD", "Msg", MOTD)
            
    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\main.ess", "ChangeMOTD"
    End If
            
    Call GlobalMsg("MOTD changed to: " & MOTD, BRIGHTCYAN)

    Call AddLog(GetPlayerName(index) & " changed MOTD to: " & MOTD, ADMIN_LOG)
End Sub

Public Sub Packet_BuyItem(ByVal index As Long, ByVal ShopIndex As Long, ByVal ItemIndex As Long)
    Dim InvItem As Long

    If ShopIndex < 1 Or ShopIndex > MAX_SHOPS Then
        Call HackingAttempt(index, "Invalid Shop Index")
        Exit Sub
    End If
    
    If ItemIndex < 1 Or ItemIndex > MAX_SHOP_ITEMS Then
        Call HackingAttempt(index, "Invalid Shop Item")
        Exit Sub
    End If

    ' Check to see if player's inventory is full.
    InvItem = FindOpenInvSlot(index, Shop(ShopIndex).ShopItem(ItemIndex).ItemNum)
    If InvItem = 0 Then
        Call PlayerMsg(index, "Your inventory has reached its maximum capacity!", BRIGHTRED)
        Exit Sub
    End If

    ' Check to see if they have enough currency.
    If HasItem(index, Shop(ShopIndex).CurrencyItem) >= Shop(ShopIndex).ShopItem(ItemIndex).Price Then
        Call TakeItem(index, Shop(ShopIndex).CurrencyItem, Shop(ShopIndex).ShopItem(ItemIndex).Price)
        Call GiveItem(index, Shop(ShopIndex).ShopItem(ItemIndex).ItemNum, Shop(ShopIndex).ShopItem(ItemIndex).Amount)

        Call PlayerMsg(index, "You bought the item.", YELLOW)
    Else
        Call PlayerMsg(index, "You cannot afford that!", RED)
    End If
End Sub

Public Sub Packet_SellItem(ByVal index As Long, ByVal ShopNum As Long, ByVal ItemNum As Long, ByVal ItemSlot As Long, ByVal ItemAmt As Long)
    If ItemIsEquipped(index, ItemNum) Then
        Call PlayerMsg(index, "You cannot sell worn items.", RED)
        Exit Sub
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        Call PlayerMsg(index, "You cannot sell currency.", RED)
        Exit Sub
    End If

    If Item(ItemNum).Stackable = YES Then
        If ItemAmt > GetPlayerInvItemValue(index, ItemSlot) Then
            Call PlayerMsg(index, "You don't have enough of that item to sell that many!", RED)
            Exit Sub
        End If
    End If

    If Item(ItemNum).Price > 0 Then
        Call TakeItem(index, ItemNum, ItemAmt)
        Call GiveItem(index, Shop(ShopNum).CurrencyItem, Item(ItemNum).Price * ItemAmt)
        Call PlayerMsg(index, "The shopkeeper hands you " & Item(ItemNum).Price * ItemAmt & " " & Trim$(Item(Shop(ShopNum).CurrencyItem).Name) & ".", YELLOW)
    Else
        Call PlayerMsg(index, "This item cannot be sold.", RED)
    End If
End Sub

Public Sub Packet_FixItem(ByVal index As Long, ByVal ShopNum As Long, ByVal InvNum As Long)
    Dim ItemNum As Long
    Dim DurNeeded As Long
    Dim GoldNeeded As Long
    Dim I As Long

    If Item(GetPlayerInvItemNum(index, InvNum)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(index, InvNum)).Type > ITEM_TYPE_NECKLACE Then
        Call PlayerMsg(index, "That item doesn't need to be fixed.", BRIGHTRED)
        Exit Sub
    End If

    If FindOpenInvSlot(index, GetPlayerInvItemNum(index, InvNum)) = 0 Then
        Call PlayerMsg(index, "You have no inventory space left!", BRIGHTRED)
        Exit Sub
    End If

    ItemNum = GetPlayerInvItemNum(index, InvNum)

    I = Int(Item(GetPlayerInvItemNum(index, InvNum)).Data2 / 5)
    If I <= 0 Then
        I = 1
    End If

    DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(index, InvNum)

    GoldNeeded = Int(DurNeeded * I / 2)
    If GoldNeeded <= 0 Then
        GoldNeeded = 1
    End If

    If DurNeeded = 0 Then
        Call PlayerMsg(index, "This item is in perfect condition!", WHITE)
        Exit Sub
    End If

    If HasItem(index, Shop(ShopNum).CurrencyItem) >= I Then
        If HasItem(index, Shop(ShopNum).CurrencyItem) >= GoldNeeded Then
            Call TakeItem(index, Shop(ShopNum).CurrencyItem, GoldNeeded)
            Call SetPlayerInvItemDur(index, InvNum, Item(ItemNum).Data1)

            Call PlayerMsg(index, "Item has been totally restored for " & GoldNeeded & " gold!", BRIGHTBLUE)
        Else
            DurNeeded = (HasItem(index, Shop(ShopNum).CurrencyItem) / I)
            GoldNeeded = Int(DurNeeded * I / 2)

            If GoldNeeded <= 0 Then
                GoldNeeded = 1
            End If

            Call TakeItem(index, Shop(ShopNum).CurrencyItem, GoldNeeded)
            Call SetPlayerInvItemDur(index, InvNum, GetPlayerInvItemDur(index, InvNum) + DurNeeded)

            Call PlayerMsg(index, "Item has been partially fixed for " & GoldNeeded & " gold!", BRIGHTBLUE)
        End If
    Else
        Call PlayerMsg(index, "You don't have enough gold to fix this item!", BRIGHTRED)
    End If
End Sub

Public Sub Packet_Search(ByVal index As Long, ByVal X As Long, ByVal Y As Long)
    Dim I As Long

    If X < 0 Or X > MAX_MAPX Then
        Exit Sub
    End If

    If Y < 0 Or Y > MAX_MAPY Then
        Exit Sub
    End If

    ' Check for a player
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(index) = GetPlayerMap(I) Then
                If GetPlayerX(I) = X Then
                    If GetPlayerY(I) = Y Then
                        If GetPlayerLevel(I) >= GetPlayerLevel(index) + 5 Then
                            Call PlayerMsg(index, "You wouldn't stand a chance.", BRIGHTRED)
                        Else
                            If GetPlayerLevel(I) > GetPlayerLevel(index) Then
                                Call PlayerMsg(index, "This one seems to have an advantage over you.", YELLOW)
                            Else
                                If GetPlayerLevel(I) = GetPlayerLevel(index) Then
                                    Call PlayerMsg(index, "This would be an even fight.", WHITE)
                                Else
                                    If GetPlayerLevel(index) >= GetPlayerLevel(I) + 5 Then
                                        Call PlayerMsg(index, "You could slaughter that player.", BRIGHTBLUE)
                                    Else
                                        If GetPlayerLevel(index) > GetPlayerLevel(I) Then
                                            Call PlayerMsg(index, "You would have an advantage over that player.", YELLOW)
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        ' Change the target.
                        Player(index).Target = I
                        Player(index).TargetType = TARGET_TYPE_PLAYER
                        If SCRIPTING = 1 Then
                            'MyScript.ExecuteStatement "Scripts\main.ess", "OnClickPlayer " & Index
                        End If

                        Call PlayerMsg(index, "Your target is now " & GetPlayerName(I) & ".", YELLOW)

                        Exit Sub
                    End If
                End If
            End If

        End If
    Next I

    ' Check for an NPC.
    For I = 1 To MAX_MAP_NPCS
        If MapNPC(GetPlayerMap(index), I).num > 0 Then
            If MapNPC(GetPlayerMap(index), I).X = X Then
                If MapNPC(GetPlayerMap(index), I).Y = Y Then
                    Player(index).TargetNPC = I
                    Player(index).TargetType = TARGET_TYPE_NPC

                    Call PlayerMsg(index, "Your target is now a " & Trim$(NPC(MapNPC(GetPlayerMap(index), I).num).Name) & ".", YELLOW)

                    Exit Sub
                End If
            End If
        End If
    Next I

    ' Check for an item on the ground.
    For I = 1 To MAX_MAP_ITEMS
        If MapItem(GetPlayerMap(index), I).num > 0 Then
            If MapItem(GetPlayerMap(index), I).X = X Then
                If MapItem(GetPlayerMap(index), I).Y = Y Then
                    Call PlayerMsg(index, "You see a " & Trim$(Item(MapItem(GetPlayerMap(index), I).num).Name) & ".", YELLOW)
                    Exit Sub
                End If
            End If
        End If
    Next I

    ' Check for an OnClick tile.
    If Map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_ONCLICK Then
        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\main.ess", "OnClick " & index & "," & Map(GetPlayerMap(index)).Tile(X, Y).Data1
        End If
    End If
End Sub

Public Sub Packet_Search2(ByVal index As Long, ByVal X As Long, ByVal Y As Long)
    Dim I As Long

    If X < 0 Or X > MAX_MAPX Then
        Exit Sub
    End If

    If Y < 0 Or Y > MAX_MAPY Then
        Exit Sub
    End If

    ' Check for a player
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(index) = GetPlayerMap(I) Then
                If GetPlayerX(I) = X Then
                    If GetPlayerY(I) = Y Then
                        
                        ' Change the target.
                        Player(index).Target = I
                        Player(index).TargetType = TARGET_TYPE_PLAYER
                        If SCRIPTING = 1 Then
                            MyScript.ExecuteStatement "Scripts\main.ess", "ShowProfile " & index & "," & I
                        End If

                        Exit Sub
                    End If
                End If
            End If

        End If
    Next I

End Sub

Public Sub Packet_PlayerChat(ByVal index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, Name & " is currently not online.", WHITE)
        Exit Sub
    End If

    If PlayerIndex = index Then
        Call PlayerMsg(index, "You cannot chat with yourself.", PINK)
        Exit Sub
    End If

    If Player(index).InChat = 1 Then
        Call PlayerMsg(index, "You're already in a chat with another player!", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).InChat = 1 Then
        Call PlayerMsg(index, Name & " is already in a chat with another player!", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "Chat request has been sent to " & GetPlayerName(PlayerIndex) & ".", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " wants you to chat with them. Type /chat to accept, or /chatdecline to decline.", PINK)

    Player(index).ChatPlayer = PlayerIndex
    Player(PlayerIndex).ChatPlayer = index
End Sub

Public Sub Packet_AcceptChat(ByVal index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "You have not requested to chat with anyone.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> index Then
        Call PlayerMsg(index, "Chat failed.", PINK)
        Exit Sub
    End If

    Call SendDataTo(index, "PPCHATTING" & SEP_CHAR & PlayerIndex & END_CHAR)
    Call SendDataTo(PlayerIndex, "PPCHATTING" & SEP_CHAR & index & END_CHAR)
End Sub

Public Sub Packet_DenyChat(ByVal index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "You have not requested to chat with anyone.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> index Then
        Call PlayerMsg(index, "Chat failed.", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "Declined chat request.", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " declined your request.", PINK)

    Player(index).ChatPlayer = 0
    Player(index).InChat = 0

    Player(PlayerIndex).ChatPlayer = 0
    Player(PlayerIndex).InChat = 0
End Sub

Public Sub Packet_QuitChat(ByVal index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "You have not requested to chat with anyone.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> index Then
        Call PlayerMsg(index, "Chat failed.", PINK)
        Exit Sub
    End If

    Call SendDataTo(index, "qchat" & END_CHAR)
    Call SendDataTo(PlayerIndex, "qchat" & END_CHAR)

    Player(index).ChatPlayer = 0
    Player(index).InChat = 0

    Player(PlayerIndex).ChatPlayer = 0
    Player(PlayerIndex).InChat = 0
End Sub

Public Sub Packet_SendChat(ByVal index As Long, ByVal Message As String)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "You have not requested to chat with anyone.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> index Then
        Call PlayerMsg(index, "Chat failed.", PINK)
        Exit Sub
    End If

    Call SendDataTo(PlayerIndex, "sendchat" & SEP_CHAR & Message & SEP_CHAR & index & END_CHAR)
End Sub

Public Sub Packet_PrepareTrade(ByVal index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, Name & " is currently not online.", WHITE)
        Exit Sub
    End If

    If PlayerIndex = index Then
        Call PlayerMsg(index, "You cannot trade with yourself!", PINK)
        Exit Sub
    End If

    If GetPlayerMap(index) <> GetPlayerMap(PlayerIndex) Then
        Call PlayerMsg(index, "You must be on the same map to trade with " & GetPlayerName(PlayerIndex) & "!", PINK)
        Exit Sub
    End If

    If Player(index).InTrade Then
        Call PlayerMsg(index, "You're already in a trade with someone else!", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).InTrade Then
        Call PlayerMsg(index, Name & " is already in a trade!", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "Trade request has been sent to " & GetPlayerName(PlayerIndex) & ".", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " wants you to trade with them. Type /accept to accept, or /decline to decline.", PINK)

    Player(index).TradePlayer = PlayerIndex
    Player(PlayerIndex).TradePlayer = index
End Sub

Public Sub Packet_AcceptTrade(ByVal index As Long)
    Dim PlayerIndex As Long
    Dim I As Long

    PlayerIndex = Player(index).TradePlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "You have not requested to trade with anyone.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).TradePlayer <> index Then
        Call PlayerMsg(index, "Trade failed.", PINK)
        Exit Sub
    End If

    If GetPlayerMap(index) <> GetPlayerMap(PlayerIndex) Then
        Call PlayerMsg(index, "You must be on the same map to trade with " & GetPlayerName(PlayerIndex) & "!", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "You are trading with " & GetPlayerName(PlayerIndex) & "!", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " accepted your trade request!", PINK)

    Call SendDataTo(index, "PPTRADING" & END_CHAR)
    Call SendDataTo(PlayerIndex, "PPTRADING" & END_CHAR)

    For I = 1 To MAX_PLAYER_TRADES
        Player(index).Trading(I).InvNum = 0
        Player(index).Trading(I).InvName = vbNullString
        Player(index).Trading(I).InvVal = 0

        Player(PlayerIndex).Trading(I).InvNum = 0
        Player(PlayerIndex).Trading(I).InvName = vbNullString
        Player(PlayerIndex).Trading(I).InvVal = 0
    Next I

    Player(index).InTrade = True
    Player(index).TradeItemMax = 0
    Player(index).TradeItemMax2 = 0

    Player(PlayerIndex).InTrade = True
    Player(PlayerIndex).TradeItemMax = 0
    Player(PlayerIndex).TradeItemMax2 = 0
End Sub

Public Sub Packet_QuitTrade(ByVal index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).TradePlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "You have not requested to trade with anyone.", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "Stopped trading.", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " stopped trading with you!", PINK)

    Player(index).TradeOk = 0
    Player(index).TradePlayer = 0
    Player(index).InTrade = False

    Player(PlayerIndex).TradeOk = 0
    Player(PlayerIndex).TradePlayer = 0
    Player(PlayerIndex).InTrade = False

    Call SendDataTo(index, "qtrade" & END_CHAR)
    Call SendDataTo(PlayerIndex, "qtrade" & END_CHAR)
End Sub

Public Sub Packet_DenyTrade(ByVal index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).TradePlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "You have not requested to trade with anyone.", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "Declined trade request.", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " declined your request.", PINK)

    Player(index).TradePlayer = 0
    Player(index).InTrade = False

    Player(PlayerIndex).TradePlayer = 0
    Player(PlayerIndex).InTrade = False
End Sub

Public Sub Packet_UpdateTradeInventory(ByVal index As Long, ByVal TradeIndex As Long, ByVal ItemNum As Long, ByVal ItemName As String, ByVal ItemValue As Long)
    Player(index).Trading(TradeIndex).InvNum = ItemNum
    Player(index).Trading(TradeIndex).InvName = Trim$(ItemName)
    Player(index).Trading(TradeIndex).InvVal = ItemValue

    If Player(index).Trading(TradeIndex).InvNum = 0 Then
        Player(index).TradeItemMax = Player(index).TradeItemMax - 1
        Player(index).TradeOk = 0
        Player(TradeIndex).TradeOk = 0

        Call SendDataTo(index, "trading" & SEP_CHAR & 0 & END_CHAR)
        Call SendDataTo(TradeIndex, "trading" & SEP_CHAR & 0 & END_CHAR)
    Else
        Player(index).TradeItemMax = Player(index).TradeItemMax + 1
    End If

    If Player(index).Trading(TradeIndex).InvNum > 0 Then
        Call SendDataTo(Player(index).TradePlayer, "updatetradeitem" & SEP_CHAR & TradeIndex & SEP_CHAR & Player(index).Trading(TradeIndex).InvNum & SEP_CHAR & Player(index).Trading(TradeIndex).InvName & SEP_CHAR & Player(index).Trading(TradeIndex).InvVal & END_CHAR)
        Else
        Call SendDataTo(Player(index).TradePlayer, "updatetradeitem" & SEP_CHAR & TradeIndex & SEP_CHAR & 0 & SEP_CHAR & vbNullString & SEP_CHAR & 0 & END_CHAR)
    End If
End Sub

Public Sub Packet_SwapItems(ByVal index As Long)
    Dim TradeIndex As Long
    Dim I As Long
    Dim X As Long
    TradeIndex = Player(index).TradePlayer

    If Player(index).TradeOk = 0 Then
        Player(index).TradeOk = 1
        Call SendDataTo(TradeIndex, "trading" & SEP_CHAR & 1 & END_CHAR)
    ElseIf Player(index).TradeOk = 1 Then
        Player(index).TradeOk = 0
        Call SendDataTo(TradeIndex, "trading" & SEP_CHAR & 0 & END_CHAR)
    End If

    If Player(index).TradeOk = 1 Then
        If Player(TradeIndex).TradeOk = 1 Then
            Player(index).TradeItemMax2 = 0
            Player(TradeIndex).TradeItemMax2 = 0
    
            For I = 1 To MAX_INV
                If Player(index).TradeItemMax = Player(index).TradeItemMax2 Then
                    Exit For
                End If

                If GetPlayerInvItemNum(TradeIndex, I) < 1 Then
                    Player(index).TradeItemMax2 = Player(index).TradeItemMax2 + 1
                End If
            Next I
    
            For I = 1 To MAX_INV
                If Player(TradeIndex).TradeItemMax = Player(TradeIndex).TradeItemMax2 Then
                    Exit For
                End If

                If GetPlayerInvItemNum(index, I) < 1 Then
                    Player(TradeIndex).TradeItemMax2 = Player(TradeIndex).TradeItemMax2 + 1
                End If
            Next I
    
            If Player(index).TradeItemMax2 = Player(index).TradeItemMax And Player(TradeIndex).TradeItemMax2 = Player(TradeIndex).TradeItemMax Then
                For I = 1 To MAX_PLAYER_TRADES
                    For X = 1 To MAX_INV
                        If GetPlayerInvItemNum(TradeIndex, X) < 1 Then
                            If Player(index).Trading(I).InvNum > 0 Then
                            If Item(GetPlayerInvItemNum(index, Player(index).Trading(I).InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, Player(index).Trading(I).InvNum)).Stackable = 1 Then
                            Call GiveItem(TradeIndex, GetPlayerInvItemNum(index, Player(index).Trading(I).InvNum), Player(index).Trading(I).InvVal)
                            Call TakeItem(index, GetPlayerInvItemNum(index, Player(index).Trading(I).InvNum), Player(index).Trading(I).InvVal)
                            Exit For
                            Else
                                Call GiveItem(TradeIndex, GetPlayerInvItemNum(index, Player(index).Trading(I).InvNum), 0)
                                Call TakeItem(index, GetPlayerInvItemNum(index, Player(index).Trading(I).InvNum), 0)
                                Exit For
                                End If
                            End If
                        End If
                    Next X
                Next I
    
                For I = 1 To MAX_PLAYER_TRADES
                    For X = 1 To MAX_INV
                        If GetPlayerInvItemNum(index, X) < 1 Then
                            If Player(TradeIndex).Trading(I).InvNum > 0 Then
                            If Item(GetPlayerInvItemNum(TradeIndex, Player(TradeIndex).Trading(I).InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(TradeIndex, Player(TradeIndex).Trading(I).InvNum)).Stackable = 1 Then
                            Call GiveItem(index, GetPlayerInvItemNum(TradeIndex, Player(TradeIndex).Trading(I).InvNum), Player(TradeIndex).Trading(I).InvVal)
                            Call TakeItem(TradeIndex, GetPlayerInvItemNum(TradeIndex, Player(TradeIndex).Trading(I).InvNum), Player(TradeIndex).Trading(I).InvVal)
                            Exit For
                            Else
                                Call GiveItem(index, GetPlayerInvItemNum(TradeIndex, Player(TradeIndex).Trading(I).InvNum), 0)
                                Call TakeItem(TradeIndex, GetPlayerInvItemNum(TradeIndex, Player(TradeIndex).Trading(I).InvNum), 0)
                                Exit For
                                End If
                        End If
                        End If
                    Next X
                Next I

                Call PlayerMsg(index, "The trade was successful!", BRIGHTGREEN)
                Call PlayerMsg(TradeIndex, "The trade was successful!", BRIGHTGREEN)

                Call SendInventory(index)
                Call SendInventory(TradeIndex)
            Else
                If Player(index).TradeItemMax2 < Player(index).TradeItemMax Then
                    Call PlayerMsg(index, "Your inventory is full!", BRIGHTRED)
                    Call PlayerMsg(TradeIndex, GetPlayerName(index) & "'s inventory is full!", BRIGHTRED)
                End If
                        
                If Player(TradeIndex).TradeItemMax2 < Player(TradeIndex).TradeItemMax Then
                    Call PlayerMsg(TradeIndex, "Your inventory is full!", BRIGHTRED)
                    Call PlayerMsg(index, GetPlayerName(TradeIndex) & "'s inventory is full!", BRIGHTRED)
                End If
            End If
    
            Player(index).TradePlayer = 0
            Player(index).InTrade = False
            Player(index).TradeOk = 0

            Player(TradeIndex).TradePlayer = 0
            Player(TradeIndex).InTrade = False
            Player(TradeIndex).TradeOk = 0

            Call SendDataTo(index, "qtrade" & END_CHAR)
            Call SendDataTo(TradeIndex, "qtrade" & END_CHAR)
        End If
    End If
End Sub

Public Sub Packet_Party(ByVal index As Long, ByVal Name As String)
    Dim PlayerIndex As Long
    Dim PartyCount As Long
    Dim I As Long

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, Name & " is currently offline.", PINK)
        Exit Sub
    End If

    If PlayerIndex = index Then
        Call PlayerMsg(index, "You cannot party with yourself!", PINK)
        Exit Sub
    End If

    If Player(index).InParty Then
        For I = 1 To MAX_PARTY_MEMBERS
            If Player(index).Party.Member(I) > 0 Then
                PartyCount = PartyCount + 1
            End If
        Next I

        If PartyCount > (MAX_PARTY_MEMBERS - 1) Then
            Call PlayerMsg(index, "Your party is full!", PINK)
            Exit Sub
        End If
    End If

    If Not Player(PlayerIndex).InParty Then
        Call PlayerMsg(index, GetPlayerName(PlayerIndex) & " has been invited to your party.", PINK)
        Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " has invited you to join their party. Type /join to join, or /leave to decline.", PINK)

        Player(PlayerIndex).InvitedBy = index
    Else
        Call PlayerMsg(index, "Player is already in a party!", PINK)
    End If
End Sub

Public Sub Packet_JoinParty(ByVal index As Long)
    Dim PlayerIndex As Long
    Dim I As Long

    PlayerIndex = Player(index).InvitedBy

    If PlayerIndex > 0 Then
        Call PlayerMsg(index, "You have joined " & GetPlayerName(PlayerIndex) & "'s party!", PINK)

        If Not Player(PlayerIndex).InParty Then
            Call SetPMember(PlayerIndex, PlayerIndex)
            Player(PlayerIndex).InParty = True
            Call SetPShare(PlayerIndex, True)
        End If

        Player(index).InParty = True
        Player(index).Party.Leader = PlayerIndex

        Call SetPMember(PlayerIndex, index)

        If GetPlayerLevel(index) + 5 < GetPlayerLevel(PlayerIndex) Or GetPlayerLevel(index) - 5 > GetPlayerLevel(PlayerIndex) Then
            Call PlayerMsg(index, "There is more then a 5 level gap between you two, you will not share experience.", PINK)
            Call PlayerMsg(PlayerIndex, "There is more then a 5 level gap between you two, you will not share experience.", PINK)
            Call SetPShare(index, False)
        Else
            Call SetPShare(index, True)
        End If

        For I = 1 To MAX_PARTY_MEMBERS
            If Player(index).Party.Member(I) > 0 Then
                If Player(index).Party.Member(I) <> index Then
                    Call PlayerMsg(Player(index).Party.Member(I), GetPlayerName(index) & " has joined your party!", PINK)
                End If
            End If
        Next I

        For I = 1 To MAX_PARTY_MEMBERS
            If Player(index).Party.Member(I) = index Then
                For PlayerIndex = 1 To MAX_PARTY_MEMBERS
                    Call SendDataTo(PlayerIndex, "updatemembers" & SEP_CHAR & I & SEP_CHAR & index & END_CHAR)
                Next PlayerIndex
            End If
        Next I

        For I = 1 To MAX_PARTY_MEMBERS
            Call SendDataTo(index, "updatemembers" & SEP_CHAR & I & SEP_CHAR & Player(index).Party.Member(I) & END_CHAR)
        Next I
    Else
        Call PlayerMsg(index, "You have not been invited into a party!", PINK)
    End If
End Sub

Public Sub Packet_LeaveParty(ByVal index As Long)
    Dim PlayerIndex As Long
    Dim I As Long

    PlayerIndex = Player(index).InvitedBy

    If PlayerIndex > 0 Or Player(index).Party.Leader = index Then
        If Player(index).InParty Then
            Call PlayerMsg(index, "You have left the party.", PINK)

            For I = 1 To MAX_PARTY_MEMBERS
                If Player(index).Party.Member(I) > 0 Then
                    Call PlayerMsg(Player(index).Party.Member(I), GetPlayerName(index) & " has left the party.", PINK)
                End If
            Next I

            Call RemovePMember(index)
            Call SendDataTo(index, "leaveparty211")
        Else
            Call PlayerMsg(index, "Declined party request.", PINK)
            Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " declined your request.", PINK)

            Player(index).InParty = False
            Player(index).InvitedBy = 0
        End If
    Else
        Call PlayerMsg(index, "You are not in a party!", PINK)
    End If
End Sub

Public Sub Packet_PartyChat(ByVal index As Long, ByVal Message As String)
    Dim I As Long

    For I = 1 To MAX_PARTY_MEMBERS
        If Player(index).Party.Member(I) > 0 Then
            Call PlayerMsg(Player(index).Party.Member(I), Message, BLUE)
        End If
    Next I
End Sub

Public Sub Packet_Spells(ByVal index As Long)
    Call SendPlayerSpells(index)
End Sub

Public Sub Packet_HotScript(ByVal index As Long, ByVal ScriptID As Long)
    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\main.ess", "HotScript " & index & "," & ScriptID
    End If
End Sub

Public Sub Packet_ScriptTile(ByVal index As Long, ByVal TileNum As Long)
    Call SendDataTo(index, "SCRIPTTILE" & SEP_CHAR & GetVar(App.Path & "\Tiles.ini", "Names", "Tile" & TileNum) & END_CHAR)
End Sub

Public Sub Packet_Cast(ByVal index As Long, ByVal SpellNum As Long)
    Call CastSpell(index, SpellNum)
End Sub

Public Sub Packet_Refresh(ByVal index As Long)
    Call SendDataToMap(GetPlayerMap(index), "playerxy" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & END_CHAR)
End Sub

Public Sub Packet_BuySprite(ByVal index As Long)
    Dim I As Long

    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
        Call PlayerMsg(index, "You need to be on a sprite tile to buy it!", BRIGHTRED)
        Exit Sub
    End If

    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 = 0 Then
        Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
        Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & END_CHAR)
        Exit Sub
    End If

    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(index, I) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 Then
            If Item(GetPlayerInvItemNum(index, I)).Type = ITEM_TYPE_CURRENCY Then
                If GetPlayerInvItemValue(index, I) >= Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3 Then
                    Call SetPlayerInvItemValue(index, I, GetPlayerInvItemValue(index, I) - Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3)

                    If GetPlayerInvItemValue(index, I) = 0 Then
                        Call SetPlayerInvItemNum(index, I, 0)
                    End If

                    Call PlayerMsg(index, "You have bought a new sprite!", BRIGHTGREEN)
                    Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                    Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & END_CHAR)
                    Call SendInventory(index)
                End If
            Else
                If GetPlayerWeaponSlot(index) <> I And GetPlayerArmorSlot(index) <> I And GetPlayerShieldSlot(index) <> I And GetPlayerHelmetSlot(index) <> I And GetPlayerLegsSlot(index) <> I And GetPlayerRingSlot(index) <> I And GetPlayerNecklaceSlot(index) <> I Then
                    Call SetPlayerInvItemNum(index, I, 0)
                    Call PlayerMsg(index, "You have bought a new sprite!", BRIGHTGREEN)
                    Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                    Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & END_CHAR)
                    Call SendInventory(index)
                End If
            End If

            If GetPlayerWeaponSlot(index) <> I And GetPlayerArmorSlot(index) <> I And GetPlayerShieldSlot(index) <> I And GetPlayerHelmetSlot(index) <> I And GetPlayerLegsSlot(index) <> I And GetPlayerRingSlot(index) <> I And GetPlayerNecklaceSlot(index) <> I Then
                Exit Sub
            End If
        End If
    Next I

    Call PlayerMsg(index, "You don't have enough to buy this sprite!", BRIGHTRED)
End Sub

Public Sub Packet_ClearOwner(ByVal index As Long)
    Dim MapNum As Long

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    MapNum = GetPlayerMap(index)

    Map(MapNum).Owner = 0
    Map(MapNum).Name = "Abandoned House"
    Map(MapNum).Revision = Map(MapNum).Revision + 1

    Call SaveMap(MapNum)
    Call MapCache_Create(MapNum)
    Call SendDataToMap(MapNum, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & (Map(MapNum).Revision + 1) & END_CHAR)

    Call PlayerMsg(index, "The house owner was successfully cleared.", BRIGHTRED)
End Sub

Public Sub Packet_RequestEditHouse(ByVal index As Long)
    If Map(GetPlayerMap(index)).Moral <> MAP_MORAL_HOUSE Then
        Call PlayerMsg(index, "This is not a house!", BRIGHTRED)
        Exit Sub
    End If

    If Map(GetPlayerMap(index)).Owner <> GetPlayerName(index) Then
        Call PlayerMsg(index, "This is not your house!", BRIGHTRED)
        Exit Sub
    End If

    Call SendDataTo(index, "EDITHOUSE" & END_CHAR)
End Sub

Public Sub Packet_BuyHouse(ByVal index As Long)
    Dim I As Long
    Dim MapNum As Long
    
    MapNum = GetPlayerMap(index)
    
    If Map(MapNum).Tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_HOUSE Then
        Call PlayerMsg(index, "You need to be on a house tile to buy it!", BRIGHTRED)
        Exit Sub
    End If

    If Map(MapNum).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 0 Then
        Map(MapNum).Owner = GetPlayerName(index)
        Map(MapNum).Name = GetPlayerName(index) & "'s House"
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        Call SaveMap(MapNum)
        Call MapCache_Create(MapNum)
        Call SendDataToMap(MapNum, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & (Map(MapNum).Revision + 1) & END_CHAR)
        Call PlayerMsg(index, "You now own this house!", BRIGHTGREEN)
        
        Exit Sub
    End If

    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(index, I) = Map(MapNum).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            If Item(GetPlayerInvItemNum(index, I)).Type = ITEM_TYPE_CURRENCY Then
                If GetPlayerInvItemValue(index, I) >= Map(MapNum).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 Then
                    Call SetPlayerInvItemValue(index, I, GetPlayerInvItemValue(index, I) - Map(MapNum).Tile(GetPlayerX(index), GetPlayerY(index)).Data2)

                    If GetPlayerInvItemValue(index, I) = 0 Then
                        Call SetPlayerInvItemNum(index, I, 0)
                    End If

                    Map(MapNum).Owner = GetPlayerName(index)
                    Map(MapNum).Name = GetPlayerName(index) & "'s House"
                    Map(MapNum).Revision = Map(GetPlayerMap(index)).Revision + 1
                    Call SaveMap(MapNum)
                    Call SendInventory(index)
                    Call PlayerMsg(index, "You have bought a new house!", BRIGHTGREEN)
                Call MapCache_Create(MapNum)
                Call SendDataToMap(MapNum, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & (Map(MapNum).Revision + 1) & END_CHAR)
                End If
            Else
                If GetPlayerWeaponSlot(index) <> I And GetPlayerArmorSlot(index) <> I And GetPlayerShieldSlot(index) <> I And GetPlayerHelmetSlot(index) <> I And GetPlayerLegsSlot(index) <> I And GetPlayerRingSlot(index) <> I And GetPlayerNecklaceSlot(index) <> I Then
                    Call SetPlayerInvItemNum(index, I, 0)

                    Map(MapNum).Owner = GetPlayerName(index)
                    Map(MapNum).Name = GetPlayerName(index) & "'s House"
                    Map(MapNum).Revision = Map(MapNum).Revision + 1

                    Call SaveMap(MapNum)
                    Call SendInventory(index)
                    Call PlayerMsg(index, "You now own a new house!", BRIGHTGREEN)
                Call MapCache_Create(MapNum)
                Call SendDataToMap(MapNum, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & (Map(MapNum).Revision + 1) & END_CHAR)
                End If
            End If

            If GetPlayerWeaponSlot(index) <> I And GetPlayerArmorSlot(index) <> I And GetPlayerShieldSlot(index) <> I And GetPlayerHelmetSlot(index) <> I And GetPlayerLegsSlot(index) <> I And GetPlayerRingSlot(index) <> I And GetPlayerNecklaceSlot(index) <> I Then
                Exit Sub
            End If
        End If
    Next I

    Call PlayerMsg(index, "You don't have enough to buy this house!", BRIGHTRED)
End Sub

Public Sub Packet_SellHouse(ByVal index As Long)
    Dim I As Long

    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_HOUSE Then
        Call PlayerMsg(index, "You need to be on the buy house tile to sell it!", BRIGHTRED)
        Exit Sub
    End If

    If Map(GetPlayerMap(index)).Owner <> GetPlayerName(index) Then
        Call PlayerMsg(index, "You don't own this house!", BRIGHTRED)
        Exit Sub
    End If
    
    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(index, I) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            If Item(GetPlayerInvItemNum(index, I)).Type = ITEM_TYPE_CURRENCY Then
                Call SetPlayerInvItemValue(index, I, GetPlayerInvItemValue(index, I) + (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 / 2))
    
                If GetPlayerInvItemValue(index, I) = 0 Then
                    Call SetPlayerInvItemNum(index, I, 0)
                End If
                Call SendInventory(index)
                     
                Call PlayerMsg(index, "You have sold this house for " & (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 / 2) & "!", BRIGHTGREEN)
                Call Packet_ClearOwner(index)
            End If
        End If
    Next I
End Sub


Public Sub Packet_CheckCommands(ByVal index As Long, ByVal Command As String)
    If SCRIPTING = 1 Then
        PutVar App.Path & "\Scripts\Command.ini", "TEMP", "Text" & index, Trim$(Command)
        MyScript.ExecuteStatement "Scripts\main.ess", "Commands " & index
    Else
        Call PlayerMsg(index, "That is not a valid command!", BRIGHTRED)
    End If
End Sub

Public Sub Packet_Prompt(ByVal index As Long, ByVal PromptNum As Long, ByVal Value As Long)
    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\main.ess", "PlayerPrompt " & index & "," & PromptNum & "," & Value
    End If
End Sub

Public Sub Packet_QueryBox(ByVal index As Long, ByVal Response As String, ByVal PromptNum As Long)
    If SCRIPTING = 1 Then
        Call PutVar(App.Path & "\Responses.ini", "Responses", CStr(index), Response)
        MyScript.ExecuteStatement "Scripts\main.ess", "QueryBox " & index & "," & PromptNum
    End If
End Sub

Public Sub Packet_RequestEditArrow(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "arroweditor" & END_CHAR)
End Sub

Public Sub Packet_EditArrow(ByVal index As Long, ByVal ArrowNum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ArrowNum < 0 Or ArrowNum > MAX_ARROWS Then
        Call HackingAttempt(index, "Invalid Arrow Index")
        Exit Sub
    End If

    Call SendEditArrowTo(index, ArrowNum)

    Call AddLog(GetPlayerName(index) & " editing arrow #" & ArrowNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveArrow(ByVal index As Long, ByVal ArrowNum As Long, ByVal Name As String, ByVal Pic As Long, ByVal Range As Long, ByVal Amount As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ArrowNum < 0 Or ArrowNum > MAX_ITEMS Then
        Call HackingAttempt(index, "Invalid Arrow Index")
        Exit Sub
    End If

    Arrows(ArrowNum).Name = Name
    Arrows(ArrowNum).Pic = Pic
    Arrows(ArrowNum).Range = Range
    Arrows(ArrowNum).Amount = Amount

    Call SendUpdateArrowToAll(ArrowNum)
    Call SaveArrow(ArrowNum)

    Call AddLog(GetPlayerName(index) & " saved arrow #" & ArrowNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_CheckArrows(ByVal index As Long, ByVal ArrowNum As Long)
    Call SendDataToMap(GetPlayerMap(index), "checkarrows" & SEP_CHAR & index & SEP_CHAR & Arrows(ArrowNum).Pic & END_CHAR)
End Sub

Public Sub Packet_RequestEditEmoticon(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "emoticoneditor" & END_CHAR)
End Sub

Public Sub Packet_RequestEditElement(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "elementeditor" & END_CHAR)
End Sub

Public Sub Packet_RequestEditQuest(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "questeditor" & END_CHAR)
End Sub

Public Sub Packet_EditEmoticon(ByVal index As Long, ByVal EmoteNum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If EmoteNum < 0 Or EmoteNum > MAX_EMOTICONS Then
        Call HackingAttempt(index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Call SendEditEmoticonTo(index, EmoteNum)

    Call AddLog(GetPlayerName(index) & " editing emoticon #" & EmoteNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_EditElement(ByVal index As Long, ByVal ElementNum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ElementNum < 0 Or ElementNum > MAX_ELEMENTS Then
        Call HackingAttempt(index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Call SendEditElementTo(index, ElementNum)

    Call AddLog(GetPlayerName(index) & " editing element #" & ElementNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveEmoticon(ByVal index As Long, ByVal EmoteNum As Long, ByVal Command As String, ByVal Pic As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If EmoteNum < 0 Or EmoteNum > MAX_EMOTICONS Then
        Call HackingAttempt(index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Emoticons(EmoteNum).Command = Command
    Emoticons(EmoteNum).Pic = Pic

    Call SendUpdateEmoticonToAll(EmoteNum)
    Call SaveEmoticon(EmoteNum)

    Call AddLog(GetPlayerName(index) & " saved emoticon #" & EmoteNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveElement(ByVal index As Long, ByVal ElementNum As Long, ByVal Name As String, ByVal Strong As Long, ByVal Weak As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ElementNum < 0 Or ElementNum > MAX_ELEMENTS Then
        Call HackingAttempt(index, "Invalid Element Index")
        Exit Sub
    End If

    Element(ElementNum).Name = Name
    Element(ElementNum).Strong = Strong
    Element(ElementNum).Weak = Weak

    Call SendUpdateElementToAll(ElementNum)
    Call SaveElement(ElementNum)

    Call AddLog(GetPlayerName(index) & " saved element #" & ElementNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_CheckEmoticon(ByVal index As Long, ByVal EmoteNum As Long)
    Call SendDataToMap(GetPlayerMap(index), "checkemoticons" & SEP_CHAR & index & SEP_CHAR & Emoticons(EmoteNum).Pic & END_CHAR)
End Sub

Public Sub Packet_MapReport(ByVal index As Long)
    Dim packet As String
    Dim I As Long

    If GetPlayerAccess(index) < ADMIN_MONITER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If

    packet = "mapreport" & SEP_CHAR

    For I = 1 To MAX_MAPS
        packet = packet & Map(I).Name & SEP_CHAR
    Next I

    packet = packet & END_CHAR

    Call SendDataTo(index, packet)
End Sub

Public Sub Packet_GMTime(ByVal index As Long, ByVal Time As Long)
    If GetPlayerAccess(index) < ADMIN_MONITER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If

    GameTime = Time

    Call SendTimeToAll
End Sub

Public Sub Packet_Weather(ByVal index As Long, ByVal WeatherNum As Long)
    If GetPlayerAccess(index) < ADMIN_MONITER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If

    WeatherType = WeatherNum

    Call SendWeatherToAll
End Sub

Public Sub Packet_WarpTo(ByVal index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    If GetPlayerAccess(index) < ADMIN_MONITER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If
    
    If X < 0 Or X > MAX_MAPX Then
        Call PlayerMsg(index, "Please enter a valid X coordinate.", BRIGHTRED)
        Exit Sub
    End If

    If Y < 0 Or Y > MAX_MAPY Then
        Call PlayerMsg(index, "Please enter a valid Y coordinate.", BRIGHTRED)
        Exit Sub
    End If

    Call PlayerWarp(index, MapNum, X, Y)
End Sub

Public Sub Packet_LocalWarp(ByVal index As Long, ByVal X As Long, ByVal Y As Long)
    If GetPlayerAccess(index) < ADMIN_MONITER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If
    
    If X < 0 Or X > MAX_MAPX Then
        Call PlayerMsg(index, "Please enter a valid X coordinate.", BRIGHTRED)
        Exit Sub
    End If

    If Y < 0 Or Y > MAX_MAPY Then
        Call PlayerMsg(index, "Please enter a valid Y coordinate.", BRIGHTRED)
        Exit Sub
    End If

    Player(index).Char(Player(index).CharNum).X = X
    Player(index).Char(Player(index).CharNum).Y = Y

    Call SendPlayerXY(index)
End Sub

Public Sub Packet_ArrowHit(ByVal index As Long, ByVal TargetType As Long, ByVal PlayerIndex As Long, ByVal X As Long, ByVal Y As Long)
    Dim Damage As Long
    
    If TargetType = TARGET_TYPE_PLAYER Then
        If PlayerIndex <> index Then
            If CanAttackPlayerWithArrow(index, PlayerIndex) Then
                Player(index).Target = PlayerIndex
                Player(index).TargetType = TARGET_TYPE_PLAYER
                If Not CanPlayerBlockHit(PlayerIndex) Then
                    If Not CanPlayerCriticalHit(index) Then
                        Damage = GetPlayerDamage(index) - GetPlayerProtection(PlayerIndex)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & END_CHAR)
                    Else
                        TargetType = GetPlayerDamage(index)
                        Damage = TargetType + Int(Rnd * Int(TargetType / 2)) + 1 - GetPlayerProtection(PlayerIndex)

                        Call BattleMsg(index, "You feel a surge of energy upon shooting!", BRIGHTCYAN, 0)
                        Call BattleMsg(PlayerIndex, GetPlayerName(index) & " shoots With amazing accuracy!", BRIGHTCYAN, 1)

                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & END_CHAR)
                    End If

                    If Damage > 0 Then
                        If SCRIPTING = 1 Then
                            MyScript.ExecuteStatement "Scripts\main.ess", "OnArrowHit " & index & "," & Damage
                        Else
                            Call AttackPlayer(index, PlayerIndex, Damage)
                        End If
                    Else
                        If SCRIPTING = 1 Then
                            MyScript.ExecuteStatement "Scripts\main.ess", "OnArrowHit " & index & "," & 0
                        End If
                        Call BattleMsg(index, "Your attack does nothing.", BRIGHTRED, 0)
                        Call BattleMsg(PlayerIndex, GetPlayerName(index) & "'s attack did nothing.", BRIGHTRED, 1)

                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                    End If
                Else
                
                    If SCRIPTING = 1 Then
                        MyScript.ExecuteStatement "Scripts\main.ess", "OnArrowHit " & index & "," & 0
                    End If
                    Call BattleMsg(index, GetPlayerName(PlayerIndex) & " blocked your hit!", BRIGHTCYAN, 0)
                    Call BattleMsg(PlayerIndex, "You blocked " & GetPlayerName(index) & "'s hit!", BRIGHTCYAN, 1)

                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                End If

                Exit Sub
            End If
        End If
    ElseIf TargetType = TARGET_TYPE_NPC Then
        If CanAttackNpcWithArrow(index, PlayerIndex) Then
        Player(index).TargetType = TARGET_TYPE_NPC
        Player(index).TargetNPC = PlayerIndex
            If Not CanPlayerCriticalHit(index) Then
                Damage = GetPlayerDamage(index) - Int(NPC(MapNPC(GetPlayerMap(index), PlayerIndex).num).DEF / 2)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & END_CHAR)
            Else
                TargetType = GetPlayerDamage(index)
                Damage = TargetType + Int(Rnd * Int(TargetType / 2)) + 1 - Int(NPC(MapNPC(GetPlayerMap(index), PlayerIndex).num).DEF / 2)

                Call BattleMsg(index, "You feel a surge of energy upon shooting!", BRIGHTCYAN, 0)

                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & END_CHAR)
            End If

            If Damage > 0 Then
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\main.ess", "OnArrowHit " & index & "," & Damage
                Else
                    Call AttackNpc(index, PlayerIndex, Damage)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & PlayerIndex & END_CHAR)
                End If
            Else
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\main.ess", "OnArrowHit " & index & "," & Damage
                End If
                Call BattleMsg(index, "Your attack does nothing.", BRIGHTRED, 0)

                Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & PlayerIndex & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
            End If

            Exit Sub
        End If
    End If
End Sub

Public Sub Packet_BankDeposit(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim BankSlot As Long
    Dim ItemNum As Long

    ItemNum = GetPlayerInvItemNum(index, InvNum)

    BankSlot = FindOpenBankSlot(index, ItemNum)
    If BankSlot = 0 Then
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Bank full!" & END_CHAR)
        Exit Sub
    End If

    If Amount > GetPlayerInvItemValue(index, InvNum) Then
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You can't deposit more than you have!" & END_CHAR)
        Exit Sub
    End If

    If GetPlayerWeaponSlot(index) = ItemNum Or GetPlayerArmorSlot(index) = ItemNum Or GetPlayerShieldSlot(index) = ItemNum Or GetPlayerHelmetSlot(index) = ItemNum Or GetPlayerLegsSlot(index) = ItemNum Or GetPlayerRingSlot(index) = ItemNum Or GetPlayerNecklaceSlot(index) = ItemNum Then
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You can't deposit worn equipment!" & END_CHAR)
        Exit Sub
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        If Amount = 0 Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You must deposit more than 0!" & END_CHAR)
            Exit Sub
        End If
    End If

    Call TakeItem(index, ItemNum, Amount)
    Call GiveBankItem(index, ItemNum, Amount, BankSlot)

    Call SendBank(index)
End Sub

Public Sub Packet_BankWithdraw(ByVal index As Long, ByVal BankInvNum As Long, ByVal Amount As Long)
    Dim BankItemNum As Long
    Dim BankInvSlot As Long

    BankItemNum = GetPlayerBankItemNum(index, BankInvNum)

    BankInvSlot = FindOpenInvSlot(index, BankItemNum)
    If BankInvSlot = 0 Then
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Inventory full!" & END_CHAR)
        Exit Sub
    End If

    If Amount > GetPlayerBankItemValue(index, BankInvNum) Then
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You can't withdraw more than you have!" & END_CHAR)
        Exit Sub
    End If

    If Item(BankItemNum).Type = ITEM_TYPE_CURRENCY Or Item(BankItemNum).Stackable = 1 Then
        If Amount = 0 Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "You must withdraw more than 0!" & END_CHAR)
            Exit Sub
        End If
    End If

    Call TakeBankItem(index, BankItemNum, Amount)
    Call GiveItem(index, BankItemNum, Amount)

    Call SendBank(index)
End Sub

Public Sub Packet_ReloadScripts(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If

    Set MyScript = Nothing
    Set clsScriptCommands = Nothing

    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands

    MyScript.ReadInCode App.Path & "\Scripts\main.ess", "Scripts\main.ess", MyScript.SControl
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True

    MyScript.ExecuteStatement "Scripts\main.ess", "OnScriptReload"

    Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
    Call AdminMsg("Scripts reloaded by " & GetPlayerName(index) & ".", WHITE)
End Sub

Public Sub Packet_CustomMenuClick(ByVal index As Long, ByVal MenuIndex As Long, ByVal ClickIndex As Long, ByVal CustomTitle As String, ByVal MenuType As Long, ByVal CustomMsg As String)
    Player(index).CustomTitle = CustomTitle
    Player(index).CustomMsg = CustomMsg

    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\main.ess", "menuscripts " & MenuIndex & "," & ClickIndex & "," & MenuType
    End If
End Sub

Public Sub Packet_CustomBoxReturnMsg(ByVal index As Long, ByVal CustomMsg As String)
    Player(index).CustomMsg = CustomMsg
End Sub

Public Sub Packet_RequestEditMain(ByVal index As Long, file As String)
    Dim f
    Dim text
    
    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If
    
    If FileExists("\Scripts\" & CStr(file)) Then
        f = FreeFile
        Open (App.Path & "\Scripts\" & CStr(file)) For Input As #f
        
        text = Input$(LOF(f), f)
        Close #f
        
        Call SendDataTo(index, "MAINEDITOR" & SEP_CHAR & file & SEP_CHAR & text & SEP_CHAR & END_CHAR)
    ElseIf FileExists("\Scripts\Events\" & CStr(file)) Then
        f = FreeFile
        Open (App.Path & "\Scripts\Events\" & CStr(file)) For Input As #f
        
        text = Input$(LOF(f), f)
        Close #f
        
        Call SendDataTo(index, "MAINEDITOR" & SEP_CHAR & file & SEP_CHAR & text & SEP_CHAR & END_CHAR)
    ElseIf FileExists("\Scripts\Functions\" & CStr(file)) Then
        f = FreeFile
        Open (App.Path & "\Scripts\Functions\" & CStr(file)) For Input As #f
        
        text = Input$(LOF(f), f)
        Close #f
        
        Call SendDataTo(index, "MAINEDITOR" & SEP_CHAR & file & SEP_CHAR & text & SEP_CHAR & END_CHAR)
    Else
        Call PlayerMsg(index, "Script file not found...", BRIGHTRED)
    End If
      
End Sub

Public Sub Packet_NewMain(ByVal index As Long, FileName As String, FileContents)
    If GetPlayerAccess(index) >= ADMIN_CREATOR Then
        Dim temp As String
        Dim f
        
        If FileExists("\Scripts\" & FileName) Then
            f = FreeFile
            Open App.Path & "\Scripts\" & FileName For Input As #f
            temp = Input$(LOF(f), f)
            Close #f
            f = FreeFile
            Open App.Path & "\Scripts\Backup.txt" For Output As #f
            Print #f, temp
            Close #f
            f = FreeFile
            Open App.Path & "\Scripts\" & FileName For Output As #f
            Print #f, FileContents
            Close #f
        ElseIf FileExists("\Scripts\Events\" & FileName) Then
            f = FreeFile
            Open App.Path & "\Scripts\Events\" & FileName For Input As #f
            temp = Input$(LOF(f), f)
            Close #f
            f = FreeFile
            Open App.Path & "\Scripts\Events\Backup.txt" For Output As #f
            Print #f, temp
            Close #f
            f = FreeFile
            Open App.Path & "\Scripts\Events\" & FileName For Output As #f
            Print #f, FileContents
            Close #f
        ElseIf FileExists("\Scripts\Functions\" & FileName) Then
            f = FreeFile
            Open App.Path & "\Scripts\Functions\" & FileName For Input As #f
            temp = Input$(LOF(f), f)
            Close #f
            f = FreeFile
            Open App.Path & "\Scripts\Functions\Backup.txt" For Output As #f
            Print #f, temp
            Close #f
            f = FreeFile
            Open App.Path & "\Scripts\Functions\" & FileName For Output As #f
            Print #f, FileContents
            Close #f
        Else
            Call PlayerMsg(index, "Script file not found...", BRIGHTRED)
        End If

        If SCRIPTING = 1 Then
            Set MyScript = Nothing
            Set clsScriptCommands = Nothing
            Set MyScript = New clsSadScript
            Set clsScriptCommands = New clsCommands
            MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
        End If
        
        Call Packet_ReloadScripts(index)
        Call AddLog(GetPlayerName(index) & " updated the script.", ADMIN_LOG)
    End If
End Sub

Public Sub Packet_GuildMsg(ByVal index As Long, ByVal Msg As String)
    Dim I
    
    ' Prevent hacking
    For I = 1 To Len(Msg)
        If Asc(Mid$(Msg, I, 1)) < 32 Or Asc(Mid$(Msg, I, 1)) > 255 Then
            Call HackingAttempt(index, "Broadcast Text Modification")
            Exit Sub
        End If
    Next I
    
    If GetPlayerGuild(index) = "" Then
        Call PlayerMsg(index, "your not in a crew...", BRIGHTRED)
        Exit Sub
    End If
            
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
        If GetPlayerGuild(index) = GetPlayerGuild(I) And GetPlayerGuild(I) <> "" Then
            Call PlayerMsg(I, "[Crew : " & GetPlayerName(index) & "]: " & Msg, GREEN)
        End If
        End If
    Next I

End Sub

Public Sub Packet_SendEmail(ByVal index As Long, ByVal msgSender As String, ByVal msgReciever As String, ByVal msgSubject As String, ByVal msgBody As String)
    Dim FileName As String
    Dim Inbox As Long
    Dim Total As Long
    
    If GetPlayerName(index) <> msgSender Then
        Call HackingAttempt(index, "Email Modification")
        Exit Sub
    End If

    FileName = App.Path & "\Mail\" & msgReciever & ".txt"
    If Not FileExists("\Mail\" & msgReciever & ".txt") Then
        Call PutVar(FileName, "INBOX", "amount", 1)
        Call PutVar(FileName, "INBOX", "total", 1)
        Call PutVar(FileName, 1, "Number", 1)
        Call PutVar(FileName, 1, "Read", 1)
        Call PutVar(FileName, 1, "Sender", msgSender)
        Call PutVar(FileName, 1, "Subject", msgSubject)
        Call PutVar(FileName, 1, "Body", msgBody)
        If IsPlaying(FindPlayer(msgReciever)) Then
            If FindPlayer(msgReciever) <> index Then
                Call Packet_Unread(FindPlayer(msgReciever), msgReciever)
            Else
                Call Packet_Unread(index, msgSender)
            End If
        End If
    Else
        If Inbox = 9999 Or Total = 9999 Then
            Exit Sub
        End If
        Inbox = GetVar(FileName, "INBOX", "amount")
        Total = GetVar(FileName, "INBOX", "total")
        
        Call PutVar(FileName, Inbox + 1, "Number", Total + 1)
        Call PutVar(FileName, Inbox + 1, "Read", 1)
        Call PutVar(FileName, Inbox + 1, "Sender", msgSender)
        Call PutVar(FileName, Inbox + 1, "Subject", msgSubject)
        Call PutVar(FileName, Inbox + 1, "Body", msgBody)
        Call PutVar(FileName, "INBOX", "amount", Inbox + 1)
        Call PutVar(FileName, "INBOX", "total", Total + 1)
        If IsPlaying(FindPlayer(msgReciever)) = True Then
            If FindPlayer(msgReciever) <> index Then
                Call Packet_Unread(FindPlayer(msgReciever), msgReciever)
            Else
                Call Packet_Unread(index, msgSender)
            End If
        End If
        Call Packet_SendEmail2(index, msgSender, msgReciever, msgSubject, msgBody)
        Exit Sub
    End If
End Sub

Public Sub Packet_SendEmail2(ByVal index As Long, ByVal msgSender As String, ByVal msgReciever As String, ByVal msgSubject As String, ByVal msgBody As String)
    Dim FileName As String
    Dim Inbox As Long
    Dim Total As Long
    
    If GetPlayerName(index) <> msgSender Then
        Call HackingAttempt(index, "Email Modification")
        Exit Sub
    End If
    
    FileName = App.Path & "\Mail\Outbox\" & msgSender & ".txt"
    If Not FileExists("\Mail\Outbox\" & msgSender & ".txt") Then
        Call PutVar(FileName, "OUTBOX", "amount", 1)
        Call PutVar(FileName, "OUTBOX", "total", 1)
        Call PutVar(FileName, 1, "Number", 1)
        Call PutVar(FileName, 1, "Sent", msgReciever)
        Call PutVar(FileName, 1, "Subject", msgSubject)
        Call PutVar(FileName, 1, "Body", msgBody)
    Else
        Inbox = GetVar(FileName, "OUTBOX", "amount")
        Total = GetVar(FileName, "OUTBOX", "total")
        
        If Inbox = 9999 Or Total = 9999 Then
            Exit Sub
        End If
        Call PutVar(FileName, Inbox + 1, "Number", Total + 1)
        Call PutVar(FileName, Inbox + 1, "Sent", msgReciever)
        Call PutVar(FileName, Inbox + 1, "Subject", msgSubject)
        Call PutVar(FileName, Inbox + 1, "Body", msgBody)
        Call PutVar(FileName, "OUTBOX", "amount", Inbox + 1)
        Call PutVar(FileName, "OUTBOX", "total", Total + 1)
        Exit Sub
    End If
End Sub

Public Sub Packet_EmailBody(ByVal index As Long, ByVal MyName As String, ByVal ListedMsg As Long, ByVal Either As Long)
    Dim FileName As String
    Dim Inbox As Long
    If Either = 1 Then
        FileName = App.Path & "\Mail\" & MyName & ".txt"
        Inbox = GetVar(FileName, "INBOX", "amount")
        Do While Inbox >= 1
            If GetVar(FileName, Val(ListedMsg), "Number") = Inbox Then
                Call SendDataTo(index, "setmsgbody" & SEP_CHAR & GetVar(FileName, Val(ListedMsg), "Sender") & SEP_CHAR & GetVar(FileName, Val(ListedMsg), "Subject") & SEP_CHAR & GetVar(FileName, Val(ListedMsg), "Body") & SEP_CHAR & Either & END_CHAR)
                Call PutVar(FileName, Val(ListedMsg), "Read", 0)
                Exit Do
            End If
            Inbox = Inbox - 1
        Loop
        Call Packet_Unread(index, MyName)
    Else
        FileName = App.Path & "\Mail\Outbox\" & MyName & ".txt"
        Inbox = GetVar(FileName, "OUTBOX", "amount")
        Do While Inbox >= 1
            If GetVar(FileName, Val(ListedMsg), "Number") = Inbox Then
                Call SendDataTo(index, "setmsgbody" & SEP_CHAR & GetVar(FileName, Val(ListedMsg), "Sent") & SEP_CHAR & GetVar(FileName, Val(ListedMsg), "Subject") & SEP_CHAR & GetVar(FileName, Val(ListedMsg), "Body") & SEP_CHAR & Either & END_CHAR)
                Exit Do
            End If
            Inbox = Inbox - 1
        Loop
    End If
End Sub

Public Sub Packet_RemoveMail(ByVal index As Long, ByVal MyName As String, ByVal DltMsgNum As Long, ByVal Either As Long)
    Dim FileName As String
    Dim NewAmount As Long
    
    If Either = 1 Then
        FileName = App.Path & "\Mail\" & MyName & ".txt"
    
        Call PutVar(FileName, Val(DltMsgNum), "Number", 0)
        Call PutVar(FileName, Val(DltMsgNum), "Read", 0)
        Call PutVar(FileName, Val(DltMsgNum), "Sender", vbNullString)
        Call PutVar(FileName, Val(DltMsgNum), "Subject", vbNullString)
        Call PutVar(FileName, Val(DltMsgNum), "Body", vbNullString)
        Call Packet_Unread(index, MyName)
    Else
        FileName = App.Path & "\Mail\Outbox\" & MyName & ".txt"
    
        Call PutVar(FileName, Val(DltMsgNum), "Number", 0)
        Call PutVar(FileName, Val(DltMsgNum), "Sent", vbNullString)
        Call PutVar(FileName, Val(DltMsgNum), "Subject", vbNullString)
        Call PutVar(FileName, Val(DltMsgNum), "Body", vbNullString)
    End If
End Sub

Public Sub Packet_MyInbox(ByVal index As Long, ByVal MyInbox As String, ByVal Either As Long)
    Dim FileName As String
    Dim Inbox As Long
    Dim Unread As Long
    
    If Either = 1 Then
        FileName = App.Path & "\Mail\" & MyInbox & ".txt"
        Inbox = GetVar(FileName, "INBOX", "amount")
        Do While Inbox >= 1
            If GetVar(FileName, Val(Inbox), "Number") > 0 Then
                Call SendDataTo(index, "xobniym" & SEP_CHAR & GetVar(FileName, Val(Inbox), "Sender") & SEP_CHAR & GetVar(FileName, Val(Inbox), "Number") & SEP_CHAR & GetVar(FileName, Val(Inbox), "Subject") & END_CHAR)
            End If
            Inbox = Inbox - 1
        Loop
        Call Packet_Unread(index, MyInbox)
    Else
        FileName = App.Path & "\Mail\Outbox\" & MyInbox & ".txt"
        Inbox = GetVar(FileName, "OUTBOX", "amount")
        Do While Inbox >= 1
            If GetVar(FileName, Val(Inbox), "Number") > 0 Then
                Call SendDataTo(index, "myoutbox" & SEP_CHAR & GetVar(FileName, Val(Inbox), "Sent") & SEP_CHAR & GetVar(FileName, Val(Inbox), "Number") & SEP_CHAR & GetVar(FileName, Val(Inbox), "Subject") & END_CHAR)
            End If
            Inbox = Inbox - 1
        Loop
    End If
    
End Sub

Public Sub Packet_Unread(ByVal index As Long, ByVal MyInbox As String)
    Dim Unread As Long
    Dim Inbox As String
    Dim FileName As String
    FileName = App.Path & "\Mail\" & MyInbox & ".txt"
    Inbox = GetVar(FileName, "INBOX", "amount")
    
    Do While Inbox >= 1
        If GetVar(FileName, Val(Inbox), "Read") = 1 Then
                Unread = Unread + 1
        End If
        Inbox = Inbox - 1
    Loop
    If Unread > 0 Then
        Call BattleMsg(index, "You Have " & Unread & " New Message!", BRIGHTRED, 0)
        Call SendDataTo(index, "sound" & SEP_CHAR & "newmsg" & END_CHAR)
    End If
    Call SendDataTo(index, "unreadmsg" & SEP_CHAR & Unread & END_CHAR)
End Sub

Public Sub Packet_CheckChar(ByVal index As Long, ByVal msgSender As String, ByVal msgReciever As String, ByVal msgSubject As String, ByVal msgBody As String)
    If FindChar(msgReciever) = True Then
        Call Packet_SendEmail(index, msgSender, msgReciever, msgSubject, msgBody)
    Else
        Call PlayerMsg(index, "That Player Does Not Exist!", BRIGHTRED)
    End If
End Sub


