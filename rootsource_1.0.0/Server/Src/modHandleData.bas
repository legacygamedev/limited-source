Attribute VB_Name = "modHandleData"
Option Explicit

' ********************************************
' **               rootSource               **
' ********************************************
Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CGetClasses) = GetAddress(AddressOf HandleGetClasses)
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CDelChar) = GetAddress(AddressOf HandleDelChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(CAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CEditItem) = GetAddress(AddressOf HandleEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CDelete) = GetAddress(AddressOf HandleDelete)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CEditNpc) = GetAddress(AddressOf HandleEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CEditShop) = GetAddress(AddressOf HandleEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditSpell)
    HandleDataSub(CEditSpell) = GetAddress(AddressOf HandleEditSpell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CTrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CFixItem) = GetAddress(AddressOf HandleFixItem)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CParty) = GetAddress(AddressOf HandleParty)
    HandleDataSub(CJoinParty) = GetAddress(AddressOf HandleJoinParty)
    HandleDataSub(CLeaveParty) = GetAddress(AddressOf HandleLeaveParty)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSync) = GetAddress(AddressOf HandleSync)
    HandleDataSub(CMapReqs) = GetAddress(AddressOf HandleMapReqs)
    HandleDataSub(CSleepinn) = GetAddress(AddressOf HandleSleepInn)
    HandleDataSub(CCreateGuild) = GetAddress(AddressOf HandleCreateGuild)
    HandleDataSub(CRemoveFromGuild) = GetAddress(AddressOf HandleRemoveFromGuild)
    HandleDataSub(CInviteGuild) = GetAddress(AddressOf HandleGuildInvite)
    HandleDataSub(CKickGuild) = GetAddress(AddressOf HandleGuildKick)
    HandleDataSub(CGuildPromote) = GetAddress(AddressOf HandleGuildPromote)
    HandleDataSub(CLeaveGuild) = GetAddress(AddressOf HandleLeaveGuild)
End Sub

' Will handle the packet data
Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim msgtype As Integer
        
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    msgtype = Buffer.ReadInteger
    
    If msgtype < 0 Or msgtype >= CMSG_COUNT Then
        HackingAttempt Index, "Packet Manipulation."
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(msgtype), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Requesting classes for making a character ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Private Sub HandleGetClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then
        Call SendNewCharClasses(Index)
    End If
End Sub

' ::::::::::::::::::::::::
' :: New account packet ::
' ::::::::::::::::::::::::
Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String
Dim Password As String
Dim i As Long
Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            
            Buffer.WriteBytes Data()
            
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
        
            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))
                If Not isNameLegal(n) Then
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next
            
            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                Call AlertMsg(Index, "Your account has been created!")
            Else
                Call AlertMsg(Index, "Sorry, that account name is already taken!")
            End If
            
        End If
    End If
    
End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String
Dim Password As String
Dim i As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            
            Buffer.WriteBytes Data()
            
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
            
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
                If LenB(Trim$(Player(Index).Char(i).Name)) > 0 Then
                    Call DeleteName(Player(Index).Char(i).Name)
                End If
            Next
            Call ClearPlayer(Index)
            
            ' Everything went ok
            Call Kill(App.Path & "\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(Index, "Your account has been deleted.")
            
        End If
    End If
    
End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String
Dim Password As String
Dim i As Long
Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            
            Buffer.WriteBytes Data()
            
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
        
            ' Check versions
            If Buffer.ReadByte < CLIENT_MAJOR Or Buffer.ReadByte < CLIENT_MINOR Or Buffer.ReadByte < CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit " & GAME_WEBSITE)
                Exit Sub
            End If
            
            If isShuttingDown Then
                Call AlertMsg(Index, "Server is either rebooting or being shutdown.")
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
    
            ' Load the player
            Call LoadPlayer(Index, Name)
            Call SendChars(Index)
            Call SendMaxes(Index)
            Call SendMapRevs(Index)
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
    
        End If
    End If
    
End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String
Dim Password As String
Dim Sex As Long
Dim Class As Long
Dim CharNum As Long
Dim i As Long
Dim n As Long

    If Not IsPlaying(Index) Then
        Set Buffer = New clsBuffer

        Buffer.WriteBytes Data()
        
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Class = Buffer.ReadLong
        CharNum = Buffer.ReadLong
    
        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(Index, "Character name must be at least three characters in length.")
            Exit Sub
        End If
        
        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))
            
            If Not isNameLegal(n) Then
                Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If
        Next
                                
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
    
        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Call HackingAttempt(Index, "Invalid Sex")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
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
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        Call AlertMsg(Index, "Character has been created!")
    End If
End Sub

' :::::::::::::::::::::::::::::::
' :: Deleting character packet ::
' :::::::::::::::::::::::::::::::
Private Sub HandleDelChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim CharNum As Long

    If Not IsPlaying(Index) Then
        Set Buffer = New clsBuffer

        Buffer.WriteBytes Data()
            
        CharNum = Buffer.ReadLong
    
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
        
        Call DelChar(Index, CharNum)
        Call AddLog("Character deleted on " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        Call AlertMsg(Index, "Character has been deleted!")
    End If
End Sub

' ::::::::::::::::::::::::::::
' :: Using character packet ::
' ::::::::::::::::::::::::::::
Private Sub HandleUseChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim CharNum As Long
Dim F As Long
Dim Buffer As clsBuffer

    If Not IsPlaying(Index) Then
        Set Buffer = New clsBuffer

        Buffer.WriteBytes Data()
            
        CharNum = Buffer.ReadLong
    
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
    
        ' Check to make sure the character exists and if so, set it as its current char
        If CharExist(Index, CharNum) Then
            TempPlayer(Index).CharNum = CharNum
            Call JoinGame(Index)
        
            CharNum = TempPlayer(Index).CharNum
            Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".")
            Call UpdateCaption
            
            ' Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
            If Not FindChar(GetPlayerName(Index)) Then
                F = FreeFile
                Open App.Path & "\accounts\charlist.txt" For Append As #F
                    Print #F, GetPlayerName(Index)
                Close #F
            End If
        Else
            Call AlertMsg(Index, "Character does not exist!")
        End If
    End If
End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
            
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Say Text Modification")
            Exit Sub
        End If
    Next
    
    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " says, '" & Msg & "'", SayColor)
End Sub

Private Sub HandleEmoteMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Emote Text Modification")
            Exit Sub
        End If
    Next
    
    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim s As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
            
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Broadcast Text Modification")
            Exit Sub
        End If
    Next
    
    s = GetPlayerName(Index) & ": " & Msg
    Call AddLog(s, PLAYER_LOG)
    Call GlobalMsg(s, BroadcastColor)
    Call TextAdd(s)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim s As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
            
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Global Text Modification")
            Exit Sub
        End If
    Next
    
    If GetPlayerAccess(Index) > 0 Then
        s = "(global) " & GetPlayerName(Index) & ": " & Msg
        Call AddLog(s, ADMIN_LOG)
        Call GlobalMsg(s, GlobalColor)
        Call TextAdd(s)
    End If
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
            
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Admin Text Modification")
            Exit Sub
        End If
    Next
    
    If GetPlayerAccess(Index) > 0 Then
        Call AddLog("(admin " & GetPlayerName(Index) & ") " & Msg, ADMIN_LOG)
        Call AdminMsg("(admin " & GetPlayerName(Index) & ") " & Msg, AdminColor)
    End If
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim i As Long
Dim MsgTo As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
            
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Player Msg Text Modification")
            Exit Sub
        End If
    Next
    
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
        Call PlayerMsg(GetPlayerName(Index), "Cannot message yourself.", BrightRed)
    End If
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Dir As Long
Dim Movement As Long
Dim Buffer As clsBuffer

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
                
    Dir = Buffer.ReadLong
    Movement = Buffer.ReadLong
    
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
    If TempPlayer(Index).CastedSpell = YES Then
        ' Check if they have already casted a spell, and if so we can't let them move
        If GetTickCount > TempPlayer(Index).AttackTimer + 1000 Then
            TempPlayer(Index).CastedSpell = NO
        Else
            Call SendPlayerXY(Index)
            Exit Sub
        End If
    End If
    
    Call PlayerMove(Index, Dir, Movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Dir As Long
Dim Buffer As clsBuffer

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    Dir = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If
    
    Call SetPlayerDir(Index, Dir)
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 10
    Buffer.WriteInteger SPlayerDir
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerDir(Index)
    
    Call SendDataToMapBut(Index, GetPlayerMap(Index), Buffer.ToArray())
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Private Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim InvNum As Long
Dim ItemNum As Long
Dim i As Long
Dim n As Long
Dim X As Long
Dim y As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    InvNum = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid InvNum")
        Exit Sub
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, InvNum)
    
    If (ItemNum > 0) And (ItemNum <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
        
        ' Find out what kind of item it is
        Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
            Case ITEM_TYPE_ARMOR
                If InvNum <> GetPlayerEquipmentSlot(Index, Armor) Then
                    If GetPlayerStat(Index, Stats.Defense) < n Then
                        Call PlayerMsg(Index, "Your defense is to low to wear this armor!  Required DEF (" & n & ")", BrightRed)
                        Exit Sub
                    End If
                    Call SetPlayerEquipmentSlot(Index, InvNum, Armor)
                Else
                    Call SetPlayerEquipmentSlot(Index, 0, Armor)
                End If
                Call SendWornEquipment(Index)
            
            Case ITEM_TYPE_WEAPON
                If InvNum <> GetPlayerEquipmentSlot(Index, Weapon) Then
                    If GetPlayerStat(Index, Stats.Strength) < n Then
                        Call PlayerMsg(Index, "Your strength is to low to hold this weapon!  Required STR (" & n & ")", BrightRed)
                        Exit Sub
                    End If
                    Call SetPlayerEquipmentSlot(Index, InvNum, Weapon)
                Else
                    Call SetPlayerEquipmentSlot(Index, 0, Weapon)
                End If
                Call SendWornEquipment(Index)
                    
            Case ITEM_TYPE_HELMET
                If InvNum <> GetPlayerEquipmentSlot(Index, Helmet) Then
                    If GetPlayerStat(Index, Stats.Speed) < n Then
                        Call PlayerMsg(Index, "Your speed coordination is to low to wear this helmet!  Required SPEED (" & n & ")", BrightRed)
                        Exit Sub
                    End If
                    Call SetPlayerEquipmentSlot(Index, InvNum, Helmet)
                Else
                    Call SetPlayerEquipmentSlot(Index, 0, Helmet)
                End If
                Call SendWornEquipment(Index)
        
            Case ITEM_TYPE_SHIELD
                If InvNum <> GetPlayerEquipmentSlot(Index, Shield) Then
                    Call SetPlayerEquipmentSlot(Index, InvNum, Shield)
                Else
                    Call SetPlayerEquipmentSlot(Index, 0, Shield)
                End If
                Call SendWornEquipment(Index)
        
            Case ITEM_TYPE_POTIONADDHP
                Call SetPlayerVital(Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) + Item(ItemNum).Data1)
                Call TakeItem(Index, ItemNum, 0)
                Call SendVital(Index, Vitals.HP)
        
            Case ITEM_TYPE_POTIONADDMP
                Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) + Item(ItemNum).Data1)
                Call TakeItem(Index, ItemNum, 0)
                Call SendVital(Index, Vitals.MP)
    
            Case ITEM_TYPE_POTIONADDSP
                Call SetPlayerVital(Index, Vitals.SP, GetPlayerVital(Index, Vitals.SP) + Item(ItemNum).Data1)
                Call TakeItem(Index, ItemNum, 0)
                Call SendVital(Index, Vitals.SP)

            Case ITEM_TYPE_POTIONSUBHP
                Call SetPlayerVital(Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) - Item(ItemNum).Data1)
                Call TakeItem(Index, ItemNum, 0)
                Call SendVital(Index, Vitals.HP)
            
            Case ITEM_TYPE_POTIONSUBMP
                Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - Item(ItemNum).Data1)
                Call TakeItem(Index, ItemNum, 0)
                Call SendVital(Index, Vitals.MP)
    
            Case ITEM_TYPE_POTIONSUBSP
                Call SetPlayerVital(Index, Vitals.SP, GetPlayerVital(Index, Vitals.SP) - Item(ItemNum).Data1)
                Call TakeItem(Index, ItemNum, 0)
                Call SendVital(Index, Vitals.SP)
                
            Case ITEM_TYPE_KEY
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        If GetPlayerY(Index) > 0 Then
                            X = GetPlayerX(Index)
                            y = GetPlayerY(Index) - 1
                        Else
                            Exit Sub
                        End If
                        
                    Case DIR_DOWN
                        If GetPlayerY(Index) < MAX_MAPY Then
                            X = GetPlayerX(Index)
                            y = GetPlayerY(Index) + 1
                        Else
                            Exit Sub
                        End If
                            
                    Case DIR_LEFT
                        If GetPlayerX(Index) > 0 Then
                            X = GetPlayerX(Index) - 1
                            y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If
                            
                    Case DIR_RIGHT
                        If GetPlayerX(Index) < MAX_MAPX Then
                            X = GetPlayerX(Index) + 1
                            y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If
                End Select
                
                ' Check if a key exists
                If Map(GetPlayerMap(Index)).Tile(X, y).Type = TILE_TYPE_KEY Then
                    ' Check if the key they are using matches the map key
                    If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(X, y).Data1 Then
                        TempTile(GetPlayerMap(Index)).DoorOpen(X, y) = YES
                        TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                        
                        Set Buffer = New clsBuffer
                        Buffer.PreAllocate 14
                        Buffer.WriteInteger SMapKey
                        Buffer.WriteLong X
                        Buffer.WriteLong y
                        Buffer.WriteLong 1
                        Call SendDataToMap(GetPlayerMap(Index), Buffer.ToArray())
                        Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
                        
                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(Index)).Tile(X, y).Data2 = 1 Then
                            Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                            Call PlayerMsg(Index, "The key disolves.", Yellow)
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
                        i = Spell(n).LevelReq
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
                                Else
                                    Call PlayerMsg(Index, "You have already learned this spell!", BrightRed)
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
                
        End Select
    End If
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim n As Long
Dim Damage As Long
Dim TempIndex As Long

    ' Try to attack a player
    For i = 1 To TotalPlayersOnline
        TempIndex = PlayersOnline(i)
        ' Make sure we dont try to attack ourselves
        If TempIndex <> Index Then
            ' Can we attack the player?
            If CanAttackPlayer(Index, TempIndex) Then
                If Not CanPlayerBlockHit(TempIndex) Then
                    ' Get the damage we can do
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = GetPlayerDamage(Index) - GetPlayerProtection(TempIndex)
                    Else
                        n = GetPlayerDamage(Index)
                        Damage = n + Int(Rnd * (n \ 2)) + 1 - GetPlayerProtection(TempIndex)
                        Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                        Call PlayerMsg(TempIndex, GetPlayerName(Index) & " swings with enormous might!", BrightCyan)
                    End If

                    Call AttackPlayer(Index, TempIndex, Damage)

                Else
                    Call PlayerMsg(Index, GetPlayerName(TempIndex) & "'s " & Trim$(Item(GetPlayerInvItemNum(TempIndex, GetPlayerEquipmentSlot(TempIndex, Shield))).Name) & " has blocked your hit!", BrightCyan)
                    Call PlayerMsg(TempIndex, "Your " & Trim$(Item(GetPlayerInvItemNum(TempIndex, GetPlayerEquipmentSlot(TempIndex, Shield))).Name) & " has blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan)
                End If
                
                Exit Sub
            End If
        End If
    Next
    
    ' Try to attack a Npc
    For i = 1 To MAX_MAP_NPCS
        ' Can we attack the Npc?
        If CanAttackNpc(Index, i) Then
            ' Get the damage we can do
            If Not CanPlayerCriticalHit(Index) Then
                Damage = GetPlayerDamage(Index) - (Npc(MapNpc(GetPlayerMap(Index), i).Num).Stat(Stats.Defense) \ 2)
            Else
                n = GetPlayerDamage(Index)
                Damage = n + Int(Rnd * (n \ 2)) + 1 - (Npc(MapNpc(GetPlayerMap(Index), i).Num).Stat(Stats.Defense) \ 2)
                Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
            End If
            
            If Damage > 0 Then
                Call AttackNpc(Index, i, Damage)
            Else
                Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
            End If
            Exit Sub
        End If
    Next
    
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Private Sub HandleUseStatPoint(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
                
    PointType = Buffer.ReadLong
    
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
                Call SetPlayerStat(Index, Stats.Strength, GetPlayerStat(Index, Stats.Strength) + 1)
                Call PlayerMsg(Index, "You have gained more strength!", White)
            Case 1
                Call SetPlayerStat(Index, Stats.Defense, GetPlayerStat(Index, Stats.Defense) + 1)
                Call PlayerMsg(Index, "You have gained more defense!", White)
            Case 2
                Call SetPlayerStat(Index, Stats.Magic, GetPlayerStat(Index, Stats.Magic) + 1)
                Call PlayerMsg(Index, "You have gained more magic abilities!", White)
            Case 3
                Call SetPlayerStat(Index, Stats.Speed, GetPlayerStat(Index, Stats.Speed) + 1)
                Call PlayerMsg(Index, "You have gained more speed!", White)
        End Select
    Else
        Call PlayerMsg(Index, "You have no skill points to train with!", BrightRed)
    End If
    
    ' Send the update
    Call SendStats(Index)
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Private Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Name As String
Dim i As Long
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
            
    Name = Buffer.ReadString
    
    i = FindPlayer(Name)
    If i > 0 Then
        Call PlayerMsg(Index, "Account: " & Trim$(Player(i).Login) & ", Name: " & GetPlayerName(i), BrightGreen)
        If GetPlayerAccess(Index) > ADMIN_MONITOR Then
            Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
            Call PlayerMsg(Index, "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), BrightGreen)
            Call PlayerMsg(Index, "HP: " & GetPlayerVital(i, Vitals.HP) & "/" & GetPlayerMaxVital(i, Vitals.HP) & "  MP: " & GetPlayerVital(i, Vitals.MP) & "/" & GetPlayerMaxVital(i, Vitals.MP) & "  SP: " & GetPlayerVital(i, Vitals.SP) & "/" & GetPlayerMaxVital(i, Vitals.SP), BrightGreen)
            Call PlayerMsg(Index, "Strength: " & GetPlayerStat(i, Stats.Strength) & "  Defense: " & GetPlayerStat(i, Stats.Defense) & "  Magic: " & GetPlayerStat(i, Stats.Magic) & "  Speed: " & GetPlayerStat(i, Stats.Speed), BrightGreen)
            n = (GetPlayerStat(i, Stats.Strength) \ 2) + (GetPlayerLevel(i) \ 2)
            i = (GetPlayerStat(i, Stats.Defense) \ 2) + (GetPlayerLevel(i) \ 2)
            If n > 100 Then n = 100
            If i > 100 Then i = 100
            Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", BrightGreen)
        End If
    Else
        Call PlayerMsg(Index, "Player is not online.", White)
    End If
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Private Sub HandleWarpMeTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
            
    ' The player
    n = FindPlayer(Buffer.ReadString)
    
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
End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Private Sub HandleWarpToMe(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' The player
    n = FindPlayer(Buffer.ReadString)
    
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
End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Private Sub HandleWarpTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' The map
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Call HackingAttempt(Index, "Invalid map")
        Exit Sub
    End If
    
    Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
    Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Private Sub HandleSetSprite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' The sprite
    n = Buffer.ReadLong
    
    Call SetPlayerSprite(Index, n)
    Call SendPlayerData(Index)
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Private Sub HandleGetStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim n As Long

    Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(Index) & " -=-", White)
    Call PlayerMsg(Index, "Level: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), White)
    Call PlayerMsg(Index, "HP: " & GetPlayerVital(Index, Vitals.HP) & "/" & GetPlayerMaxVital(Index, Vitals.HP) & "  MP: " & GetPlayerVital(Index, Vitals.MP) & "/" & GetPlayerMaxVital(Index, Vitals.MP) & "  SP: " & GetPlayerVital(Index, Vitals.SP) & "/" & GetPlayerMaxVital(Index, Vitals.SP), White)
    Call PlayerMsg(Index, "STR: " & GetPlayerStat(Index, Stats.Strength) & "  DEF: " & GetPlayerStat(Index, Stats.Defense) & "  MAGI: " & GetPlayerStat(Index, Stats.Magic) & "  Speed: " & GetPlayerStat(Index, Stats.Speed), White)
    n = (GetPlayerStat(Index, Stats.Strength) \ 2) + (GetPlayerLevel(Index) \ 2)
    i = (GetPlayerStat(Index, Stats.Defense) \ 2) + (GetPlayerLevel(Index) \ 2)
    If n > 100 Then n = 100
    If i > 100 Then i = 100
    Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", White)
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Private Sub HandleRequestNewMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Dir As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    Dir = Buffer.ReadLong
    
    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If
            
    Call PlayerMove(Index, Dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Private Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim MapNum As Long
Dim MapSize As Long
Dim MapData() As Byte
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
                
    MapNum = GetPlayerMap(Index)
    
    i = Map(MapNum).Revision + 1
    
    Call ClearMap(MapNum)
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MapSize = LenB(Map(MapNum))
    ReDim MapData(MapSize - 1)
    MapData = Buffer.ReadBytes(MapSize)
    CopyMemory ByVal VarPtr(Map(MapNum)), ByVal VarPtr(MapData(0)), MapSize
    
    ' set the new revision
    Map(MapNum).Revision = i
    
    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i, MapNum)
    Next
    
    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)
    
    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next
    
    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    
    ' Save the map
    Call SaveMap(MapNum)
    
    Call MapCache_Create(MapNum)
    
    ' Refresh map for everyone online
    For i = 1 To TotalPlayersOnline
        i = PlayersOnline(i)
        If IsPlaying(i) Then
            Call SendMap(i, MapNum)
        End If
    Next
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Private Sub HandleNeedMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim s As Byte
Dim i As Byte
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' Get yes/no value
    s = Buffer.ReadByte
            
    Set Buffer = Nothing
    
    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If
   
    
    For i = 1 To MAX_MAPS
        Call SendMapItemsTo(Index, i)
        Call SendMapNpcsTo(Index, i)
    Next i
    Call SendJoinMap(Index)
    TempPlayer(Index).GettingMap = NO
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SMapDone
    Call SendDataTo(Index, Buffer.ToArray())
            
    Call SendDoorData(Index)
    
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Private Sub HandleMapGetItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(Index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Private Sub HandleMapDropItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim InvNum As Long
Dim Amount As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    InvNum = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Then
        Call HackingAttempt(Index, "Invalid InvNum")
        Exit Sub
    End If
    
    ' Prevent hacking
    If Amount > GetPlayerInvItemValue(Index, InvNum) Then
        Call HackingAttempt(Index, "Item amount modification")
        Exit Sub
    End If
    
    ' Prevent hacking
    If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
        ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
        If Amount <= 0 Then
            'Call HackingAttempt(Index, "Trying to drop 0 amount of currency")
            Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0) ' remove item
            Exit Sub
        End If
    End If
        
    Call PlayerMapDropItem(Index, InvNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Private Sub HandleMapRespawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next
    
    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    
    ' Respawn NpcS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(Index))
    Next
    
    Call PlayerMsg(Index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Private Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim s As String
Dim i As Long
Dim tMapStart As Long
Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    s = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1
    
    For i = 1 To MAX_MAPS
        If LenB(Trim$(Map(i).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else
            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If
            tMapStart = i + 1
            tMapEnd = i + 1
        End If
    Next
    
    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    
    Call PlayerMsg(Index, s, Brown)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Private Sub HandleKickPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' The player index
    n = FindPlayer(Buffer.ReadString)
    
    If n <> Index Then
        If n > 0 Then
       
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(Index) & "!", White)
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If
    Else
        Call PlayerMsg(Index, "You cannot kick yourself!", White)
    End If
End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Private Sub HandleBanList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim F As Long
Dim s As String
Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    n = 1
    F = FreeFile
    Open App.Path & "\data\banlist.txt" For Input As #F
    Do While Not EOF(F)
        Input #F, s
        Input #F, Name
        
        Call PlayerMsg(Index, n & ": Banned IP " & s & " by " & Name, White)
        n = n + 1
    Loop
    Close #F
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Private Sub HandleBanDestroy(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim FileName As String
Dim File As Long
Dim F As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    FileName = App.Path & "\data\banlist.txt"
    
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
    
    Kill FileName

    Call PlayerMsg(Index, "Ban list destroyed.", White)
    
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Private Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' The player index
    n = FindPlayer(Buffer.ReadString)
    
    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call BanIndex(n, Index)
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If
    Else
        Call PlayerMsg(Index, "You cannot ban yourself!", White)
    End If
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Private Sub HandleRequestEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SEditMap
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Private Sub HandleRequestEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SItemEditor
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

' ::::::::::::::::::::::
' :: Edit item packet ::
' ::::::::::::::::::::::
Private Sub HandleEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' The item #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Item Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing item #" & n & ".", ADMIN_LOG)
    Call SendEditItemTo(Index, n)
End Sub

Private Sub HandleDelete(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Editor As Byte
Dim EditorIndex As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    Editor = Buffer.ReadByte
    EditorIndex = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    Select Case Editor
    
        Case EDITOR_ITEM
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then
                Call HackingAttempt(Index, "Invalid Item Index")
                Exit Sub
            End If
            
            Call ClearItem(EditorIndex)
            
            Call SendUpdateItemToAll(EditorIndex)
            Call SaveItem(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted item #" & EditorIndex & ".", ADMIN_LOG)
        
        Case EDITOR_Npc
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then
                Call HackingAttempt(Index, "Invalid Npc Index")
                Exit Sub
            End If
        
            Call ClearNpc(EditorIndex)
        
            Call SendUpdateNpcToAll(EditorIndex)
            Call SaveNpc(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted Npc #" & EditorIndex & ".", ADMIN_LOG)
        
        Case EDITOR_SPELL
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then
                Call HackingAttempt(Index, "Invalid Spell Index")
                Exit Sub
            End If
        
            Call ClearSpell(EditorIndex)
            
            Call SendUpdateSpellToAll(EditorIndex)
            Call SaveSpell(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted spell #" & EditorIndex & ".", ADMIN_LOG)
        
        Case EDITOR_SHOP
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then
                Call HackingAttempt(Index, "Invalid Shop Index")
                Exit Sub
            End If
            
            Call ClearShop(EditorIndex)
            
            Call SendUpdateShopToAll(EditorIndex)
            Call SaveShop(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted shop #" & EditorIndex & ".", ADMIN_LOG)
    End Select
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SREditor
    Call SendDataTo(Index, Buffer.ToArray())

End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Private Sub HandleSaveItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ItemNum As Long
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ItemNum = Buffer.ReadLong
    
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Item Index")
        Exit Sub
    End If
    
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(ItemNum)), ByVal VarPtr(ItemData(0)), ItemSize
    
    ' Save it
    Call SendUpdateItemToAll(ItemNum)
    Call SaveItem(ItemNum)
    Call AddLog(GetPlayerName(Index) & " saved item #" & ItemNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Npc packet ::
' :::::::::::::::::::::::::::::
Private Sub HandleRequestEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SNpcEditor
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

' :::::::::::::::::::::
' :: Edit Npc packet ::
' :::::::::::::::::::::
Private Sub HandleEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' The Npc #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_NPCS Then
        Call HackingAttempt(Index, "Invalid Npc Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing Npc #" & n & ".", ADMIN_LOG)
    Call SendEditNpcTo(Index, n)
End Sub

' :::::::::::::::::::::
' :: Save Npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim NpcNum As Long
Dim Buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    NpcNum = Buffer.ReadLong
    
    ' Prevent hacking
    If NpcNum < 0 Or NpcNum > MAX_NPCS Then
        Call HackingAttempt(Index, "Invalid Npc Index")
        Exit Sub
    End If
    
    NpcSize = LenB(Npc(NpcNum))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(NpcNum)), ByVal VarPtr(NpcData(0)), NpcSize
    
    ' Save it
    Call SendUpdateNpcToAll(NpcNum)
    Call SaveNpc(NpcNum)
    Call AddLog(GetPlayerName(Index) & " saved Npc #" & NpcNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Private Sub HandleRequestEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SShopEditor
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

' ::::::::::::::::::::::
' :: Edit shop packet ::
' ::::::::::::::::::::::
Private Sub HandleEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' The shop #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing shop #" & n & ".", ADMIN_LOG)
    Call SendEditShopTo(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Private Sub HandleSaveShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim ShopNum As Long
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ShopNum = Buffer.ReadLong
    
    ' Prevent hacking
    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    ' Save it
    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)
    Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::::
Private Sub HandleRequestEditSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger SSpellEditor
    Call SendDataTo(Index, Buffer.ToArray())
End Sub

' :::::::::::::::::::::::
' :: Edit spell packet ::
' :::::::::::::::::::::::
Private Sub HandleEditSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' The spell #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_SPELLS Then
        Call HackingAttempt(Index, "Invalid Spell Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing spell #" & n & ".", ADMIN_LOG)
    Call SendEditSpellTo(Index, n)
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Private Sub HandleSaveSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim SpellNum As Long
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' Spell #
    SpellNum = Buffer.ReadLong
    
    ' Prevent hacking
    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Call HackingAttempt(Index, "Invalid Spell Index")
        Exit Sub
    End If
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
            
    ' Save it
    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)
    Call AddLog(GetPlayerName(Index) & " saving spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Private Sub HandleSetAccess(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Call HackingAttempt(Index, "Trying to use powers not available")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' The index
    n = FindPlayer(Buffer.ReadString)
    ' The access
    i = Buffer.ReadLong
    
    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then
        ' Check if player is on
        If n > 0 Then
        
        'check to see if same level access is trying to change another access of the very same level and boot them if they are.
        If GetPlayerAccess(n) = GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "Invalid access level.", Red)
        Exit Sub
    End If
    
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
End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Private Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(Index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Private Sub HandleSetMotd(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    MOTD = Buffer.ReadString
    Call PutVar(App.Path & "\data\motd.ini", "MOTD", "Msg", MOTD)
    Call GlobalMsg("MOTD changed to: " & MOTD, BrightCyan)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & MOTD, ADMIN_LOG)
End Sub

' ::::::::::::::::::
' :: Trade packet ::
' ::::::::::::::::::
Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Map(GetPlayerMap(Index)).Shop > 0 Then
        Call SendTrade(Index, Map(GetPlayerMap(Index)).Shop)
    Else
        Call PlayerMsg(Index, "There is no shop here.", BrightRed)
    End If
End Sub

' ::::::::::::::::::::::::::
' :: Trade request packet ::
' ::::::::::::::::::::::::::
Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim X As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' Trade num
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If (n <= 0) Or (n > MAX_TRADES) Then
        Call HackingAttempt(Index, "Trade Request Modification")
        Exit Sub
    End If
    
    ' Index for shop
    i = Map(GetPlayerMap(Index)).Shop
    
    ' Check if inv full
    X = FindOpenInvSlot(Index, Shop(i).TradeItem(n).GetItem)
    If X = 0 Then
        Call PlayerMsg(Index, "Trade unsuccessful, inventory full.", BrightRed)
        Exit Sub
    End If
    
    ' Check if they have the item
    If HasItem(Index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
        Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
        Call GiveItem(Index, Shop(i).TradeItem(n).GetItem, Shop(i).TradeItem(n).GetValue)
        Call PlayerMsg(Index, "The trade was successful!", Yellow)
    Else
        Call PlayerMsg(Index, "Trade unsuccessful.", BrightRed)
    End If
End Sub

' :::::::::::::::::::::
' :: Fix item packet ::
' :::::::::::::::::::::
Private Sub HandleFixItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim i As Long
Dim ItemNum As Long
Dim DurNeeded As Long
Dim GoldNeeded As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' Inv num
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n <= 0 Or n > MAX_INV Then
        Call HackingAttempt(Index, "Fix item modification")
        Exit Sub
    End If
    
    ' check for bad data
    If GetPlayerInvItemNum(Index, n) <= 0 Or GetPlayerInvItemNum(Index, n) > MAX_ITEMS Then
        Exit Sub
    End If
    
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
    i = (Item(GetPlayerInvItemNum(Index, n)).Data2 \ 5)
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
            GoldNeeded = Int(DurNeeded * i \ 2)
            If GoldNeeded <= 0 Then GoldNeeded = 1
            
            Call TakeItem(Index, 1, GoldNeeded)
            Call SetPlayerInvItemDur(Index, n, GetPlayerInvItemDur(Index, n) + DurNeeded)
            Call PlayerMsg(Index, "Item has been partially fixed for " & GoldNeeded & " gold!", BrightBlue)
        End If
    Else
        Call PlayerMsg(Index, "Insufficient gold to fix this item!", BrightRed)
    End If
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Private Sub HandleSearch(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Long
Dim y As Long
Dim i As Long
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    y = Buffer.ReadLong
    
    ' Prevent subscript out of range
    If X < 0 Or X > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then
        Exit Sub
    End If
    
    ' Check for a player
    For i = 1 To TotalPlayersOnline
        If GetPlayerMap(Index) = GetPlayerMap(PlayersOnline(i)) Then
            If GetPlayerX(PlayersOnline(i)) = X Then
                If GetPlayerY(PlayersOnline(i)) = y Then
        
                    ' Consider the player
                    If PlayersOnline(i) <> Index Then
                    
                        If GetPlayerLevel(PlayersOnline(i)) >= GetPlayerLevel(Index) + 5 Then
                            Call PlayerMsg(Index, "You wouldn't stand a chance.", BrightRed)
                        Else
                            If GetPlayerLevel(PlayersOnline(i)) > GetPlayerLevel(Index) Then
                                Call PlayerMsg(Index, "This one seems to have an advantage over you.", Yellow)
                            Else
                                If GetPlayerLevel(PlayersOnline(i)) = GetPlayerLevel(Index) Then
                                    Call PlayerMsg(Index, "This would be an even fight.", White)
                                Else
                                    If GetPlayerLevel(Index) >= GetPlayerLevel(PlayersOnline(i)) + 5 Then
                                        Call PlayerMsg(Index, "You could slaughter that player.", BrightBlue)
                                    Else
                                        If GetPlayerLevel(Index) > GetPlayerLevel(PlayersOnline(i)) Then
                                            Call PlayerMsg(Index, "You would have an advantage over that player.", Yellow)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                    End If
                
                    ' Change target
                    TempPlayer(Index).Target = PlayersOnline(i)
                    TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER
                    Call PlayerMsg(Index, "Your target is now " & GetPlayerName(PlayersOnline(i)) & ".", Yellow)
                    Exit Sub
                    
                End If
            End If
        End If
                
    Next
    
    ' Check for an item
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(GetPlayerMap(Index), i).Num > 0 Then
            If MapItem(GetPlayerMap(Index), i).X = X Then
                If MapItem(GetPlayerMap(Index), i).y = y Then
                    Call PlayerMsg(Index, "You see a " & Trim$(Item(MapItem(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        End If
    Next
    
    ' Check for an Npc
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(Index), i).Num > 0 Then
            If MapNpc(GetPlayerMap(Index), i).X = X Then
                If MapNpc(GetPlayerMap(Index), i).y = y Then
                    ' Change target
                    TempPlayer(Index).Target = i
                    TempPlayer(Index).TargetType = TARGET_TYPE_Npc
                    Call PlayerMsg(Index, "Your target is now a " & Trim$(Npc(MapNpc(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub

' ::::::::::::::::::
' :: Party packet ::
' ::::::::::::::::::
Private Sub HandleParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    n = FindPlayer(Buffer.ReadString)
    
    ' Prevent partying with self
    If n = Index Then
        Exit Sub
    End If
            
    ' Check for a previous party and if so drop it
    If TempPlayer(Index).InParty = YES Then
        Call PlayerMsg(Index, "You are already in a party!", Pink)
        Exit Sub
    End If
    
    If n > 0 Then
        ' Check if its an admin
        If GetPlayerAccess(Index) > ADMIN_MONITOR Then
            Call PlayerMsg(Index, "You can't join a party, you are an admin!", BrightBlue)
            Exit Sub
        End If
    
        If GetPlayerAccess(n) > ADMIN_MONITOR Then
            Call PlayerMsg(Index, "Admins cannot join parties!", BrightBlue)
            Exit Sub
        End If
        
        ' Make sure they are in right level range
        If GetPlayerLevel(Index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(Index) - 5 > GetPlayerLevel(n) Then
            Call PlayerMsg(Index, "There is more then a 5 level gap between you two, party failed.", Pink)
            Exit Sub
        End If
        
        ' Check to see if player is already in a party
        If TempPlayer(n).InParty = NO Then
            Call PlayerMsg(Index, "Party request has been sent to " & GetPlayerName(n) & ".", Pink)
            Call PlayerMsg(n, GetPlayerName(Index) & " wants you to join their party.  Type /join to join, or /leave to decline.", Pink)
        
            TempPlayer(Index).PartyStarter = YES
            TempPlayer(Index).PartyPlayer = n
            TempPlayer(n).PartyPlayer = Index
        Else
            Call PlayerMsg(Index, "Player is already in a party!", Pink)
        End If
    Else
        Call PlayerMsg(Index, "Player is not online.", White)
    End If
End Sub

' :::::::::::::::::::::::
' :: Join party packet ::
' :::::::::::::::::::::::
Private Sub HandleJoinParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long

    n = TempPlayer(Index).PartyPlayer
    
    If n > 0 Then
        ' Check to make sure they aren't the starter
        If TempPlayer(Index).PartyStarter = NO Then
            ' Check to make sure that each of there party players match
            If TempPlayer(n).PartyPlayer = Index Then
                Call PlayerMsg(Index, "You have joined " & GetPlayerName(n) & "'s party!", Pink)
                Call PlayerMsg(n, GetPlayerName(Index) & " has joined your party!", Pink)
                
                TempPlayer(Index).InParty = YES
                TempPlayer(n).InParty = YES
            Else
                Call PlayerMsg(Index, "Party failed.", Pink)
            End If
        Else
            Call PlayerMsg(Index, "You have not been invited to join a party!", Pink)
        End If
    Else
        Call PlayerMsg(Index, "You have not been invited into a party!", Pink)
    End If
End Sub

' ::::::::::::::::::::::::
' :: Leave party packet ::
' ::::::::::::::::::::::::
Private Sub HandleLeaveParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long

    n = TempPlayer(Index).PartyPlayer
    
    If n > 0 Then
        If TempPlayer(Index).InParty = YES Then
            Call PlayerMsg(Index, "You have left the party.", Pink)
            Call PlayerMsg(n, GetPlayerName(Index) & " has left the party.", Pink)
            
            TempPlayer(Index).PartyPlayer = 0
            TempPlayer(Index).PartyStarter = NO
            TempPlayer(Index).InParty = NO
            TempPlayer(n).PartyPlayer = 0
            TempPlayer(n).PartyStarter = NO
            TempPlayer(n).InParty = NO
        Else
            Call PlayerMsg(Index, "Declined party request.", Pink)
            Call PlayerMsg(n, GetPlayerName(Index) & " declined your request.", Pink)
            
            TempPlayer(Index).PartyPlayer = 0
            TempPlayer(Index).PartyStarter = NO
            TempPlayer(Index).InParty = NO
            TempPlayer(n).PartyPlayer = 0
            TempPlayer(n).PartyStarter = NO
            TempPlayer(n).InParty = NO
        End If
    Else
        Call PlayerMsg(Index, "You are not in a party!", Pink)
    End If
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Private Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(Index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Private Sub HandleCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteBytes Data()
    
    ' Spell slot
    n = Buffer.ReadLong
    
    Call CastSpell(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Private Sub HandleQuit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(Index)
End Sub

' ::::::::::
' :: Sync ::
' ::::::::::

Private Sub HandleSync(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteInteger SSync
    Buffer.WriteByte GetPlayerX(Index)
    Buffer.WriteByte GetPlayerY(Index)
    Buffer.WriteLong GetPlayerMap(Index)

    Call SendDataTo(Index, Buffer.ToArray())
    
End Sub

Private Sub HandleMapReqs(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_MAPS
        If Buffer.ReadByte = 1 Then
            SendMap Index, i
        End If
    Next i
    
End Sub

Private Sub HandleSleepInn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long, n As Long
        For i = 1 To MAX_INV
            Select Case GetPlayerInvItemName(Index, i)
                Case "Gold"
                    If Map(GetPlayerMap(Index)).Moral = MAP_MORAL_INN Then
                        If GetPlayerInvItemValue(Index, i) >= Val(GetPlayerLevel(Index) * 10) Then
                            Call TakeItem(Index, i, Val(GetPlayerLevel(Index) * 10))
                            For n = 1 To Vitals.Vital_Count - 1
                                Call SetPlayerVital(Index, n, GetPlayerMaxVital(Index, n))
                                Call SendVital(Index, n)
                            Next
                            Call PlayerMsg(Index, "You sleep and wake up feeling refreshed!", BrightGreen)
                            Exit Sub
                        ElseIf GetPlayerInvItemValue(Index, i) < Val(GetPlayerLevel(Index) * 10) Then
                            Call PlayerMsg(Index, "You do not have enough money to sleep here!", BrightRed)
                            Exit Sub
                        End If
                    Else
                        Call PlayerMsg(Index, "There is no Inn here. You cannot sleep!", BrightRed)
                        Exit Sub
                    End If
                Case Else
                    Call PlayerMsg(Index, "You do not have any gold with which to pay!", BrightRed)
                    Exit Sub
            End Select
        Next i
End Sub

Private Sub HandleCreateGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer
Dim user As String
Dim Guild As String

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    user = Buffer.ReadString
    Guild = Buffer.ReadString
    
    If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
        For i = 1 To High_Index
            If user = LCase(GetPlayerName(i)) Then
                Call SetPlayerGuild(i, Guild)
                Call SetPlayerGAccess(i, 2)
                Call SendPlayerData(i)
                
                Exit Sub
            End If
        Next
    End If
    
End Sub

Private Sub HandleRemoveFromGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer
Dim user As String
Dim Guild As String

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    user = Buffer.ReadString
    Guild = Buffer.ReadString
    
    If GetPlayerAccess(Index) >= ADMIN_CREATOR Then
        For i = 1 To High_Index
            If user = GetPlayerName(i) Then
                If GetPlayerGuild(i) = Guild Then
                    Call SetPlayerGAccess(i, 0)
                    Call SetPlayerGuild(i, "")
                    Call SendPlayerData(i)
                    Exit Sub
                End If
            End If
        Next
    End If
    
End Sub

Private Sub HandleGuildInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer
Dim user As String
Dim Guild As String

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    user = Buffer.ReadString
    Guild = GetPlayerGuild(Index)
    
    If GetPlayerGAccess(Index) >= 2 Then
        For i = 1 To High_Index
            If user = GetPlayerName(i) Then
                If GetPlayerGuild(i) = "" Then
                    Call SetPlayerGAccess(i, 1)
                    Call SetPlayerGuild(i, Guild)
                    Call SendPlayerData(i)
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub

Private Sub HandleGuildKick(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer
Dim user As String
Dim Guild As String

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    user = Buffer.ReadString
    Guild = GetPlayerGuild(Index)
    
    If GetPlayerGAccess(Index) >= 2 Then
        For i = 1 To High_Index
            If user = GetPlayerName(i) Then
                If GetPlayerGuild(i) = Guild Then
                    Call SetPlayerGAccess(i, 0)
                    Call SetPlayerGuild(i, "")
                    Call SendPlayerData(i)
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub

Private Sub HandleGuildPromote(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer
Dim user As String
Dim Guild As String
Dim Access As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    user = Buffer.ReadString
    Access = Buffer.ReadLong
    Guild = GetPlayerGuild(Index)
    
    If GetPlayerGAccess(Index) >= 2 Then
        If Access > GetPlayerGAccess(Index) Then
            Call PlayerMsg(Index, "You cannot set access higher than your own.", BrightRed)
            Exit Sub
        End If
        For i = 1 To High_Index
            If user = GetPlayerName(i) Then
                If GetPlayerGuild(i) = Guild Then
                    Call SetPlayerGAccess(i, Access)
                    Call SendPlayerData(i)
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub

Private Sub HandleLeaveGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Buffer As clsBuffer
Dim user As String
Dim Guild As String

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    'user = Buffer.ReadString
    Guild = GetPlayerGuild(Index)
    
    If GetPlayerGAccess(Index) >= 1 Then
            Call SetPlayerGAccess(Index, 0)
            Call SetPlayerGuild(Index, "")
            Call SendPlayerData(Index)
    End If
End Sub
