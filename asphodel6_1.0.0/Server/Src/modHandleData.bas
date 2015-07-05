Attribute VB_Name = "modHandleData"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

Public Sub HandleData(ByVal Index As Long, ByVal Data As String)
Dim Parse() As String

    ' Handle incoming data
    Parse = Split(Data, SEP_CHAR)
    
    Select Case Parse(0)
        Case CGetClasses
            HandleGetClasses Index
        Case CNewAccount
            HandleNewAccount Index, Parse
        Case CLogin
            HandleLogin Index, Parse
        Case CAddChar
            HandleAddChar Index, Parse
        Case CDelChar
            HandleDelChar Index, Parse
        Case CUseChar
            HandleUseChar Index, Parse
        Case CMessage
            HandleMessage Index, Parse
        Case CPlayerMove
            HandlePlayerMove Index, Parse
        Case CPlayerDir
            HandlePlayerDir Index, Parse
        Case CUseItem
            HandleUseItem Index, Parse
        Case CAttack
            HandleAttack Index
        Case CUseStatPoint
            HandleUseStatPoint Index, Parse
        Case CPlayerInfoRequest
            HandlePlayerInfoRequest Index, Parse
        Case CWarpMeTo
            HandleWarpMeTo Index, Parse
        Case CWarpToMe
            HandleWarpToMe Index, Parse
        Case CWarpTo
            HandleWarpTo Index, Parse
        Case CSetSprite
            HandleSetSprite Index, Parse
        Case CGetStats
            HandleGetStats Index
        Case CRequestNewMap
            HandleRequestNewMap Index, Parse
        Case CMapData
            HandleMapData Index, Parse
        Case CNeedMap
            HandleNeedMap Index, Parse
        Case CMapGetItem
            HandleMapGetItem Index
        Case CMapDropItem
            HandleMapDropItem Index, Parse
        Case CMapRespawn
            HandleMapRespawn Index
        Case CMapReport
            HandleMapReport Index
        Case CKickPlayer
            HandleKickPlayer Index, Parse
        Case CBanList
            HandleBanList Index
        Case CBanDestroy
            HandleBanDestroy Index
        Case CBanPlayer
            HandleBanPlayer Index, Parse
        Case CRequestEditMap
            HandleRequestEditMap Index
        Case CRequestEditItem
            HandleRequestEditItem Index
        Case CEditItem
            HandleEditItem Index, Parse
        Case CSaveItem
            HandleSaveItem Index, Parse
        Case CDelete
            HandleDelete Index, Parse
        Case CRequestEditNpc
            HandleRequestEditNpc Index
        Case CEditNpc
            HandleEditNpc Index, Parse
        Case CSaveNpc
            HandleSaveNpc Index, Parse
        Case CRequestEditShop
            HandleRequestEditShop Index
        Case CEditShop
            HandleEditShop Index, Parse
        Case CSaveShop
            HandleSaveShop Index, Parse
        Case CRequestEditSpell
            HandleRequestEditSpell Index
        Case CEditSpell
            HandleEditSpell Index, Parse
        Case CSaveSpell
            HandleSaveSpell Index, Parse
        Case CSetAccess
            HandleSetAccess Index, Parse
        Case CWhosOnline
            HandleWhosOnline Index
        Case CSetMotd
            HandleSetMotd Index, Parse
        Case CTradeRequest
            HandleTradeRequest Index, Parse
        Case CFixItem
            HandleFixItem Index, Parse
        Case CSearch
            HandleSearch Index, Parse
        Case CParty
            HandleParty Index, Parse
        Case CJoinParty
            HandleJoinParty Index
        Case CLeaveParty
            HandleLeaveParty Index
        Case CSpells
            HandleSpells Index
        Case CCast
            HandleCast Index, Parse
        Case CQuit
            HandleQuit Index
        Case CConfigPass
            HandleConfigPass Index, Parse
        Case CACPAction
            HandleACPAction Index, Parse
        Case CRCWarp
            HandleRCWarp Index, Parse
        Case CRequestEditSign
            HandleRequestEditSign Index
        Case CSaveSign
            HandleSaveSign Index, Parse
        Case CEditSign
            HandleEditSign Index, Parse
        Case CPressReturn
            HandlePressReturn Index, Parse
        Case CGuildCreation
            HandleGuildCreation Index, Parse
        Case CGuildDisband
            HandleGuildDisband Index
        Case CGuildInvite
            HandleGuildInvite Index, Parse
        Case CInviteResponse
            HandleInviteResponse Index, Parse
        Case CGuildPromoteDemote
            HandleGuildPromoteDemote Index, Parse
        Case CRequestEditAnim
            HandleRequestEditAnim Index
        Case CSaveAnim
            HandleSaveAnim Index, Parse
        Case CEditAnim
            HandleEditAnim Index, Parse
        Case CPing
            HandlePing Index
        Case CLogout
            HandleLogout Index
        Case CSellItem
            HandleSellItem Index, Parse
    End Select
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Requesting classes for making a character ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Private Sub HandleGetClasses(ByVal Index As Long)
    If Not IsPlaying(Index) Then SendNewCharClasses Index
End Sub

' ::::::::::::::::::::::::
' :: New account packet ::
' ::::::::::::::::::::::::
Private Sub HandleNewAccount(ByVal Index As Long, ByRef Parse() As String)
Dim Name As String
Dim Password As String
Dim i As Long
Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
            
            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call NormalMsg(Index, "Your name and password must be at least three characters in length", Window_State.New_Account)
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))
                If Not isNameLegal(n) Then
                    Call NormalMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.", Window_State.New_Account)
                    Exit Sub
                End If
            Next
            
            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                If Not IsBanned(Name, False) Then
                    Call AddAccount(Index, Name, Password)
                    Call TextAdd(frmServer.txtText, "Account " & Name & " has been created.")
                    Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                    Call NormalMsg(Index, "Your account has been created!", Window_State.Main_Menu)
                    ClearPlayer Index
                Else
                    Call NormalMsg(Index, "That account name has been banned from use.", Window_State.Main_Menu)
                    Exit Sub
                End If
            Else
                Call NormalMsg(Index, "Sorry, that account name is already taken!", Window_State.New_Account)
            End If
        End If
    End If
End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal Index As Long, ByRef Parse() As String)
Dim Name As String
Dim Password As String

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Trim$(Parse(1))
            Password = Trim$(Parse(2))
            
            If IsBanned(Name, False) Then
                Call NormalMsg(Index, "You are banned from " & GAME_NAME & "!" & vbNewLine & "Contact an admin at " & GAME_WEBSITE & ".", Window_State.Login)
                Exit Sub
            End If
            
            ' Check versions
            If Val(Parse(3)) <> App.Major Or Val(Parse(4)) <> App.Minor Or Val(Parse(5)) <> App.Revision Then
                Call NormalMsg(Index, "Version outdated, please visit " & GAME_WEBSITE & "!", Window_State.Login)
                Exit Sub
            End If
            
            If Len(Name) < 3 Or Len(Password) < 3 Then
                Call NormalMsg(Index, "Your name and password must be at least three characters in length", Window_State.Login)
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call NormalMsg(Index, "That account name does not exist.", Window_State.Login)
                Exit Sub
            End If
            
            If Not PasswordOK(Name, Password) Then
                Call NormalMsg(Index, "Incorrect password.", Window_State.Login)
                Exit Sub
            End If
            
            If IsMultiAccounts(Name) Then
                Call NormalMsg(Index, "Multiple account logins is not authorized.", Window_State.Login)
                Exit Sub
            End If
            
            ' Load the player
            Call LoadPlayer(Index, Name)
            Call SendChars(Index)
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
        End If
    End If
End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal Index As Long, ByRef Parse() As String)
Dim Name As String
Dim Sex As Long
Dim Class As Long
Dim CharNum As Long
Dim i As Long
Dim n As Long

    If Not IsPlaying(Index) Then
        Name = Parse(1)
        Sex = Val(Parse(2))
        Class = Val(Parse(3))
        CharNum = Val(Parse(4))
        
        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call NormalMsg(Index, "Character name must be at least three characters in length.", Window_State.New_Char)
            Exit Sub
        End If
        
        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))
            
            If Not isNameLegal(n) Then
                Call NormalMsg(Index, "Invalid name! Only letters, numbers, spaces, and _ allowed in names.", Window_State.New_Char)
                Exit Sub
            End If
        Next
        
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid character number.")
            Exit Sub
        End If
        
        ' Prevent hacking
        If (Sex < GenderType.Male_) Or (Sex > GenderType.Female_) Then
            Call HackingAttempt(Index, "Invalid gender.")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Call HackingAttempt(Index, "Invalid class.")
            Exit Sub
        End If
        
        ' Check if char already exists in slot
        If CharExist(Index, CharNum) Then
            Call NormalMsg(Index, "Character already exists!", Window_State.Chars)
            Exit Sub
        End If
        
        ' Check if name is already in use
        If FindChar(Name) Then
            Call NormalMsg(Index, "Sorry, but that name is in use!", Window_State.New_Char)
            Exit Sub
        End If
        
        ' Everything went ok, add the character
        Call AddChar(Index, Name, Sex, Class, CharNum)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        'Call NormalMsg(Index, "Character has been created!", Window_State.Chars)
        SendChars Index
        
    End If
End Sub

' :::::::::::::::::::::::::::::::
' :: Deleting character packet ::
' :::::::::::::::::::::::::::::::
Private Sub HandleDelChar(ByVal Index As Long, ByRef Parse() As String)
Dim CharNum As Long
Dim LoopI As Long
Dim LoopI2 As Long

    If Not IsPlaying(Index) Then
        CharNum = Val(Parse(1))
    
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
        
        For LoopI = 1 To MAX_GUILDS
            If LenB(Trim$(Guild(LoopI).Name)) > 0 Then
                For LoopI2 = 0 To UBound(Guild(LoopI).Member_Account)
                    If GetPlayerLogin(Index) = Trim$(Guild(LoopI).Member_Account(LoopI2)) Then
                        If CharNum = Guild(LoopI).Member_CharNum(LoopI2) Then
                            If LoopI2 = UBound(Guild(LoopI).Member_Account) Then
                                ReDim Preserve Guild(LoopI).Member_Account(0 To UBound(Guild(LoopI).Member_Account) - 1)
                                ReDim Preserve Guild(LoopI).Member_CharNum(0 To UBound(Guild(LoopI).Member_CharNum) - 1)
                            Else
                                Guild(LoopI).Member_Account(LoopI2) = vbNullString
                                Guild(LoopI).Member_CharNum(LoopI2) = 0
                            End If
                        End If
                    End If
                Next
            End If
        Next
        
        Call DelChar(Index, CharNum)
        Call AddLog("Character deleted on " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        'Call NormalMsg(Index, "Character has been deleted!", Window_State.Chars)
        
        SendChars Index
    End If
    
End Sub

' ::::::::::::::::::::::::::::
' :: Using character packet ::
' ::::::::::::::::::::::::::::
Private Sub HandleUseChar(ByVal Index As Long, ByRef Parse() As String)
Dim CharNum As Long
Dim F As Long

    If Not IsPlaying(Index) Then
        CharNum = Val(Parse(1))
        
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
        
        ' Check to make sure the character exists and if so, set it as its current char
        If CharExist(Index, CharNum) Then
            If frmServer.chkStaffOnly.Value = 1 Then
                If Player(Index).Char(CharNum).Access < 1 Then
                    NormalMsg Index, "Sorry, only staff are allowed to enter!", Window_State.Chars
                    Exit Sub
                End If
            End If
            
            TempPlayer(Index).CharNum = CharNum
            Call JoinGame(Index)
            
            CharNum = TempPlayer(Index).CharNum
            Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
            Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".")
            Call UpdateCaption
            
            ' we'll send an update of the player's misc stuff for their stat window
            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index))
            Call SetPlayerLevel(Index, GetPlayerLevel(Index))
            Call SetPlayerExp(Index, GetPlayerExp(Index))
            
            SendDataTo Index, SClassName & SEP_CHAR & Trim$(Class(GetPlayerClass(Index)).Name) & END_CHAR
            
            CheckIfGuildStillExists Index
            SendPlayerGuildToAll Index
            
            For F = 1 To MAX_PLAYERS
                If IsPlaying(F) Then
                    SendPlayerGuildTo Index, F
                End If
            Next
            
            For F = 1 To MAX_INV
                If GetPlayerInvItemNum(Index, F) > 0 Then SendUpdateItemTo Index, GetPlayerInvItemNum(Index, F)
            Next
            
            ' Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
            If Not FindChar(GetPlayerName(Index)) Then
                F = FreeFile
                Open App.Path & "\accounts\charlist.txt" For Append As #F
                    Print #F, GetPlayerName(Index)
                Close #F
            End If
        Else
            Call NormalMsg(Index, "Character does not exist!", Window_State.Chars)
        End If
    End If
    
End Sub

' ::::::::::::::::::::
' :: Social packet  ::
' ::::::::::::::::::::

Private Sub HandleMessage(ByVal Index As Long, ByRef Parse() As String)
Dim Message As String
Dim ChatTag As String
Dim ChatType As Byte
Dim ChatColor As Long
Dim SendTo As Long
Dim SendToName As String
Dim i As Long

    If Player(Index).Char(TempPlayer(Index).CharNum).Muted Then
        PlayerMsg Index, "You are muted, you cannot talk! (" & CInt((Player(Index).Char(TempPlayer(Index).CharNum).MuteTime - GetTickCountNew) / 60000) & " minutes left)", Color.BrightRed
        Exit Sub
    End If
    
    ChatType = CByte(Parse(1))
    Message = Parse(2)
    
    ' Prevent hacking
    For i = 1 To Len(Message)
        If AscW(Mid$(Message, i, 1)) < 32 Or AscW(Mid$(Message, i, 1)) > 126 Then
            Call HackingAttempt(Index, "Chat message modification.")
            Exit Sub
        End If
    Next
    
    Select Case ChatType
        Case E_ChatType.MapMsg_
            If frmServer.chkMapChat.Value <> 1 Then
                PlayerMsg Index, "Map chat has been disabled by the server.", Color.BrightRed
                Exit Sub
            End If
            ChatTag = "[MAP]"
            ChatColor = SayColor
        Case E_ChatType.EmoteMsg_
            If frmServer.chkEmoteChat.Value <> 1 Then
                PlayerMsg Index, "Emote chat has been disabled by the server.", Color.BrightRed
                Exit Sub
            End If
            ChatColor = EmoteColor
            Message = GetPlayerName(Index) & " " & Right$(Message, Len(Message) - 1)
            MapMsg GetPlayerMap(Index), Message, ChatColor
            AddLog "[EMOTE] " & Message, PLAYER_LOG
            TextAdd frmServer.txtText, "[EMOTE] " & Message
            Exit Sub
        Case E_ChatType.BroadcastMsg_
            If frmServer.chkGlobalChat.Value <> 1 Then
                PlayerMsg Index, "Global chat has been disabled by the server.", Color.BrightRed
                Exit Sub
            End If
            ChatColor = BroadcastColor
            ChatTag = "[GLOBAL]"
        Case E_ChatType.GlobalMsg_
            ChatColor = GlobalColor
            ChatTag = "[ALERT]"
        Case E_ChatType.AdminMsg_
            ChatColor = AdminColor
            ChatTag = "[ADMIN]"
        Case E_ChatType.PrivateMsg_
            If frmServer.chkPrivateChat.Value <> 1 Then
                PlayerMsg Index, "Private chat has been disabled by the server.", Color.BrightRed
                Exit Sub
            End If
            ChatColor = TellColor
            ChatTag = "[PM]"
            For i = 1 To Len(Message)
                If Mid$(Message, i, 1) <> " " Then
                    SendToName = SendToName & Mid$(Message, i, 1)
                Else
                    Exit For
                End If
            Next
            Message = Mid$(Message, i, Len(Message) - 1)
            If LenB(Message) > 0 Then
                If FindPlayer(SendToName) Then
                    If FindPlayer(SendToName) <> Index Then
                        SendMessage Index, ChatType, ChatTag, Message, ChatColor, SendTo
                    Else
                        PlayerMsg Index, "You cannot send a message to yourself.", BrightRed
                    End If
                Else
                    PlayerMsg Index, "Player is not online.", BrightRed
                End If
            Else
                PlayerMsg Index, "You need to enter a message!", BrightRed
            End If
            Exit Sub
    End Select
    
    AddLog ChatTag & " " & GetPlayerName(Index) & ":" & Message, PLAYER_LOG
    TextAdd frmServer.txtText, ChatTag & " " & GetPlayerName(Index) & ": " & Message
    SendMessage Index, ChatType, ChatTag, Message, ChatColor
    
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Parse() As String)
Dim Dir As Long
Dim Movement As Long

    If TempPlayer(Index).GettingMap = YES Then Exit Sub

    Dir = Val(Parse(1))
    Movement = Val(Parse(2))
    
    ' Prevent hacking
    If Dir < E_Direction.Up_ Or Dir > E_Direction.Right_ Then
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
        If GetTickCountNew > TempPlayer(Index).AttackTimer + 1000 Then
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
Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Parse() As String)
Dim Dir As Long

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Val(Parse(1))
    
    ' Prevent hacking
    If Dir < E_Direction.Up_ Or Dir > E_Direction.Right_ Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If
    
    Call SetPlayerDir(Index, Dir)
    Call SendDataToMapBut(Index, GetPlayerMap(Index), SPlayerDir & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)

End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Private Sub HandleUseItem(ByVal Index As Long, ByRef Parse() As String)
Dim InvNum As Long
Dim CharNum As Long
Dim i As Long
Dim n As Long
Dim X As Long
Dim Y As Long
Dim UseAnim As Boolean
Dim ItemNum As Long

    InvNum = Val(Parse(1))
    CharNum = TempPlayer(Index).CharNum
    
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
    
    If (GetPlayerInvItemNum(Index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
            
            ItemNum = GetPlayerInvItemNum(Index, InvNum)
            
            ' Find out what kind of item it is
            Select Case Item(ItemNum).Type
                Case ItemType.Armor_
                    If InvNum <> GetPlayerEquipmentSlot(Index, Armor) Then
                        If Not Meets_ItemRequired(Index, InvNum) Then
                            Call PlayerMsg(Index, "You don't have the required stats to wear this!", BrightRed)
                            Exit Sub
                        End If
                        If GetPlayerEquipmentSlot(Index, Armor) > 0 Then
                            For i = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffStats(i), False
                            Next
                            For i = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffVitals(i), False
                            Next
                            Call SetPlayerEquipmentSlot(Index, 0, Armor)
                        End If
                        Call SetPlayerEquipmentSlot(Index, InvNum, Armor)
                        For i = 1 To Stats.Stat_Count - 1
                            AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffStats(i)
                        Next
                        For i = 1 To Vitals.Vital_Count - 1
                            AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffVitals(i)
                        Next
                        PlayerMsg Index, "You have equipped a " & Trim$(Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).Name) & "!", Color.BrightBlue
                    Else
                        PlayerMsg Index, "You have unequipped a " & Trim$(Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).Name) & "!", Color.BrightBlue
                        For i = 1 To Stats.Stat_Count - 1
                            AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffStats(i), False
                        Next
                        For i = 1 To Vitals.Vital_Count - 1
                            AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Armor))).BuffVitals(i), False
                        Next
                        Call SetPlayerEquipmentSlot(Index, 0, Armor)
                    End If
                    Call SendWornEquipment(Index)
                    
                Case ItemType.Weapon_
                    If InvNum <> GetPlayerEquipmentSlot(Index, Weapon) Then
                        If Not Meets_ItemRequired(Index, InvNum) Then
                            Call PlayerMsg(Index, "You don't have the required stats to wear this!", BrightRed)
                            Exit Sub
                        End If
                        If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
                            For i = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffStats(i), False
                            Next
                            For i = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffVitals(i), False
                            Next
                            Call SetPlayerEquipmentSlot(Index, 0, Weapon)
                        End If
                        Call SetPlayerEquipmentSlot(Index, InvNum, Weapon)
                        For i = 1 To Stats.Stat_Count - 1
                            AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffStats(i)
                        Next
                        For i = 1 To Vitals.Vital_Count - 1
                            AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffVitals(i)
                        Next
                        PlayerMsg Index, "You have equipped a " & Trim$(Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).Name) & "!", Color.BrightBlue
                    Else
                        PlayerMsg Index, "You have unequipped a " & Trim$(Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).Name) & "!", Color.BrightBlue
                        For i = 1 To Stats.Stat_Count - 1
                            AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffStats(i), False
                        Next
                        For i = 1 To Vitals.Vital_Count - 1
                            AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).BuffVitals(i), False
                        Next
                        Call SetPlayerEquipmentSlot(Index, 0, Weapon)
                    End If
                    Call SendWornEquipment(Index)
                    
                Case ItemType.Helmet_
                    If InvNum <> GetPlayerEquipmentSlot(Index, Helmet) Then
                        If Not Meets_ItemRequired(Index, InvNum) Then
                            Call PlayerMsg(Index, "You don't have the required stats to wear this!", BrightRed)
                            Exit Sub
                        End If
                        If GetPlayerEquipmentSlot(Index, Helmet) > 0 Then
                            For i = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffStats(i), False
                            Next
                            For i = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffVitals(i), False
                            Next
                            Call SetPlayerEquipmentSlot(Index, 0, Helmet)
                        End If
                        Call SetPlayerEquipmentSlot(Index, InvNum, Helmet)
                        For i = 1 To Stats.Stat_Count - 1
                            AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffStats(i)
                        Next
                        For i = 1 To Vitals.Vital_Count - 1
                            AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffVitals(i)
                        Next
                        PlayerMsg Index, "You have equipped a " & Trim$(Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).Name) & "!", Color.BrightBlue
                    Else
                        PlayerMsg Index, "You have unequipped a " & Trim$(Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).Name) & "!", Color.BrightBlue
                        For i = 1 To Stats.Stat_Count - 1
                            AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffStats(i), False
                        Next
                        For i = 1 To Vitals.Vital_Count - 1
                            AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Helmet))).BuffVitals(i), False
                        Next
                        Call SetPlayerEquipmentSlot(Index, 0, Helmet)
                    End If
                    Call SendWornEquipment(Index)
                    
                Case ItemType.Shield_
                    If InvNum <> GetPlayerEquipmentSlot(Index, Shield) Then
                        If Not Meets_ItemRequired(Index, InvNum) Then
                            Call PlayerMsg(Index, "You don't have the required stats to wear this!", BrightRed)
                            Exit Sub
                        End If
                        If GetPlayerEquipmentSlot(Index, Shield) > 0 Then
                            For i = 1 To Stats.Stat_Count - 1
                                AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffStats(i), False
                            Next
                            For i = 1 To Vitals.Vital_Count - 1
                                AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffVitals(i), False
                            Next
                            Call SetPlayerEquipmentSlot(Index, 0, Shield)
                        End If
                        Call SetPlayerEquipmentSlot(Index, InvNum, Shield)
                        For i = 1 To Stats.Stat_Count - 1
                            AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffStats(i)
                        Next
                        For i = 1 To Vitals.Vital_Count - 1
                            AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffVitals(i)
                        Next
                        PlayerMsg Index, "You have equipped a " & Trim$(Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).Name) & "!", Color.BrightBlue
                    Else
                        For i = 1 To Stats.Stat_Count - 1
                            AdjustStatBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffStats(i), False
                        Next
                        For i = 1 To Vitals.Vital_Count - 1
                            AdjustVitalBonus Index, i, Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).BuffVitals(i), False
                        Next
                        PlayerMsg Index, "You have unequipped a " & Trim$(Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Shield))).Name) & "!", Color.BrightBlue
                        Call SetPlayerEquipmentSlot(Index, 0, Shield)
                    End If
                    Call SendWornEquipment(Index)
                    
                Case ItemType.Potion
                    If Not Meets_ItemRequired(Index, InvNum) Then
                        Call PlayerMsg(Index, "You don't have the required stats to use this!", BrightRed)
                        Exit Sub
                    End If
                    
                    If Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 <> 0 Then
                        Call SetPlayerVital(Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                        If Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 > 0 Then
                            Call PlayerMsg(Index, "You have recovered " & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & " HP!", Color.Green)
                        Else
                            DirectDamagePlayer Index, Abs(Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1), "You have lost " & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & " HP!"
                        End If
                    End If
                    
                    If Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data2 <> 0 Then
                        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data2)
                        If Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data2 > 0 Then
                            Call PlayerMsg(Index, "You have recovered " & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data2 & " MP!", Color.Green)
                        Else
                            Call PlayerMsg(Index, "You have lost " & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data2 & " MP!", Color.BrightRed)
                        End If
                    End If
                    
                    If Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data3 <> 0 Then
                        Call SetPlayerVital(Index, Vitals.SP, GetPlayerVital(Index, Vitals.SP) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data3)
                        If Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data3 > 0 Then
                            Call PlayerMsg(Index, "You have recovered " & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data3 & " SP!", Color.Green)
                        Else
                            Call PlayerMsg(Index, "You have lost " & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data3 & " SP!", Color.BrightRed)
                        End If
                    End If
                    
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    UseAnim = True
                    
                Case ItemType.Key
                    If Not Meets_ItemRequired(Index, InvNum) Then
                        Call PlayerMsg(Index, "You don't have the required stats to use this!", BrightRed)
                        Exit Sub
                    End If
                    Select Case GetPlayerDir(Index)
                        Case E_Direction.Up_
                            If GetPlayerY(Index) > 0 Then
                                X = GetPlayerX(Index)
                                Y = GetPlayerY(Index) - 1
                            Else
                                Exit Sub
                            End If
                            
                        Case E_Direction.Down_
                            If GetPlayerY(Index) < MAX_MAPY Then
                                X = GetPlayerX(Index)
                                Y = GetPlayerY(Index) + 1
                            Else
                                Exit Sub
                            End If
                                
                        Case E_Direction.Left_
                            If GetPlayerX(Index) > 0 Then
                                X = GetPlayerX(Index) - 1
                                Y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                                
                        Case E_Direction.Right_
                            If GetPlayerX(Index) < MAX_MAPX Then
                                X = GetPlayerX(Index) + 1
                                Y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                    End Select
                    
                    ' Check if a key exists
                    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = Tile_Type.Key_ Then
                        ' Check if the key they are using matches the map key
                        If ItemNum = Map(GetPlayerMap(Index)).Tile(X, Y).Data1 Then
                            TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCountNew + 5000
                            
                            Call SendDataToMap(GetPlayerMap(Index), SMapKey & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                            Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", Color.White)
                            UseAnim = True
                            
                            ' Check if we are supposed to take away the item
                            If Map(GetPlayerMap(Index)).Tile(X, Y).Data2 = 1 Then
                                Call TakeItem(Index, ItemNum, 0)
                                Call PlayerMsg(Index, "The key disolves.", Color.Yellow)
                            End If
                        End If
                    End If
                    
                Case ItemType.Spell_
                
                    If Not Meets_ItemRequired(Index, InvNum) Then
                        Call PlayerMsg(Index, "You don't have the required stats to learn this spell!", BrightRed)
                        Exit Sub
                    End If
                    
                    ' Get the spell num
                    n = Item(ItemNum).Data1
                    
                    If n > 0 Then
                        ' Make sure they are the right class
                        'If Spell(n).ClassReq = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                            ' Make sure they are the right level
                            'I = GetSpellReqLevel(n)
                            'If I <= GetPlayerLevel(Index) Then
                                i = FindOpenSpellSlot(Index)
                                
                                ' Make sure they have an open spell slot
                                If i > 0 Then
                                    ' Make sure they dont already have the spell
                                    If Not HasSpell(Index, n) Then
                                        Call SetPlayerSpell(Index, i, n)
                                        Call TakeItem(Index, ItemNum, 0)
                                        Call PlayerMsg(Index, "You study the spell carefully...", Color.Yellow)
                                        Call PlayerMsg(Index, "You have learned a new spell!", Color.White)
                                        Call SendPlayerSpells(Index)
                                        UseAnim = True
                                    Else
                                        Call PlayerMsg(Index, "You have already learned this spell!", BrightRed)
                                    End If
                                Else
                                    Call PlayerMsg(Index, "You have learned all that you can learn!", BrightRed)
                                End If
                            'Else
                            '    Call PlayerMsg(Index, "You must be level " & I & " to learn this spell.", Color.White)
                            'End If
                        'Else
                        '    Call PlayerMsg(Index, "This spell can only be learned by a " & GetClassName(Spell(n).ClassReq) & ".", Color.White)
                        'End If
                    Else
                        Call PlayerMsg(Index, "This scroll is not connected to a spell, please inform an admin!", Color.White)
                    End If
                    
            End Select
        End If
    End If
    
    If UseAnim Then
        If Item(ItemNum).Anim > 0 Then
            SendDataToMap GetPlayerMap(Index), SAnimation & SEP_CHAR & Item(ItemNum).Anim & SEP_CHAR & Index & SEP_CHAR & E_Target.Player_ & END_CHAR
        End If
    End If
    
    SendVital Index, HP
    SendVital Index, MP
    SendVital Index, SP
    
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAttack(ByVal Index As Long)
Dim i As Long
Dim n As Long
Dim Damage As Long
Dim TempIndex As Long

    ' Try to attack a player
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            
            TempIndex = i
            
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
                            Damage = n + Random(1, n * 0.5) - GetPlayerProtection(TempIndex)
                            Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                            Call PlayerMsg(TempIndex, GetPlayerName(Index) & " swings with enormous might!", BrightCyan)
                        End If
                        
                        If Damage > 0 Then
                            If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
                                If Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).Anim > 0 Then
                                    SendDataToMap GetPlayerMap(Index), SAnimation & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).Anim & SEP_CHAR & TempIndex & SEP_CHAR & E_Target.Player_ & END_CHAR
                                End If
                            End If
                            Call AttackPlayer(Index, TempIndex, Damage)
                        Else
                            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
                        End If
                    Else
                        Call PlayerMsg(Index, GetPlayerName(TempIndex) & "'s " & Trim$(Item(GetPlayerInvItemNum(TempIndex, GetPlayerEquipmentSlot(TempIndex, Shield))).Name) & " has blocked your hit!", BrightCyan)
                        Call PlayerMsg(TempIndex, "Your " & Trim$(Item(GetPlayerInvItemNum(TempIndex, GetPlayerEquipmentSlot(TempIndex, Shield))).Name) & " has blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan)
                    End If
                    
                    Exit Sub
                End If
            End If
        End If
    Next
    
    ' Try to attack a npc
    For i = 1 To UBound(MapSpawn(GetPlayerMap(Index)).Npc)
        ' Can we attack the npc?
        If CanAttackNpc(Index, i) Then
            ' Get the damage we can do
            If Not CanPlayerCriticalHit(Index) Then
                Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index)).MapNpc(i).Num).Stat(Stats.Defense) * 0.5)
            Else
                n = GetPlayerDamage(Index)
                Damage = n + Random(1, n / 2) - Int(Npc(MapNpc(GetPlayerMap(Index)).MapNpc(i).Num).Stat(Stats.Defense) * 0.5)
                Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
            End If
            
            If Damage > 0 Then
                If GetPlayerEquipmentSlot(Index, Weapon) > 0 Then
                    If Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).Anim > 0 Then
                        SendDataToMap GetPlayerMap(Index), SAnimation & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, Weapon))).Anim & SEP_CHAR & i & SEP_CHAR & E_Target.NPC_ & END_CHAR
                    End If
                End If
                Call AttackNpc(Index, i, Damage)
                ' if they didn't kill the NPC, then do the reflection
                If MapNpc(GetPlayerMap(Index)).MapNpc(i).Num > 0 Then
                    If Npc(MapNpc(GetPlayerMap(Index)).MapNpc(i).Num).Reflection(NPC_Reflection.Magic_) > 0 Then
                        Damage = Damage * (Npc(MapNpc(GetPlayerMap(Index)).MapNpc(i).Num).Reflection(NPC_Reflection.Melee_) * 0.01)
                        NpcAttackPlayer i, Index, Damage, True
                    End If
                End If
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
Private Sub HandleUseStatPoint(ByVal Index As Long, ByRef Parse() As String)
Dim PointType As Long

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
                Call SetPlayerStat(Index, Stats.Strength, GetPlayerStat(Index, Stats.Strength) + 1)
                Call PlayerMsg(Index, "You have gained more Strength!", Color.White)
            Case 1
                Call SetPlayerStat(Index, Stats.Defense, GetPlayerStat(Index, Stats.Defense) + 1)
                Call PlayerMsg(Index, "You have gained more Defense!", Color.White)
            Case 2
                Call SetPlayerStat(Index, Stats.Magic, GetPlayerStat(Index, Stats.Magic) + 1)
                Call PlayerMsg(Index, "You have gained more Magic!", Color.White)
            Case 3
                Call SetPlayerStat(Index, Stats.SPEED, GetPlayerStat(Index, Stats.SPEED) + 1)
                Call PlayerMsg(Index, "You have gained more Speed!", Color.White)
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
Private Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Parse() As String)
Dim Name As String
Dim i As Long
Dim n As Long

    Name = Parse(1)
    
    i = FindPlayer(Name)
    If i > 0 Then
        Call PlayerMsg(Index, "Account: " & Trim$(Player(i).Login) & ", Name: " & GetPlayerName(i), BrightGreen)
        If GetPlayerAccess(Index) > StaffType.Monitor Then
            Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
            Call PlayerMsg(Index, "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), BrightGreen)
            Call PlayerMsg(Index, "HP: " & GetPlayerVital(i, Vitals.HP) & "/" & GetPlayerMaxVital(i, Vitals.HP) & "  MP: " & GetPlayerVital(i, Vitals.MP) & "/" & GetPlayerMaxVital(i, Vitals.MP) & "  SP: " & GetPlayerVital(i, Vitals.SP) & "/" & GetPlayerMaxVital(i, Vitals.SP), BrightGreen)
            Call PlayerMsg(Index, "Strength: (" & GetPlayerStat_withBonus(i, Stats.Strength) & "/" & GetPlayerStat(i, Strength) & ")  Defense: (" & GetPlayerStat_withBonus(i, Defense) & "/" & GetPlayerStat(i, Stats.Defense) & ")  Magic: (" & GetPlayerStat_withBonus(i, Stats.Magic) & "/" & GetPlayerStat(i, Stats.Magic) & ")  Speed: (" & GetPlayerStat_withBonus(i, Stats.SPEED) & "/" & GetPlayerStat(i, Stats.SPEED) & ")", BrightGreen)
            n = Int(GetPlayerStat_withBonus(i, Stats.Strength) * 0.5) + Int(GetPlayerLevel(i) * 0.5)
            i = Int(GetPlayerStat_withBonus(i, Stats.Defense) * 0.5) + Int(GetPlayerLevel(i) * 0.5)
            If n > 100 Then n = 100
            If i > 100 Then i = 100
            Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", BrightGreen)
        End If
    Else
        Call PlayerMsg(Index, "Player is not online.", Color.White)
    End If
    
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Private Sub HandleWarpMeTo(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Mapper Then
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
            Call PlayerMsg(Index, "Player is not online.", Color.White)
        End If
    Else
        Call PlayerMsg(Index, "You cannot warp to yourself!", Color.White)
    End If
    
End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Private Sub HandleWarpToMe(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Mapper Then
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
            Call PlayerMsg(Index, "Player is not online.", Color.White)
        End If
    Else
        Call PlayerMsg(Index, "You cannot warp yourself to yourself!", Color.White)
    End If
    
End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Private Sub HandleWarpTo(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Mapper Then
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
    
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Private Sub HandleSetSprite(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Mapper Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The sprite
    n = Val(Parse(1))
    
    Call SetPlayerSprite(Index, n)
    Call SendPlayerData(Index)
    Exit Sub
    
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Private Sub HandleGetStats(ByVal Index As Long)
Dim i As Long
Dim n As Long

    Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(Index) & " -=-", Color.White)
    Call PlayerMsg(Index, "Level: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), Color.White)
    Call PlayerMsg(Index, "HP: " & GetPlayerVital(Index, Vitals.HP) & "/" & GetPlayerMaxVital(Index, Vitals.HP) & "  MP: " & GetPlayerVital(Index, Vitals.MP) & "/" & GetPlayerMaxVital(Index, Vitals.MP) & "  SP: " & GetPlayerVital(Index, Vitals.SP) & "/" & GetPlayerMaxVital(Index, Vitals.SP), Color.White)
    Call PlayerMsg(Index, "STR: (" & GetPlayerStat_withBonus(Index, Stats.Strength) & "/" & GetPlayerStat(Index, Strength) & ")  DEF: (" & GetPlayerStat_withBonus(Index, Defense) & "/" & GetPlayerStat(Index, Stats.Defense) & ")  MAGI: (" & GetPlayerStat_withBonus(Index, Stats.Magic) & "/" & GetPlayerStat(Index, Stats.Magic) & ")  Speed: (" & GetPlayerStat_withBonus(Index, Stats.SPEED) & "/" & GetPlayerStat(Index, Stats.SPEED) & ")", Color.White)
    
    n = Int(GetPlayerStat_withBonus(Index, Stats.Strength) * 0.5) + Int(GetPlayerLevel(Index) * 0.5)
    i = Int(GetPlayerStat_withBonus(Index, Stats.Defense) * 0.5) + Int(GetPlayerLevel(Index) * 0.5)
    
    If n > 100 Then n = 100
    If i > 100 Then i = 100
    
    Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", Color.White)
    
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Private Sub HandleRequestNewMap(ByVal Index As Long, ByRef Parse() As String)
Dim Dir As Long

    Dir = Val(Parse(1))
    
    ' Prevent hacking
    If Dir < E_Direction.Up_ Or Dir > E_Direction.Right_ Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If
            
    Call PlayerMove(Index, Dir, 1)
    
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Private Sub HandleMapData(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long
Dim i As Long
Dim MapNum As Long
Dim X As Long
Dim Y As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Mapper Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    n = 0
    
    MapNum = GetPlayerMap(Index)
    
    i = Map(MapNum).Revision + 1
    
    Call ClearMap(MapNum)
    
    MapNum = GetPlayerMap(Index)
    Map(MapNum).Name = Parse(n + 1)
    Map(MapNum).Revision = i
    Map(MapNum).Moral = Val(Parse(n + 2))
    Map(MapNum).Up = Val(Parse(n + 3))
    Map(MapNum).Down = Val(Parse(n + 4))
    Map(MapNum).Left = Val(Parse(n + 5))
    Map(MapNum).Right = Val(Parse(n + 6))
    Map(MapNum).Music = Parse(n + 7)
    Map(MapNum).BootMap = Val(Parse(n + 8))
    Map(MapNum).BootX = Val(Parse(n + 9))
    Map(MapNum).BootY = Val(Parse(n + 10))
    
    n = n + 11
    
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
        
            For i = 0 To UBound(Map(MapNum).Tile(X, Y).Layer)
                Map(MapNum).Tile(X, Y).Layer(i) = Val(Parse(n))
                Map(MapNum).Tile(X, Y).LayerSet(i) = Val(Parse(n + 1))
                n = n + 2
            Next
            
            Map(MapNum).Tile(X, Y).Type = Val(Parse(n))
            Map(MapNum).Tile(X, Y).Data1 = Val(Parse(n + 1))
            Map(MapNum).Tile(X, Y).Data2 = Val(Parse(n + 2))
            Map(MapNum).Tile(X, Y).Data3 = Val(Parse(n + 3))
            
            n = n + 4
            
        Next
    Next
    
    ReDim MapSpawn(MapNum).Npc(1 To Val(Parse(n)))
    ReDim MapNpc(MapNum).MapNpc(1 To Val(Parse(n)))
    
    n = n + 1
    
    For X = 1 To UBound(MapSpawn(MapNum).Npc)
        MapSpawn(MapNum).Npc(X).Num = Val(Parse(n))
        MapSpawn(MapNum).Npc(X).X = Val(Parse(n + 1))
        MapSpawn(MapNum).Npc(X).Y = Val(Parse(n + 2))
        n = n + 3
        Call ClearMapNpc(X, MapNum)
    Next
    
    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).X, MapItem(GetPlayerMap(Index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next
    
    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    
    ' Save the map
    Call SaveMap(MapNum)
    
    Call MapCache_Create(MapNum)
    
    ' Refresh map for everyone online
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
            End If
        End If
    Next
    
    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)
    
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Private Sub HandleNeedMap(ByVal Index As Long, ByRef Parse() As String)

    ' Check if map data is needed to be sent
    If Parse(1) = 1 Then Call SendMap(Index, GetPlayerMap(Index))
    
    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)
    
    TempPlayer(Index).GettingMap = NO
    
    Call SendDataTo(Index, SMapDone & END_CHAR)
    
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Private Sub HandleMapGetItem(ByVal Index As Long)
    Call PlayerMapGetItem(Index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Private Sub HandleMapDropItem(ByVal Index As Long, ByRef Parse() As String)
Dim InvNum As Long
Dim Ammount As Long

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
    If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ItemType.Currency_ Then
        ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
        If Ammount <= 0 Then
            Call HackingAttempt(Index, "Invalid drop value.")
            Exit Sub
        End If
    End If
    
    Call PlayerMapDropItem(Index, InvNum, Ammount)
    
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Private Sub HandleMapRespawn(ByVal Index As Long)
Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Mapper Then
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
    For i = 1 To UBound(MapSpawn(GetPlayerMap(Index)).Npc)
        Call SpawnNpc(i, GetPlayerMap(Index))
    Next
    
    Call PlayerMsg(Index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
    
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Private Sub HandleMapReport(ByVal Index As Long)
Dim s As String
Dim i As Long
Dim tMapStart As Long
Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Mapper Then
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
Private Sub HandleKickPlayer(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The player index
    n = FindPlayer(Parse(1))
    
    If n <> Index Then
        If n > 0 Then
       
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(Index) & "!", Color.White)
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", Color.White)
            End If
        Else
            Call PlayerMsg(Index, "Player is not online.", Color.White)
        End If
    Else
        Call PlayerMsg(Index, "You cannot kick yourself!", Color.White)
    End If
    
End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Private Sub HandleBanList(ByVal Index As Long)
Dim n As Long
Dim F As Long
Dim s As String
Dim FileName As String

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Mapper Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    FileName = App.Path & "\data\bans.ini"
    
    If Val(GetVar(FileName, "IP", "Total")) = 0 Then
        PlayerMsg Index, "No IPs in ban list.", Color.BrightRed
    Else
        For n = 1 To Val(GetVar(FileName, "IP", "Total"))
            PlayerMsg Index, "Ban IP " & n & ": " & GetVar(FileName, "IP", "IP" & n), Color.Black
        Next
    End If
    
    If Val(GetVar(FileName, "ACCOUNT", "Total")) = 0 Then
        PlayerMsg Index, "No Accounts in ban list.", Color.BrightRed
    Else
        For n = 1 To Val(GetVar(FileName, "ACCOUNT", "Total"))
            PlayerMsg Index, "Ban Account " & n & ": " & GetVar(FileName, "ACCOUNT", "Account" & n), Color.Black
        Next
    End If
    
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Private Sub HandleBanDestroy(ByVal Index As Long)
Dim FileName As String
Dim F As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Creator Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    FileName = App.Path & "\data\bans.ini"
    
    If Not FileExist("data\bans.ini") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    Else
        Kill FileName
    End If
    
    Call PlayerMsg(Index, "Ban list destroyed.", Color.White)
    
    Load_BanTable
    
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Private Sub HandleBanPlayer(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Mapper Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The player index
    n = FindPlayer(Parse(1))
    
    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call BanIndex(n, Index)
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", Color.White)
            End If
        Else
            Call PlayerMsg(Index, "Player is not online.", Color.White)
        End If
    Else
        Call PlayerMsg(Index, "You cannot ban yourself!", Color.White)
    End If
    
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Private Sub HandleRequestEditMap(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Mapper Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SEditMap & END_CHAR)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Private Sub HandleRequestEditItem(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SItemEditor & END_CHAR)
End Sub

' ::::::::::::::::::::::
' :: Edit item packet ::
' ::::::::::::::::::::::
Private Sub HandleEditItem(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
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
    
End Sub

Private Sub HandleDelete(ByVal Index As Long, ByRef Parse() As String)
Dim Editor As Byte
Dim EditorIndex As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Editor = CByte(Parse(1))
    EditorIndex = CLng(Parse(2))

    Select Case Editor
    
        Case GameEditor.Item_
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then
                Call HackingAttempt(Index, "Invalid Item Index")
                Exit Sub
            End If
            
            Call ClearItem(EditorIndex)
            
            Call SendUpdateItemToAll(EditorIndex)
            Call SaveItem(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted item #" & EditorIndex & ".", ADMIN_LOG)
        
        Case GameEditor.NPC_
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then
                Call HackingAttempt(Index, "Invalid NPC Index")
                Exit Sub
            End If
        
            Call ClearNpc(EditorIndex)
        
            Call SendUpdateNpcToAll(EditorIndex)
            Call SaveNpc(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted npc #" & EditorIndex & ".", ADMIN_LOG)
        
        Case GameEditor.Spell_
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then
                Call HackingAttempt(Index, "Invalid Spell Index")
                Exit Sub
            End If
        
            Call ClearSpell(EditorIndex)
            
            Call SendUpdateSpellToAll(EditorIndex)
            Call SaveSpell(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted spell #" & EditorIndex & ".", ADMIN_LOG)
        
        Case GameEditor.Shop_
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_SHOPS Then
                Call HackingAttempt(Index, "Invalid Shop Index")
                Exit Sub
            End If
            
            Call ClearShop(EditorIndex)
            
            Call SendUpdateShopToAll(EditorIndex)
            Call SaveShop(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted shop #" & EditorIndex & ".", ADMIN_LOG)
            
        Case GameEditor.Sign_
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_SIGNS Then
                Call HackingAttempt(Index, "Invalid Sign Index")
                Exit Sub
            End If
            
            Call ClearSign(EditorIndex)
            
            Call SendUpdateSignToAll(EditorIndex)
            Call SaveSign(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted sign #" & EditorIndex & ".", ADMIN_LOG)
            
        Case GameEditor.Anim_
            ' Prevent hacking
            If EditorIndex < 1 Or EditorIndex > MAX_SIGNS Then
                Call HackingAttempt(Index, "Invalid Anim Index")
                Exit Sub
            End If
            
            Call ClearAnim(EditorIndex)
            
            Call SendUpdateAnimToAll(EditorIndex)
            Call SaveAnim(EditorIndex)
            Call AddLog(GetPlayerName(Index) & "Deleted anim #" & EditorIndex & ".", ADMIN_LOG)
            
    End Select
    
    Call SendDataTo(Index, SREditor & END_CHAR)

End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Private Sub HandleSaveItem(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long
Dim LoopI As Long
Dim PacketCount As Long
Dim LoopI2 As Long
Dim LoopI3 As Long
Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    n = Val(Parse(1))
    
    If n < 0 Or n > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Item Index")
        Exit Sub
    End If
    
    If Item(n).Type > ItemType.Shield_ Then GoTo skippy
    
    ' this huge loop updates the player's bonuses properly...
    For LoopI = 1 To MAX_PLAYERS
        If IsPlaying(LoopI) Then
            For LoopI2 = 1 To Equipment.Equipment_Count - 1
                If GetPlayerEquipmentSlot(Index, LoopI2) > 0 Then
                    If GetPlayerInvItemNum(Index, GetPlayerEquipmentSlot(Index, LoopI2)) = n Then
                        For i = 1 To Stats.Stat_Count - 1
                            AdjustStatBonus LoopI, i, Item(n).BuffStats(i), False
                        Next
                        For i = 1 To Vitals.Vital_Count - 1
                            AdjustVitalBonus LoopI, i, Item(n).BuffVitals(i), False
                        Next
                        SetPlayerEquipmentSlot LoopI, 0, LoopI2
                        GoTo skippy
                    End If
                End If
            Next
        End If
    Next
skippy:
    ' Update the item
    Item(n).Name = Parse(2)
    Item(n).Pic = Val(Parse(3))
    Item(n).Type = Val(Parse(4))
    Item(n).Data1 = Val(Parse(5))
    Item(n).Data2 = Val(Parse(6))
    Item(n).Data3 = Val(Parse(7))
    If Val(Parse(8)) = 0 Then Item(n).Durability = -1 Else Item(n).Durability = Val(Parse(8))
    Item(n).Anim = Val(Parse(9))
    Item(n).CostItem = Val(Parse(10))
    Item(n).CostAmount = Val(Parse(11))
    
    PacketCount = 12
    
    For LoopI = 1 To Stats.Stat_Count - 1
        Item(n).BuffStats(LoopI) = Val(Parse(PacketCount))
        PacketCount = PacketCount + 1
    Next
    
    For LoopI = 1 To Vitals.Vital_Count - 1
        Item(n).BuffVitals(LoopI) = Val(Parse(PacketCount))
        PacketCount = PacketCount + 1
    Next
    
    For LoopI = 0 To Item_Requires.Count - 1
        Item(n).Required(LoopI) = Val(Parse(PacketCount))
        PacketCount = PacketCount + 1
    Next
    
    If Item(n).Type > ItemType.Shield_ Then GoTo skippy2
    
    ' this huge loop updates the player's bonuses properly...
    For LoopI = 1 To MAX_PLAYERS
        If IsPlaying(LoopI) Then
            For LoopI3 = 1 To MAX_INV
                If GetPlayerInvItemNum(LoopI, LoopI3) > 0 Then
                    If GetPlayerInvItemNum(LoopI, LoopI3) = n Then
                        SetPlayerEquipmentSlot LoopI, LoopI3, Item(n).Type
                        For i = 1 To Stats.Stat_Count - 1
                            AdjustStatBonus LoopI, i, Item(n).BuffStats(i)
                        Next
                        For i = 1 To Vitals.Vital_Count - 1
                            AdjustVitalBonus LoopI, i, Item(n).BuffVitals(i)
                        Next
                        GoTo skippy2
                    End If
                End If
            Next
            SendWornEquipment LoopI
        End If
    Next
skippy2:
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
    
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Private Sub HandleRequestEditNpc(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SNpcEditor & END_CHAR)
End Sub

' :::::::::::::::::::::
' :: Edit npc packet ::
' :::::::::::::::::::::
Private Sub HandleEditNpc(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
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
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long
Dim i As Long
Dim ii As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
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
    
    Npc(n).Stat(Stats.Strength) = Val(Parse(11))
    Npc(n).Stat(Stats.Defense) = Val(Parse(12))
    Npc(n).Stat(Stats.SPEED) = Val(Parse(13))
    Npc(n).Stat(Stats.Magic) = Val(Parse(14))
    Npc(n).HP = Val(Parse(15))
    Npc(n).Experience = Val(Parse(16))
    Npc(n).GivesGuild = Val(Parse(17))
    
    ii = 18
    
    For i = 0 To UBound(Npc(n).Sound)
        Npc(n).Sound(i) = Parse(ii)
        ii = ii + 1
    Next
    
    For i = 0 To UBound(Npc(n).Reflection)
        Npc(n).Reflection(i) = Val(Parse(ii))
        ii = ii + 1
    Next
    
    ' Save it
    Call SendUpdateNpcToAll(n)
    Call SaveNpc(n)
    Call AddLog(GetPlayerName(Index) & " saved npc #" & n & ".", ADMIN_LOG)
    
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Private Sub HandleRequestEditShop(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SShopEditor & END_CHAR)
End Sub

' ::::::::::::::::::::::
' :: Edit shop packet ::
' ::::::::::::::::::::::
Private Sub HandleEditShop(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
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
End Sub

' ::::::::::::::::::::::
' :: Edit anim packet ::
' ::::::::::::::::::::::
Private Sub HandleEditAnim(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The shop #
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n < 0 Or n > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Anim Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing anim #" & n & ".", ADMIN_LOG)
    Call SendEditAnimTo(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Private Sub HandleSaveShop(ByVal Index As Long, ByRef Parse() As String)
Dim ShopNum As Long
Dim n As Long
Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
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
    
    n = 6
    For i = 1 To MAX_TRADES
        Shop(ShopNum).TradeItem(i).GiveItem = Val(Parse(n))
        Shop(ShopNum).TradeItem(i).GiveValue = Val(Parse(n + 1))
        Shop(ShopNum).TradeItem(i).GetItem = Val(Parse(n + 2))
        Shop(ShopNum).TradeItem(i).GetValue = Val(Parse(n + 3))
        n = n + 4
    Next
    
    ' Save it
    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)
    Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::::
' :: Request edit anim packet  ::
' :::::::::::::::::::::::::::::::
Private Sub HandleRequestEditAnim(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SAnimEditor & END_CHAR)
End Sub

' :::::::::::::::::::::::::::::::
' :: Request edit sign packet  ::
' :::::::::::::::::::::::::::::::
Private Sub HandleRequestEditSign(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SSignEditor & END_CHAR)
End Sub

' :::::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::::
Private Sub HandleRequestEditSpell(ByVal Index As Long)
    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    Call SendDataTo(Index, SSpellEditor & END_CHAR)
End Sub

' :::::::::::::::::::::::
' :: Edit spell packet ::
' :::::::::::::::::::::::
Private Sub HandleEditSpell(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
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
End Sub

' :::::::::::::::::::::::
' :: Edit sign packet  ::
' :::::::::::::::::::::::
Private Sub HandleEditSign(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' The sign #
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n < 1 Or n > MAX_SIGNS Then
        Call HackingAttempt(Index, "Invalid Sign Index")
        Exit Sub
    End If
    
    Call AddLog(GetPlayerName(Index) & " editing sign #" & n & ".", ADMIN_LOG)
    Call SendEditSignTo(Index, n)
    
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Private Sub HandleSaveSign(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long
Dim LoopI As Long
Dim PacketCount As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' sign #
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n < 1 Or n > MAX_SIGNS Then
        Call HackingAttempt(Index, "Invalid Sign Index")
        Exit Sub
    End If
    
    Sign(n).Name = Parse(2)
    
    PacketCount = 4
    
    ' Update the spell
    With Sign(n)
        ReDim Preserve .Section(0 To Val(Parse(3)))
        For LoopI = 0 To UBound(.Section)
            .Section(LoopI) = Parse(PacketCount)
            PacketCount = PacketCount + 1
        Next
    End With
    
    ' Save it
    Call SendUpdateSignToAll(n)
    Call SaveSign(n)
    Call AddLog(GetPlayerName(Index) & " saving sign #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Save anim packet  ::
' :::::::::::::::::::::::
Private Sub HandleSaveAnim(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    ' Anim #
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n < 0 Or n > MAX_SPELLS Then
        Call HackingAttempt(Index, "Invalid Anim Index")
        Exit Sub
    End If
    
    ' Update the spell
    With Animation(n)
        .Name = Parse(2)
        .Delay = Val(Parse(3))
        .Width = Val(Parse(4))
        .Height = Val(Parse(5))
        .Pic = Val(Parse(6))
    End With
    
    ' Save it
    Call SendUpdateAnimToAll(n)
    Call SaveAnim(n)
    Call AddLog(GetPlayerName(Index) & " saving anim #" & n & ".", ADMIN_LOG)
    
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Private Sub HandleSaveSpell(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Developer Then
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
    With Spell(n)
        .Name = Parse(2)
        .CastSound = Parse(3)
        .MPReq = Val(Parse(4))
        .Type = Val(Parse(5))
        .Anim = Val(Parse(6))
        .Icon = Val(Parse(7))
        .Range = Val(Parse(8))
        .AOE = Val(Parse(9))
        .Data1 = Val(Parse(10))
        .Data2 = Val(Parse(11))
        .Data3 = Val(Parse(12))
        .Timer = Val(Parse(13))
    End With
    
    ' Save it
    Call SendUpdateSpellToAll(n)
    Call SaveSpell(n)
    Call AddLog(GetPlayerName(Index) & " saving spell #" & n & ".", ADMIN_LOG)
    
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Private Sub HandleSetAccess(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long
Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Creator Then
        Call HackingAttempt(Index, "Trying to use powers not available")
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
        
            'check to see if they can change the person's access
            If GetPlayerAccess(n) >= GetPlayerAccess(Index) Then
                Call PlayerMsg(Index, "You cannot change somebody's access that is higher than your own.", Red)
                Exit Sub
            End If
            
            If GetPlayerAccess(n) <= 0 Then Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            
            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
            
            UpdatePlayerTable n
        Else
            Call PlayerMsg(Index, "Player is not online.", Color.White)
        End If
    Else
        Call PlayerMsg(Index, "Invalid access level.", Red)
    End If
End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Private Sub HandleWhosOnline(ByVal Index As Long)
    Call SendWhosOnline(Index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Private Sub HandleSetMotd(ByVal Index As Long, ByRef Parse() As String)

    ' Prevent hacking
    If GetPlayerAccess(Index) < StaffType.Mapper Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    MOTD = Trim$(Parse(1))
    Call PutVar(App.Path & "\data\motd.ini", "MOTD", "Msg", MOTD)
    Call GlobalMsg("MOTD changed to: " & MOTD, BrightCyan)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & MOTD, ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::
' :: Trade request packet ::
' ::::::::::::::::::::::::::
Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long
Dim i As Long
Dim X As Long

    ' Trade num
    n = Val(Parse(1))
    
    ' Prevent hacking
    If (n <= 0) Or (n > MAX_TRADES) Then
        Call HackingAttempt(Index, "Trade Request Modification")
        Exit Sub
    End If
    
    ' Index for shop
    i = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
    
    If i < 1 Or i > MAX_SHOPS Or Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type <> Tile_Type.Shop_ Then
        PlayerMsg Index, "You are not at a shop!", BrightRed
        Exit Sub
    End If
    
    ' Check if inv full
    X = FindOpenInvSlot(Index, Shop(i).TradeItem(n).GetItem)
    
    If X = 0 Then
        Call PlayerMsg(Index, "Trade unsuccessful! Your inventory is currently full.", BrightRed)
        Exit Sub
    End If
    
    ' Check if they have the item
    If HasItem(Index, Shop(i).TradeItem(n).GiveItem) >= Shop(i).TradeItem(n).GiveValue Then
        If Item(Shop(i).TradeItem(n).GetItem).Type = ItemType.Currency_ Then
            Call GiveItem(Index, Shop(i).TradeItem(n).GetItem, Shop(i).TradeItem(n).GetValue)
            Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem, Shop(i).TradeItem(n).GiveValue)
        Else
            X = 0
            Do Until FindOpenInvSlot(Index, Shop(i).TradeItem(n).GetItem) = 0 Or X = Shop(i).TradeItem(n).GetValue
                Call GiveItem(Index, Shop(i).TradeItem(n).GetItem, 0)
                X = X + 1
            Loop
            Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem, CInt(Shop(i).TradeItem(n).GiveValue * (X / Shop(i).TradeItem(n).GetValue)))
        End If
        
        Call PlayerMsg(Index, "You successfully bought " & X & " " & Trim$(Item(Shop(i).TradeItem(n).GetItem).Name) & " for " & CInt(Shop(i).TradeItem(n).GiveValue * (X / Shop(i).TradeItem(n).GetValue)) & " " & Trim$(Item(Shop(i).TradeItem(n).GiveItem).Name) & "!", Color.Yellow)
    Else
        Call PlayerMsg(Index, "Trade unsuccessful! You can't afford that!", BrightRed)
    End If
    
End Sub

' :::::::::::::::::::::
' :: Fix item packet ::
' :::::::::::::::::::::
Private Sub HandleFixItem(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long
Dim i As Long
Dim ItemNum As Long
Dim DurNeeded As Long
Dim GoldNeeded As Long

    ' Inv num
    n = Val(Parse(1))
    
    ' Prevent hacking
    If n <= 0 Or n > MAX_INV Then
        Call HackingAttempt(Index, "Fix item modification")
        Exit Sub
    End If
    
    ' check for bad data
    If GetPlayerInvItemNum(Index, n) <= 0 Or GetPlayerInvItemNum(Index, n) > MAX_ITEMS Then Exit Sub
    
    ' Make sure its a equipable item
    If Item(GetPlayerInvItemNum(Index, n)).Type < ItemType.Weapon_ Or Item(GetPlayerInvItemNum(Index, n)).Type > ItemType.Shield_ Then
        Call PlayerMsg(Index, "You can only fix equipment type items!", Color.BrightRed)
        Exit Sub
    End If
    
    ' Check if they have a full inventory
    If FindOpenInvSlot(Index, GetPlayerInvItemNum(Index, n)) <= 0 Then
        Call PlayerMsg(Index, "You have no inventory space left!", Color.BrightRed)
        Exit Sub
    End If
    
    ' Now check the rate of pay
    ItemNum = GetPlayerInvItemNum(Index, n)
    
    DurNeeded = Item(ItemNum).Durability - GetPlayerInvItemDur(Index, n)
    
    ' Check if they even need it repaired
    If DurNeeded <= 0 Then
        Call PlayerMsg(Index, "This item is in perfect condition!", Color.White)
        Exit Sub
    End If
    
    If Item(ItemNum).CostAmount = 0 Or Item(ItemNum).CostItem = 0 Then
        Call SetPlayerInvItemDur(Index, n, Item(ItemNum).Durability)
        Call PlayerMsg(Index, "Item has been totally restored for nothing!", BrightBlue)
        Exit Sub
    End If
    
    GoldNeeded = CInt(Item(ItemNum).CostAmount * (DurNeeded / Item(ItemNum).Durability))
    If GoldNeeded <= 0 Then GoldNeeded = 1
    
    ' Check if they have enough for at least one point
    If HasItem(Index, Item(ItemNum).CostItem) >= i Then
        ' Check if they have enough for a total restoration
        If HasItem(Index, Item(ItemNum).CostItem) >= GoldNeeded Then
            Call TakeItem(Index, Item(ItemNum).CostItem, GoldNeeded)
            Call SetPlayerInvItemDur(Index, n, Item(ItemNum).Durability)
            Call PlayerMsg(Index, "Item has been totally restored for " & GoldNeeded & " " & Trim$(Item(Item(ItemNum).CostItem).Name) & "!", BrightBlue)
        Else
            DurNeeded = Item(ItemNum).Durability * (Item(ItemNum).CostAmount / HasItem(Index, Item(ItemNum).CostItem))
            Call TakeItem(Index, Item(ItemNum).CostItem, HasItem(Index, Item(ItemNum).CostItem))
            Call SetPlayerInvItemDur(Index, n, GetPlayerInvItemDur(Index, n) + DurNeeded)
            Call PlayerMsg(Index, "Item has been partially restored for " & GoldNeeded & " " & Trim$(Item(Item(ItemNum).CostItem).Name) & "!", BrightBlue)
        End If
    Else
        Call PlayerMsg(Index, "Insufficient " & Trim$(Item(Item(ItemNum).CostItem).Name) & " to fix this item!", BrightRed)
    End If
    
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Private Sub HandleSearch(ByVal Index As Long, ByRef Parse() As String)
Dim X As Long
Dim Y As Long
Dim i As Long

    X = Val(Parse(1))
    Y = Val(Parse(2))
    
    ' Prevent subscript out of range
    If X < 0 Or X > MAX_MAPX Or Y < 0 Or Y > MAX_MAPY Then
        Exit Sub
    End If
    
    ' Check for a player
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = X And GetPlayerY(i) = Y Then
                    
                    If i <> Index Then
                        ' Consider the player
                        If GetPlayerLevel(i) >= GetPlayerLevel(Index) + 5 Then
                            Call PlayerMsg(Index, "You wouldn't stand a chance.", BrightRed)
                        Else
                            If GetPlayerLevel(i) > GetPlayerLevel(Index) Then
                                Call PlayerMsg(Index, "This one seems to have an advantage over you.", Color.Yellow)
                            Else
                                If GetPlayerLevel(i) = GetPlayerLevel(Index) Then
                                    Call PlayerMsg(Index, "This would be an even fight.", Color.White)
                                Else
                                    If GetPlayerLevel(Index) >= GetPlayerLevel(i) + 5 Then
                                        Call PlayerMsg(Index, "You could slaughter that player.", BrightBlue)
                                    Else
                                        If GetPlayerLevel(Index) > GetPlayerLevel(i) Then
                                            Call PlayerMsg(Index, "You would have an advantage over that player.", Color.Yellow)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    ' Change target
                    TempPlayer(Index).Target = i
                    TempPlayer(Index).TargetType = E_Target.Player_
                    
                    If i <> Index Then
                        Call PlayerMsg(Index, "Your target is now " & GetPlayerName(i) & ".", Color.Yellow)
                    Else
                        Call PlayerMsg(Index, "Your target is now yourself.", Color.Yellow)
                    End If
                    Exit Sub
                    
                End If
            End If
        End If
    Next
    
    ' Check for an item
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(GetPlayerMap(Index), i).Num > 0 Then
            If MapItem(GetPlayerMap(Index), i).X = X Then
                If MapItem(GetPlayerMap(Index), i).Y = Y Then
                    Call PlayerMsg(Index, "You see a " & Trim$(Item(MapItem(GetPlayerMap(Index), i).Num).Name) & ".", Color.Yellow)
                    Exit Sub
                End If
            End If
        End If
    Next
    
    ' Check for an npc
    For i = 1 To UBound(MapSpawn(GetPlayerMap(Index)).Npc)
        If MapNpc(GetPlayerMap(Index)).MapNpc(i).Num > 0 Then
            If MapNpc(GetPlayerMap(Index)).MapNpc(i).X = X Then
                If MapNpc(GetPlayerMap(Index)).MapNpc(i).Y = Y Then
                    ' Change target
                    TempPlayer(Index).Target = i
                    TempPlayer(Index).TargetType = E_Target.NPC_
                    Call PlayerMsg(Index, "Your target is now a " & Trim$(Npc(MapNpc(GetPlayerMap(Index)).MapNpc(i).Num).Name) & ".", Color.Yellow)
                    Exit Sub
                End If
            End If
        End If
    Next
    
    ' Check for sign
    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = Tile_Type.Sign_ Then
        PlayerMsg Index, "You see a sign.", Color.Yellow
    End If
    
End Sub

' ::::::::::::::::::
' :: Party packet ::
' ::::::::::::::::::
Private Sub HandleParty(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long

    n = FindPlayer(Parse(1))
    
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
        If GetPlayerAccess(Index) > StaffType.Monitor Then
            Call PlayerMsg(Index, "You can't join a party, you are an admin!", BrightBlue)
            Exit Sub
        End If
    
        If GetPlayerAccess(n) > StaffType.Monitor Then
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
        Call PlayerMsg(Index, "Player is not online.", Color.White)
    End If
End Sub

' :::::::::::::::::::::::
' :: Join party packet ::
' :::::::::::::::::::::::
Private Sub HandleJoinParty(ByVal Index As Long)
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
Private Sub HandleLeaveParty(ByVal Index As Long)
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
Private Sub HandleSpells(ByVal Index As Long)
    Call SendPlayerSpells(Index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Private Sub HandleCast(ByVal Index As Long, ByRef Parse() As String)
Dim n As Long
Dim LoopI As Long
Dim HasFailed As Boolean

    ' Spell slot
    n = Val(Parse(1))
    
    If Not CanCastSpell(Index, n) Then Exit Sub
    
    If Spell(Player(Index).Char(TempPlayer(Index).CharNum).Spell(n)).AOE = 0 Then
        If TempPlayer(Index).Target = 0 Or TempPlayer(Index).TargetType = E_Target.None Then
            PlayerMsg Index, "You have no target!", Color.BrightRed
            Exit Sub
        End If
    End If
    
    ' handling AOE spells
    If Player(Index).Char(TempPlayer(Index).CharNum).Spell(n) > 0 Then
        If Spell(Player(Index).Char(TempPlayer(Index).CharNum).Spell(n)).AOE = 1 And Spell(GetPlayerSpell(Index, n)).Type <> Spell_Type.GiveItem_ Then
            Dim OldTarget As Long
            Dim OldTargetType As Long
            
            OldTarget = TempPlayer(Index).Target
            OldTargetType = TempPlayer(Index).TargetType
            
            For LoopI = 1 To MAX_PLAYERS
                If Spell(GetPlayerSpell(Index, n)).Type = Spell_Type.SubHP_ Or _
                   Spell(GetPlayerSpell(Index, n)).Type = Spell_Type.SubMP_ Or _
                   Spell(GetPlayerSpell(Index, n)).Type = Spell_Type.SubSP_ Then
                    If LoopI = Index Then GoTo Skipper
                End If
                
                If IsPlaying(LoopI) Then
                    If GetPlayerMap(LoopI) = GetPlayerMap(Index) Then
                        If Spell(Player(Index).Char(TempPlayer(Index).CharNum).Spell(n)).Range > 0 And Not IsInRange(GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(LoopI), GetPlayerY(LoopI), Spell(Player(Index).Char(TempPlayer(Index).CharNum).Spell(n)).Range) Then GoTo Skipper
                        
                        TempPlayer(Index).Target = LoopI
                        TempPlayer(Index).TargetType = E_Target.Player_
                        
                        HasFailed = CastSpell(Index, n, True)
                    End If
                End If
Skipper:
            Next
            
            For LoopI = 1 To UBound(MapSpawn(GetPlayerMap(Index)).Npc)
                If MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Num > 0 Then
                    If Spell(Player(Index).Char(TempPlayer(Index).CharNum).Spell(n)).Range > 0 And Not IsInRange(GetPlayerX(Index), GetPlayerY(Index), MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).X, MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Y, Spell(Player(Index).Char(TempPlayer(Index).CharNum).Spell(n)).Range) Then GoTo Skipperrr
                    
                    TempPlayer(Index).Target = LoopI
                    TempPlayer(Index).TargetType = E_Target.NPC_
                    
                    HasFailed = CastSpell(Index, n, True)
                End If
Skipperrr:
            Next
            
            TempPlayer(Index).Target = OldTarget
            TempPlayer(Index).TargetType = OldTargetType
            
            GoTo SkipAhead
        End If
    End If
    
    ' handling normal spells
    HasFailed = CastSpell(Index, n)
    
SkipAhead:
    
    If HasFailed Then
        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " casts " & Trim$(Spell(Player(Index).Char(TempPlayer(Index).CharNum).Spell(n)).Name) & "!", Color.BrightBlue)
        
        SendDataTo Index, SCastSuccess & SEP_CHAR & n & END_CHAR
    Else
        PlayerMsg Index, "You failed to cast the spell!", Color.BrightRed
    End If
    
    ' Take away the mana points
    Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - Spell(GetPlayerSpell(Index, n)).MPReq)
    Call SendVital(Index, Vitals.MP)
    
    TempPlayer(Index).CastTimer(n) = GetTickCountNew
    TempPlayer(Index).CastedSpell = YES
    
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Private Sub HandleQuit(ByVal Index As Long)
    Call CloseSocket(Index)
End Sub

' ::::::::::::::::::::::::
' :: Send client config ::
' ::::::::::::::::::::::::
Private Sub HandleConfigPass(ByVal Index As Long, ByRef Parse() As String)

    If IsBanned(GetPlayerIP(Index), True) Then Exit Sub
    
    If Encryption(CONFIG_PASSWORD, Parse(1)) = CONFIG_PASSWORD Then
        SendGameOptions Index
        SendDataTo Index, SConfigPass & SEP_CHAR & 1 & END_CHAR
    Else
        SendDataTo Index, SConfigPass & SEP_CHAR & 0 & END_CHAR
    End If
    
End Sub

Private Sub HandleACPAction(ByVal Index As Long, ByRef Parse() As String)
Dim UseIndex As Long

    If GetPlayerAccess(Index) < StaffType.Monitor Then
        HackingAttempt Index, "Admin cloning."
        Exit Sub
    End If
    
    Select Case Val(Parse(1))
    
        Case ACP_Action.LevelSelf
            SetPlayerExp Index, GetPlayerExp(Index) + GetPlayerNextLevel(Index)
            CheckPlayerLevelUp Index
            
        Case ACP_Action.LevelTarget
            UseIndex = FindPlayer(Parse(2))
            
            If UseIndex < 1 Then
                PlayerMsg Index, "The player is offline!", Color.BrightRed
                Exit Sub
            End If
            
            SetPlayerExp UseIndex, GetPlayerExp(UseIndex) + GetPlayerNextLevel(UseIndex)
            CheckPlayerLevelUp UseIndex
            
        Case ACP_Action.SetTargetSprite
            UseIndex = FindPlayer(Parse(2))
            
            If UseIndex < 1 Then
                PlayerMsg Index, "The player is offline!", Color.BrightRed
                Exit Sub
            End If
            
            SetPlayerSprite UseIndex, Val(Parse(3))
            SendPlayerData UseIndex
            
        Case ACP_Action.CheckAccount
            If Not AccountExist(Parse(2)) Then
                PlayerMsg Index, "That account doesn't exist!", Color.BrightRed
                Exit Sub
            End If
            
            Dim TempAccount As AccountRec
            Dim FileName As String
            Dim F As Long
            Dim FF As Long
            Dim s As String
            Dim Name As String
            
            FileName = App.Path & "\accounts\" & Trim$(Parse(2)) & ".bin"
            
            F = FreeFile
            
            Open FileName For Binary As #F
                Get #F, , TempAccount
            Close #F
            
            PlayerMsg Index, "Account: [" & Trim$(TempAccount.Login) & "] + Password: [" & Trim$(TempAccount.Password) & "]", Color.Grey
            
            FileName = App.Path & "\data\bans.ini"
            
            If Not FileExist(FileName) Then
                PlayerMsg Index, "** " & Trim$(Name) & " is not banned **", Color.BrightRed
                Exit Sub
            End If
            
            For FF = 1 To Val(GetVar(FileName, "ACCOUNT", "Total"))
                If Trim$(GetVar(FileName, "ACCOUNT", "Account" & FF)) = Trim$(TempAccount.Login) Then
                    PlayerMsg Index, "** " & Trim$(GetVar(FileName, "ACCOUNT", "Account" & FF)) & " IS banned **", Color.BrightRed
                    Close #FF
                    Exit Sub
                End If
            Next
            
            PlayerMsg Index, "** " & Trim$(TempAccount.Login) & " is not banned **", Color.BrightBlue
            
        Case ACP_Action.GiveSelfPK
        
            If GetPlayerPK(Index) < 1 Then
                SetPlayerPK Index, 1
                Call GlobalMsg(GetPlayerName(Index) & " has been deemed a Player Killer!", Color.BrightRed)
            Else
                SetPlayerPK Index, 0
                Call GlobalMsg(GetPlayerName(Index) & " has lost his Player Killer status!", Color.BrightRed)
            End If
            
            SendPlayerData Index
            
        Case ACP_Action.GiveTargetPK
            UseIndex = FindPlayer(Parse(2))
            
            If UseIndex < 1 Then
                PlayerMsg Index, "The player is offline!", Color.BrightRed
                Exit Sub
            End If
            
            If GetPlayerPK(UseIndex) < 1 Then
                SetPlayerPK UseIndex, 1
                Call GlobalMsg(GetPlayerName(UseIndex) & " has been deemed a Player Killer!", Color.BrightRed)
            Else
                SetPlayerPK UseIndex, 0
                Call GlobalMsg(GetPlayerName(UseIndex) & " has lost his Player Killer status!", Color.BrightRed)
            End If
            
            SendPlayerData UseIndex
            
        Case ACP_Action.CheckInventory
            UseIndex = FindPlayer(Parse(2))
            
            If UseIndex < 1 Then
                PlayerMsg Index, "The player is offline!", Color.BrightRed
                Exit Sub
            End If
            
            Dim i As Long
            Dim OutputS As String
            
            For i = 1 To MAX_INV
                If GetPlayerInvItemNum(UseIndex, i) > 0 Then
                    If GetPlayerInvItemValue(UseIndex, i) = 0 Then
                        OutputS = OutputS & i & ") " & Trim$(Item(GetPlayerInvItemNum(UseIndex, i)).Name) & " (1), "
                    Else
                        OutputS = OutputS & i & ") " & Trim$(Item(GetPlayerInvItemNum(UseIndex, i)).Name) & " (" & GetPlayerInvItemValue(UseIndex, i) & "), "
                    End If
                End If
            Next
            
            PlayerMsg Index, GetPlayerName(UseIndex) & "'s inventory:", Color.DarkGrey
            PlayerMsg Index, Left$(OutputS, Len(OutputS) - 2), Color.Grey
            
        Case ACP_Action.MutePlayer
            UseIndex = FindPlayer(Parse(2))
            
            If UseIndex < 1 Then
                PlayerMsg Index, "The player is offline!", Color.BrightRed
                Exit Sub
            End If
            
            If GetPlayerAccess(UseIndex) >= GetPlayerAccess(Index) Then
                PlayerMsg Index, "You can't mute somebody that is your access or higher!", Color.BrightRed
                Exit Sub
            End If
            
            If Not Player(UseIndex).Char(TempPlayer(UseIndex).CharNum).Muted Then
                Player(UseIndex).Char(TempPlayer(UseIndex).CharNum).MuteTime = GetTickCountNew + (Val(Parse(3)) * 60000)
                Player(UseIndex).Char(TempPlayer(UseIndex).CharNum).Muted = True
                
                PlayerMsg UseIndex, "You have been muted for " & Val(Parse(3)) & " minutes by " & GetPlayerName(Index) & "!", Color.BrightRed
                PlayerMsg Index, "You have muted " & GetPlayerName(UseIndex) & " for " & Val(Parse(3)) & " minutes!", Color.BrightBlue
                TextAdd frmServer.txtText, GetPlayerName(Index) & " (access " & GetPlayerAccess(Index) & ") has muted " & GetPlayerName(UseIndex) & " (access " & GetPlayerAccess(UseIndex) & ") for " & Val(Parse(3)) & " minutes!"
            Else
                Player(UseIndex).Char(TempPlayer(UseIndex).CharNum).Muted = False
                Player(UseIndex).Char(TempPlayer(UseIndex).CharNum).MuteTime = 0
                
                PlayerMsg UseIndex, "You have been unmuted by " & GetPlayerName(Index) & "!", Color.BrightGreen
                PlayerMsg Index, "You have unmuted " & GetPlayerName(UseIndex) & "!", Color.BrightGreen
                TextAdd frmServer.txtText, GetPlayerName(Index) & " (access " & GetPlayerAccess(Index) & ") has unmuted " & GetPlayerName(UseIndex) & " (access " & GetPlayerAccess(UseIndex) & ")!"
            End If
            
            UpdatePlayerTable UseIndex
            
    End Select
    
End Sub

Private Sub HandleRCWarp(ByVal Index As Long, ByRef Parse() As String)
Dim X As Long
Dim Y As Long

    If GetPlayerAccess(Index) < StaffType.Mapper Then
        HackingAttempt Index, "Admin cloning."
        Exit Sub
    End If
    
    X = Val(Parse(1))
    Y = Val(Parse(2))
    
    PlayerWarp Index, GetPlayerMap(Index), X, Y
    
End Sub

Private Sub HandlePressReturn(ByVal Index As Long, ByRef Parse() As String)
Dim LoopI As Long

    Select Case GetPlayerDir(Index)
        Case E_Direction.Up_
            If GetPlayerY(Index) - 1 < 0 Then Exit Sub
        Case E_Direction.Down_
            If GetPlayerY(Index) + 1 > MAX_MAPY Then Exit Sub
        Case E_Direction.Left_
            If GetPlayerX(Index) - 1 < 0 Then Exit Sub
        Case E_Direction.Right_
            If GetPlayerX(Index) + 1 > MAX_MAPX Then Exit Sub
    End Select
    
    Select Case UBound(Parse)
        Case 0
            For LoopI = 1 To UBound(MapSpawn(GetPlayerMap(Index)).Npc)
                If MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Num > 0 Then
                    If Npc(MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Num).GivesGuild = YES Then
                        Select Case GetPlayerDir(Index)
                            Case E_Direction.Up_
                                If MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Y = GetPlayerY(Index) - 1 Then
                                    If MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).X = GetPlayerX(Index) Then
                                        If CanMakeGuild(Index, Trim$(Npc(MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Num).Name)) Then SendDataTo Index, SGuildCreation & END_CHAR
                                        Exit Sub
                                    End If
                                End If
                            Case E_Direction.Down_
                                If MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Y = GetPlayerY(Index) + 1 Then
                                    If MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).X = GetPlayerX(Index) Then
                                        If CanMakeGuild(Index, Trim$(Npc(MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Num).Name)) Then SendDataTo Index, SGuildCreation & END_CHAR
                                        Exit Sub
                                    End If
                                End If
                            Case E_Direction.Left_
                                If MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).X = GetPlayerX(Index) - 1 Then
                                    If MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Y = GetPlayerY(Index) Then
                                        If CanMakeGuild(Index, Trim$(Npc(MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Num).Name)) Then SendDataTo Index, SGuildCreation & END_CHAR
                                        Exit Sub
                                    End If
                                End If
                            Case E_Direction.Right_
                                If MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).X = GetPlayerX(Index) + 1 Then
                                    If MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Y = GetPlayerY(Index) Then
                                        If CanMakeGuild(Index, Trim$(Npc(MapNpc(GetPlayerMap(Index)).MapNpc(LoopI).Num).Name)) Then SendDataTo Index, SGuildCreation & END_CHAR
                                        Exit Sub
                                    End If
                                End If
                        End Select
                    End If
                End If
            Next
            
            With Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1)
                If .Type = Tile_Type.Sign_ Then
                    SendDataTo Index, SScrollingText & SEP_CHAR & Trim$(Sign(.Data1).Section(0)) & SEP_CHAR & .Data1 & SEP_CHAR & 0 & END_CHAR
                End If
            End With
            
        Case 2
            If Val(Parse(2)) + 1 > UBound(Sign(Val(Parse(1))).Section) Then
                SendDataTo Index, SScrollingText & END_CHAR
            Else
                With Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1)
                    SendDataTo Index, SScrollingText & SEP_CHAR & Trim$(Sign(.Data1).Section(Val(Parse(2)) + 1)) & SEP_CHAR & .Data1 & SEP_CHAR & Val(Parse(2)) + 1 & END_CHAR
                End With
            End If
            
    End Select
    
End Sub

Private Sub HandleGuildCreation(ByVal Index As Long, ByRef Parse() As String)
Dim GuildIndex As Long
Dim LoopI As Long
    
    If Player(Index).Char(TempPlayer(Index).CharNum).Guild > 0 Then
        PlayerMsg Index, "You are already in a guild, you can't make another until you disband your current one!", Color.BrightRed
        Exit Sub
    End If
    
    If Guild_Creation_Item > 0 Then
        For LoopI = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, LoopI) > 0 Then
                If GetPlayerInvItemNum(Index, LoopI) = Guild_Creation_Item Then
                    If GetPlayerInvItemValue(Index, LoopI) >= Guild_Creation_Cost Then
                        TakeItem Index, Guild_Creation_Item, Guild_Creation_Cost
                        Exit For
                    Else
                        PlayerMsg Index, "Sorry, but to start a guild it costs " & Guild_Creation_Cost & " " & Trim$(Item(Guild_Creation_Item).Name) & "!", Color.BrightRed
                        Exit Sub
                    End If
                End If
            End If
        Next
    End If
    
    If LoopI = MAX_INV + 1 Then
        PlayerMsg Index, "Sorry, but to start a guild it costs " & Guild_Creation_Cost & " " & Trim$(Item(Guild_Creation_Item).Name) & "!", Color.BrightRed
        Exit Sub
    End If
    
    GuildIndex = FindOpenGuildSlot
    
    If GuildIndex = 0 Then
        PlayerMsg Index, "Sorry, but no guilds are available to be made at this time!", Color.BrightRed
        Exit Sub
    End If
    
    For LoopI = 1 To MAX_GUILDS
        If Trim$(Guild(LoopI).Name) = Trim$(Parse(1)) Then
            PlayerMsg Index, "Sorry, but that guild name is taken!", Color.BrightRed
            Exit Sub
        End If
    Next
    
    With Guild(GuildIndex)
        ReDim .Member_Account(0 To 0)
        ReDim .Member_CharNum(0 To 0)
        
        .Name = Trim$(Parse(1))
        .TotalMembers = 1
        
        .Member_Account(0) = GetPlayerLogin(Index)
        .Member_CharNum(0) = TempPlayer(Index).CharNum
    End With
    
    SaveGuild GuildIndex
    
    Player(Index).Char(TempPlayer(Index).CharNum).Guild = GuildIndex
    Player(Index).Char(TempPlayer(Index).CharNum).GuildRank = Guild_Rank.Rank4
    
    GlobalMsg GetPlayerName(Index) & " has started a new guild: " & Trim$(Parse(1)) & "!", Color.BrightBlue
    
    If Guild_Creation_Item > 0 Then
        PlayerMsg Index, "You just spent " & Guild_Creation_Cost & " " & Trim$(Item(Guild_Creation_Item).Name) & " on making a guild!", Color.BrightBlue
    End If
    
    SendPlayerGuildToAll Index
    
End Sub

Private Sub HandleGuildDisband(ByVal Index As Long)
Dim LoopI As Long
Dim GuildLoginIndex As Long

    If Player(Index).Char(TempPlayer(Index).CharNum).GuildRank = Guild_Rank.Rank4 Then
        GlobalMsg Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Name & " has been disbanded!", Color.BrightBlue
        
        ClearGuild Player(Index).Char(TempPlayer(Index).CharNum).Guild
        SaveGuild Player(Index).Char(TempPlayer(Index).CharNum).Guild
        
        For LoopI = 1 To MAX_PLAYERS
            If IsPlaying(LoopI) Then
                If Player(LoopI).Char(TempPlayer(LoopI).CharNum).Guild = Player(Index).Char(TempPlayer(Index).CharNum).Guild Then
                    If LoopI <> Index Then
                        Player(LoopI).Char(TempPlayer(LoopI).CharNum).Guild = 0
                        Player(LoopI).Char(TempPlayer(LoopI).CharNum).GuildRank = 0
                    End If
                End If
            End If
        Next
        
        Player(Index).Char(TempPlayer(Index).CharNum).Guild = 0
        Player(Index).Char(TempPlayer(Index).CharNum).GuildRank = 0
        
        SendPlayerGuildToAll Index
    Else
        GlobalMsg GetPlayerName(Index) & " has left " & Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Name & "!", Color.BrightBlue
        
        For LoopI = 0 To UBound(Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Member_Account)
            If Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Member_Account(LoopI) = GetPlayerLogin(Index) Then
                GuildLoginIndex = LoopI
                Exit For
            End If
        Next
        
        If GuildLoginIndex > 0 Then
            Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Member_Account(GuildLoginIndex) = vbNullString
            Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Member_CharNum(GuildLoginIndex) = 0
        End If
        
        Player(Index).Char(TempPlayer(Index).CharNum).Guild = 0
        Player(Index).Char(TempPlayer(Index).CharNum).GuildRank = 0
        
        SendPlayerGuildToAll Index
    End If
    
End Sub

Private Sub HandleGuildInvite(ByVal Index As Long, ByRef Parse() As String)
Dim LoopI As Long
Dim TargetIndex As Long

    If Player(Index).Char(TempPlayer(Index).CharNum).GuildRank < Guild_Rank.Rank4 Then
        PlayerMsg Index, "You don't have the appropriate rank to invite members!", Color.BrightRed
        Exit Sub
    End If
    
    TargetIndex = FindPlayer(Parse(1))
    
    If TargetIndex < 1 Then
        PlayerMsg Index, "Player is offline!", Color.BrightRed
        Exit Sub
    End If
    
    If TempPlayer(TargetIndex).GInviteWaiting Then
        PlayerMsg Index, "This player is currently deciding if they want to join a guild, please be patient!", Color.BrightRed
        Exit Sub
    End If
    
    TempPlayer(TargetIndex).GInviteWaiting = True
    SendDataTo TargetIndex, SGuildInvite & SEP_CHAR & Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Name & SEP_CHAR & Player(Index).Char(TempPlayer(Index).CharNum).Guild & END_CHAR
    
End Sub

Private Sub HandleInviteResponse(ByVal Index As Long, ByRef Parse() As String)
Dim GuildMemberIndex As Long

    If Not TempPlayer(Index).GInviteWaiting Then Exit Sub
    
    TempPlayer(Index).GInviteWaiting = False
    
    If CBool(Parse(1)) = True Then
        GuildMemberIndex = FindOpenGuildMemberSlot(Val(Parse(2)))
        If GuildMemberIndex > 0 Then
Reserver:
            Player(Index).Char(TempPlayer(Index).CharNum).Guild = Val(Parse(2))
            Player(Index).Char(TempPlayer(Index).CharNum).GuildRank = Guild_Rank.Rank1
            Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Member_Account(GuildMemberIndex) = GetPlayerLogin(Index)
            Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Member_CharNum(GuildMemberIndex) = TempPlayer(Index).CharNum
        Else
            GuildMemberIndex = UBound(Guild(Val(Parse(2))).Member_Account) + 1
            ReDim Preserve Guild(Val(Parse(2))).Member_Account(0 To GuildMemberIndex)
            ReDim Preserve Guild(Val(Parse(2))).Member_CharNum(0 To GuildMemberIndex)
            GoTo Reserver
        End If
        SaveGuild Val(Parse(2))
        SendPlayerGuildToAll Index
        GlobalMsg GetPlayerName(Index) & " has joined " & Guild(Val(Parse(2))).Name & "!", Color.Green
    Else
        GlobalMsg GetPlayerName(Index) & " has declined to join " & Guild(Val(Parse(2))).Name & "!", Color.BrightBlue
    End If
    
End Sub

Private Sub HandleGuildPromoteDemote(ByVal Index As Long, ByRef Parse() As String)
Dim TargetIndex As Long
Dim LoopI As Long
Dim GuildLoginIndex As Long

    If Player(Index).Char(TempPlayer(Index).CharNum).Guild < 1 Then
        PlayerMsg Index, "You are not in a guild!", Color.BrightRed
        Exit Sub
    End If
    
    If Player(Index).Char(TempPlayer(Index).CharNum).GuildRank <> Guild_Rank.Rank4 Then
        If Val(Parse(1)) = 1 Or Val(Parse(1)) = 2 Then
            PlayerMsg Index, "You have to be rank 4 to demote or promote members!", Color.BrightRed
        Else
            PlayerMsg Index, "You have to be rank 4 to kick members!", Color.BrightRed
        End If
        Exit Sub
    End If
    
    TargetIndex = FindPlayer(Parse(2))
    
    If TargetIndex < 1 Then
        PlayerMsg Index, "That player is offline!", Color.BrightRed
        Exit Sub
    End If
    
    If Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).GuildRank >= Player(Index).Char(TempPlayer(Index).CharNum).GuildRank Then
        If Trim$(Guild(Player(Index).Char(TempPlayer(Index).CharNum).Guild).Member_Account(0)) <> GetPlayerLogin(Index) Then
            If Val(Parse(1)) = 1 Or Val(Parse(1)) = 2 Then
                PlayerMsg Index, "You cannot promote or demote other members with your rank or higher unless you are the owner!", Color.BrightRed
            Else
                PlayerMsg Index, "You cannot kick other members with your rank or higher unless you are the owner!", Color.BrightRed
            End If
            Exit Sub
        End If
    End If
    
    Select Case Val(Parse(1))
    
        'promote
        Case 1
            If Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).GuildRank + 1 < 5 Then
                Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).GuildRank = Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).GuildRank + 1
                PlayerMsg TargetIndex, "You have been promoted to rank " & Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).GuildRank & " by " & GetPlayerName(Index) & "!", Color.BrightBlue
                SendPlayerGuildToAll TargetIndex
            End If
            
        'demote
        Case 2
            If Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).GuildRank - 1 > 0 Then
                Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).GuildRank = Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).GuildRank - 1
                PlayerMsg TargetIndex, "You have been demoted to rank " & Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).GuildRank & " by " & GetPlayerName(Index) & "!", Color.BrightBlue
                SendPlayerGuildToAll TargetIndex
            End If
            
        'kick
        Case 3
            GlobalMsg GetPlayerName(TargetIndex) & " has been kicked from " & Guild(Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).Guild).Name & "!", Color.BrightBlue
            
            For LoopI = 1 To UBound(Guild(Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).Guild).Member_Account)
                If Trim$(Guild(Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).Guild).Member_Account(LoopI)) = GetPlayerLogin(TargetIndex) Then
                    GuildLoginIndex = LoopI
                    Exit For
                End If
            Next
            
            If GuildLoginIndex > 0 Then
                Guild(Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).Guild).Member_Account(GuildLoginIndex) = vbNullString
                Guild(Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).Guild).Member_CharNum(GuildLoginIndex) = 0
            End If
            
            Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).Guild = 0
            Player(TargetIndex).Char(TempPlayer(TargetIndex).CharNum).GuildRank = 0
            
            SendPlayerGuildToAll TargetIndex
            
    End Select
    
End Sub

Private Sub HandlePing(ByVal Index As Long)
    SendDataTo Index, SPing & END_CHAR
End Sub

Private Sub HandleLogout(ByVal Index As Long)
    ClearPlayer Index
End Sub

Private Sub HandleSellItem(ByVal Index As Long, ByRef Parse() As String)
Dim ItemNum As Long
Dim ItemSlot As Long
Dim ItemWorth As Long

    ItemSlot = CLng(Parse(1))
    
    If ItemSlot < 1 Or ItemSlot > MAX_INV Then
        HackingAttempt Index, "Invalid item slot"
        Exit Sub
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, ItemSlot)
    
    If Item(ItemNum).Type = ItemType.Currency_ Then
        PlayerMsg Index, "You cannot sell currency type items!", Color.BrightRed
        Exit Sub
    End If
    
    If Item(ItemNum).CostItem < 1 Or Item(ItemNum).CostAmount < 1 Then
        PlayerMsg Index, "This item isn't worth anything.", Color.BrightRed
        Exit Sub
    End If
    
    ItemWorth = CInt(Item(ItemNum).CostAmount * (frmServer.scrlSellBack.Value * 0.01))
    
    If ItemWorth < 1 Then
        PlayerMsg Index, "This item isn't worth anything.", Color.BrightRed
        Exit Sub
    End If
    
    TakeItem Index, ItemNum, GetPlayerInvItemValue(Index, ItemSlot)
    GiveItem Index, Item(ItemNum).CostItem, ItemWorth
    
    PlayerMsg Index, "You sold your " & Trim$(Item(ItemNum).Name) & " for " & ItemWorth & " " & Trim$(Item(Item(ItemNum).CostItem).Name) & "!", Color.Black
    
End Sub
