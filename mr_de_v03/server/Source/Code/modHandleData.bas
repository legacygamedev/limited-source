Attribute VB_Name = "modHandleData"
Option Explicit

Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SMsgGetClasses) = GetAddress(AddressOf HandleGetClasses)
    HandleDataSub(SMsgNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(SMsgLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(SMsgRequestEditEmoticon) = GetAddress(AddressOf HandleRequestEditEmoticon)
    HandleDataSub(SMsgEditEmoticon) = GetAddress(AddressOf HandleEditEmoticon)
    HandleDataSub(SMsgSaveEmoticon) = GetAddress(AddressOf HandleSaveEmoticon)
    HandleDataSub(SMsgCheckEmoticon) = GetAddress(AddressOf HandleCheckEmoticon)
    HandleDataSub(SMsgAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(SMsgDelChar) = GetAddress(AddressOf HandleDelChar)
    HandleDataSub(SMsgUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(SMsgSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SMsgEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(SMsgGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SMsgAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SMsgPartyMsg) = GetAddress(AddressOf HandlePartyMsg)
    HandleDataSub(SMsgPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMsgPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SMsgPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SMsgUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(SMsgUnequipSlot) = GetAddress(AddressOf HandleUnequipSlot)
    HandleDataSub(SMsgAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SMsgUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(SMsgPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(SMsgWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(SMsgWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(SMsgWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(SMsgSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(SMsgGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(SMsgClickWarp) = GetAddress(AddressOf HandleClickWarp)
    HandleDataSub(SMsgRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(SMsgMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMsgNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(SMsgMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(SMsgMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(SMsgMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(SMsgMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(SMsgKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(SMsgListBans) = GetAddress(AddressOf HandleListBans)
    HandleDataSub(SMsgBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(SMsgBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(SMsgRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(SMsgRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(SMsgEditItem) = GetAddress(AddressOf HandleEditItem)
    HandleDataSub(SMsgSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(SMsgRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(SMsgEditNpc) = GetAddress(AddressOf HandleEditNpc)
    HandleDataSub(SMsgSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(SMsgRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(SMsgEditShop) = GetAddress(AddressOf HandleEditShop)
    HandleDataSub(SMsgSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(SMsgRequestEditSpell) = GetAddress(AddressOf HandleRequestEditSpell)
    HandleDataSub(SMsgEditSpell) = GetAddress(AddressOf HandleEditSpell)
    HandleDataSub(SMsgSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(SMsgRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(SMsgEditAnimation) = GetAddress(AddressOf HandleEditAnimation)
    HandleDataSub(SMsgSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(SMsgSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(SMsgWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(SMsgSetMOTD) = GetAddress(AddressOf HandleSetMOTD)
    HandleDataSub(SMsgTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SMsgSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(SMsgParty) = GetAddress(AddressOf HandleParty)
    HandleDataSub(SMsgJoinParty) = GetAddress(AddressOf HandleJoinParty)
    HandleDataSub(SMsgLeaveParty) = GetAddress(AddressOf HandleLeaveParty)
    HandleDataSub(SMsgCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(SMsgRequestLocation) = GetAddress(AddressOf HandleRequestLocation)
    HandleDataSub(SMsgFix) = GetAddress(AddressOf HandleFix)
    HandleDataSub(SMsgChangeInvSlots) = GetAddress(AddressOf HandleChangeInvSlots)
    HandleDataSub(SMsgClearTarget) = GetAddress(AddressOf HandleClearTarget)
    HandleDataSub(SMsgGCreate) = GetAddress(AddressOf HandleGCreate)
    HandleDataSub(SMsgSetGMOTD) = GetAddress(AddressOf HandleSetGMOTD)
    HandleDataSub(SMsgGQuit) = GetAddress(AddressOf HandleGQuit)
    HandleDataSub(SMsgGDelete) = GetAddress(AddressOf HandleGDelete)
    HandleDataSub(SMsgGPromote) = GetAddress(AddressOf HandleGPromote)
    HandleDataSub(SMsgGDemote) = GetAddress(AddressOf HandleGDemote)
    HandleDataSub(SMsgGKick) = GetAddress(AddressOf HandleGKick)
    HandleDataSub(SMsgGInvite) = GetAddress(AddressOf HandleGInvite)
    HandleDataSub(SMsgGJoin) = GetAddress(AddressOf HandleGJoin)
    HandleDataSub(SMsgGDecline) = GetAddress(AddressOf HandleGDecline)
    HandleDataSub(SMsgGuildMsg) = GetAddress(AddressOf HandleGuildMsg)
    HandleDataSub(SMsgKill) = GetAddress(AddressOf HandleKill)
    HandleDataSub(SMsgSetBound) = GetAddress(AddressOf HandleSetBound)
    HandleDataSub(SMsgCancelSpell) = GetAddress(AddressOf HandleCancelSpell)
    HandleDataSub(SMsgRelease) = GetAddress(AddressOf HandleRelease)
    HandleDataSub(SMsgRevive) = GetAddress(AddressOf HandleRevive)
    HandleDataSub(SMsgRequestEditQuest) = GetAddress(AddressOf HandleRequestEditQuest)
    HandleDataSub(SMsgEditQuest) = GetAddress(AddressOf HandleEditQuest)
    HandleDataSub(SMsgSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(SMsgAcceptQuest) = GetAddress(AddressOf HandleAcceptQuest)
    HandleDataSub(SMsgCompleteQuest) = GetAddress(AddressOf HandleCompleteQuest)
    HandleDataSub(SMsgDropQuest) = GetAddress(AddressOf HandleDropQuest)
End Sub
 
Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
        
    If EncryptPackets Then
        Encryption_XOR_DecryptByte Data(), PacketKeys(Player(Index).PacketInIndex)
        Player(Index).PacketInIndex = Player(Index).PacketInIndex + 1
        If Player(Index).PacketInIndex > PacketEncKeys - 1 Then Player(Index).PacketInIndex = 0
    End If
            
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        HackingAttempt Index, "Packet Manipulation."
        Exit Sub
    End If
    
    If MsgType >= SMSG_COUNT Then
        HackingAttempt Index, "Packet Manipulation."
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Private Sub HandleGetClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not IsPlaying(Index) Then
        SendNewCharClasses Index
    End If
End Sub

Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String, Password As String
    
    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
        
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
            
            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                SendClientMsg Index, "Account / Password is too short. (Err: #8)", NewAccount
                Exit Sub
            End If
            
'            If Not IsAlpha(Name) Then
'                SendClientMsg Index, "Adventurer name contains invalid characters. (Err: #3)"
'                Exit Sub
'            End If
            
            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                AddAccount Index, Name, Password
                AddLog "Account " & Name & " has been created.", PLAYER_LOG
                SendChars Index
            Else
                SendClientMsg Index, "Sorry, that account is already taken. (Err: #6)", NewAccount
            End If
        End If
    End If
End Sub

Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String, Password As String, Version As String

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
        
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()

            Name = Buffer.ReadString
            Password = Buffer.ReadString
            Version = Buffer.ReadString
            
'            If IsMultiIPOnline(Current_IP(Index)) Then
'                SendAlertMsg Index, "Only one account per computer can be logged in."
'                Exit Sub
'            End If
            
            If InStr(Name, "/") Then
                SendClientMsg Index, "Name contains illegal characters.", Login
                Exit Sub
            End If
            
            If InStr(Name, "\") Then
                SendClientMsg Index, "Name contains illegal characters.", Login
                Exit Sub
            End If
            
            If IsMultiAccounts(Name) Then
                SendClientMsg Index, "Account is already online. (Err: #11)", Login
                Exit Sub
            End If
            
            If Len(Name) < 3 Or Len(Password) < 3 Then
                SendClientMsg Index, "Account / Password is too short. (Err: #8)", Login
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                SendClientMsg Index, "Account name does not exist. (Err: #9)", Login
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                SendClientMsg Index, "Incorrect password. (Err: #10)", Login
                Exit Sub
            End If

            'version
            If Version <> "3.0.7" Then
                SendClientMsg Index, "Incorrect client version. (Err: #12)", Login
                Exit Sub
            End If

            ' Everything went ok

            ' Load the player
            Update_Login Index, Name
            Update_Password Index, Password
            SendChars Index

            ' Show the player up on the socket status
            AddLog Current_Login(Index) & " has logged in from " & Current_IP(Index) & ".", PLAYER_LOG
            AddText frmServer.txtText, Current_Login(Index) & " has logged in from " & Current_IP(Index) & "."
        End If
    End If
End Sub

Private Sub HandleRequestEditEmoticon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If

    SendEmoticonEditor Index
End Sub

Private Sub HandleEditEmoticon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long

    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong

    If n < 1 Or n > MAX_EMOTICONS Then
        HackingAttempt Index, "Invalid Emoticon Index"
        Exit Sub
    End If

    AddLog Current_Name(Index) & " editing emoticon #" & n & ".", ADMIN_LOG
    SendEditEmoticonTo Index, n
End Sub

Private Sub HandleSaveEmoticon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long, Pic As Long
Dim Command As String
    
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
            
    n = Buffer.ReadLong
    Command = Buffer.ReadString
    Pic = Buffer.ReadLong
    
    If n < 1 Or n > MAX_EMOTICONS Then
        HackingAttempt Index, "Invalid Emoticon Index"
        Exit Sub
    End If

    Emoticons(n).Command = Command
    Emoticons(n).Pic = Pic
    
    SendUpdateEmoticonToAll (n)
    SaveEmoticon (n)
    AddLog Current_Name(Index) & " saved emoticon #" & n & ".", ADMIN_LOG
    
    ' Update the cache now
    CacheEmoticons
End Sub

Private Sub HandleCheckEmoticon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SendCheckEmoticon Index, Current_Map(Index), Buffer.ReadLong
End Sub

Private Sub HandleAddChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String
Dim Sex As Byte, ClassNum As Byte, CharNum As Byte
Dim SpriteNum As Long
    
    If Not IsPlaying(Index) Then
        
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
    
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        ClassNum = Buffer.ReadLong
        CharNum = Buffer.ReadLong
        SpriteNum = Buffer.ReadLong
        
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            HackingAttempt Index, "Invalid CharNum"
            Exit Sub
        End If
        
        ' Update the charnum
        Update_CharNum Index, CharNum
        LoadPlayer Index, Current_Login(Index)
        
        If Len(Trim$(Name)) < 3 Then
            SendClientMsg Index, "Adventurer name is too short. (Err: #8)", NewChar
            Exit Sub
        End If
        
        If Len(Trim$(Name)) > 20 Then
            SendClientMsg Index, "Adventurer name is too long. (Err: #8)", NewChar
            Exit Sub
        End If
        
        If Not IsAlpha(Name) Then
            SendClientMsg Index, "Adventurer name has invalid characters. (Err: #3)", NewChar
            Exit Sub
        End If
        
        ' Check if char already exists in slot
        If CharExist(Index) Then
            SendClientMsg Index, "Adventurer slot already taken. (Err: #13)", NewChar
            Exit Sub
        End If
        
        ' Check if name is already in use
        If FindChar(Name) Then
            SendClientMsg Index, "That adventurer already exists. (Err: #14)", NewChar
            Exit Sub
        End If
        
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            HackingAttempt Index, "Invalid CharNum"
            Exit Sub
        End If
    
        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            HackingAttempt Index, "Invalid Sex (dont laugh)"
            Exit Sub
        End If
        
        ' Prevent hacking
        If ClassNum < 0 Or ClassNum > MAX_CLASSES Then
            HackingAttempt Index, "Invalid Class"
            Exit Sub
        End If
        
        ' Prevent hacking
        If Sex = SEX_MALE Then
            If InStr(Class(ClassNum).MaleSprite, SpriteNum) = 0 Then
                HackingAttempt Index, "Invalid Sprite Number"
                Exit Sub
            End If
        Else
            If InStr(Class(ClassNum).FemaleSprite, SpriteNum) = 0 Then
                HackingAttempt Index, "Invalid Sprite Number"
                Exit Sub
            End If
        End If
    
        ' Everything went ok, add the character
        PlayerAddChar Index, Name, Sex, ClassNum, CharNum, SpriteNum
        AddLog "Character " & Name & " added to " & Current_Login(Index) & "'s account.", PLAYER_LOG
        SendChars Index
    End If
End Sub

Private Sub HandleDelChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim CharNum As Byte
    
    If Not IsPlaying(Index) Then
    
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        
        CharNum = Buffer.ReadByte
    
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            HackingAttempt Index, "Invalid CharNum"
            Exit Sub
        End If
        
        ' Have to set the char num for saving and such
        Update_CharNum Index, CharNum
        LoadPlayer Index, Player(Index).Login
        
        DeleteName Current_Name(Index)
        
        ClearChar Index
        SavePlayer Index
    
        AddLog "Character deleted on " & Current_Login(Index) & "'s account.", PLAYER_LOG
        SendChars Index
    End If
End Sub

Private Sub HandleUseChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim CharNum As Byte
Dim f As Long
    
    If Not IsPlaying(Index) Then
    
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
    
        CharNum = Buffer.ReadByte
        
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            HackingAttempt Index, "Invalid Character Number!"
            Exit Sub
        End If
    
        Update_CharNum Index, CharNum
        LoadPlayer Index, Player(Index).Login
        
        ' Check to make sure the character exists and if so, set it as its current char
        If CharExist(Index) Then
            
            JoinGame Index
        
            AddLog Current_Login(Index) & "/" & Current_Name(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG
            AddText frmServer.txtText, Current_Login(Index) & "/" & Current_Name(Index) & " has began playing " & GAME_NAME & "."
            UpdateCaption
            
'            ' Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
'            If Not FindChar(Current_Name(Index)) Then
'                f = FreeFile
'                Open App.Path & "\accounts\charlist.txt" For Append As #f
'                    Print #f, Current_Name(Index)
'                Close #f
'            End If
        Else
            SendClientMsg Index, "Adventurer does not exist! (Err: #15)", Chars
        End If
    End If
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
            HackingAttempt Index, "Say Text Modification"
            Exit Sub
        End If
    Next
    
    AddLog "Map #" & Current_Map(Index) & ": " & Current_Name(Index) & ": " & Msg, PLAYER_LOG
    SendMapMsg Current_Map(Index), Current_Name(Index) & ": " & Msg, SayColor
End Sub

Private Sub HandleEmoteMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
        
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
            HackingAttempt Index, "Emote Text Modification"
            Exit Sub
        End If
    Next
    
    AddLog "Map #" & Current_Map(Index) & ": " & Current_Name(Index) & " " & Msg, PLAYER_LOG
    SendMapMsg Current_Map(Index), Current_Name(Index) & Msg, SayColor
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String, s As String
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
        
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
            HackingAttempt Index, "Global Text Modification"
            Exit Sub
        End If
    Next

    s = "[Realm] " & Current_Name(Index) & ": " & Msg
    AddLog s, PLAYER_LOG
    SendGlobalMsg s, GlobalColor
    AddText frmServer.txtText, s
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim i As Long

    If Current_Access(Index) <= 0 Then
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
        
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
            HackingAttempt Index, "Admin Text Modification"
            Exit Sub
        End If
    Next
   
    AddLog "[Realm Master] " & Current_Name(Index) & ": " & Msg, ADMIN_LOG
    SendAdminMsg "[Realm Master] " & Current_Name(Index) & ": " & Msg, AdminColor
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String
Dim MsgTo As Long, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString
    
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
            HackingAttempt Index, "Player Msg Text Modification"
            Exit Sub
        End If
    Next
    
    ' Check if they are trying to talk to themselves
    If MsgTo <> Index Then
        If MsgTo > 0 Then
            AddLog Current_Name(Index) & " tells " & Current_Name(MsgTo) & ", " & Msg & "'", PLAYER_LOG
            SendPlayerMsg MsgTo, Current_Name(Index) & ": " & Msg, TellColor
            SendPlayerMsg Index, "[" & Current_Name(MsgTo) & "] " & Current_Name(Index) & ": " & Msg, Green
        Else
            'PlayerMsg(Index, "Adventurer is not in the realm.", AlertColor)
            SendActionMsg Current_Map(Index), "Adventurer not online.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
        End If
    Else
        AddLog "Map #" & Current_Map(Index) & ": " & Current_Name(Index) & " begins to mumble to himself, what a wierdo...", PLAYER_LOG
        SendMapMsg Current_Map(Index), Current_Name(Index) & " engages in meaningful conversation with thin air.", SayColor
    End If
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Dir As Byte
Dim Movement As Long
    
    ' Can't do while dead
    If Current_IsDead(Index) Then Exit Sub
    If Player(Index).GettingMap Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Dir = Buffer.ReadLong
    Movement = Buffer.ReadLong
    
    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        HackingAttempt Index, "Invalid Direction"
        Exit Sub
    End If
    
    ' Prevent hacking
    If Movement < 1 Or Movement > 2 Then
        HackingAttempt Index, "Invalid Movement"
        Exit Sub
    End If
    
    ' Can't move if they are dead
    ' If they aren't dead...
    If Current_IsDead(Index) Then
        Exit Sub
    End If
    
    ' If they are casting , cancel it
    CheckCasting Index
    
    ' We can now move the player
    PlayerMove Index, Dir, Movement
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Dir As Byte

    ' Can't do while dead
    If Current_IsDead(Index) Then Exit Sub
    If Player(Index).GettingMap = 1 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Dir = Buffer.ReadLong
        
    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        HackingAttempt Index, "Invalid Direction"
        Exit Sub
    End If
    
    Update_Dir Index, Dir
    SendPlayerDir Index
End Sub

Private Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim InvNum As Byte

    ' Doesn't matter if they are dead
    If Current_IsDead(Index) Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    InvNum = Buffer.ReadLong
    
    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        HackingAttempt Index, "Invalid InvNum"
        Exit Sub
    End If
    
    OnUseItem Index, InvNum
End Sub

Private Sub HandleUnequipSlot(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim ItemType As Byte
    
    ' Can't do while dead
    If Current_IsDead(Index) Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ItemType = Buffer.ReadLong
    
    OnUnequipSlot Index, ItemType
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim MapNum As Long

    ' Doesn't matter if they are dead
    If Current_IsDead(Index) Then Exit Sub
    
    ' Send this packet so they can see the person attacking
    SendAttack Index
    
    ' Try to attack a player
    MapNum = Current_Map(Index)
    For i = 1 To MapData(MapNum).MapPlayersCount
        If MapData(MapNum).MapPlayers(i) <> Index Then
            If CanAttackPlayer(Index, MapData(MapNum).MapPlayers(i)) Then
                AttackPlayer Index, MapData(MapNum).MapPlayers(i)
                Exit Sub
            End If
        End If
    Next
'    If Map(Current_Map(Index)).Moral = MAP_MORAL_NONE Then
'        For i = 1 To MAX_PLAYERS
'            ' Make sure we dont try to attack ourselves
'            If i <> Index Then
'                ' Can we attack the player?
'                If CanAttackPlayer(Index, i) Then
'                    AttackPlayer Index, i
'                    Exit Sub
'                End If
'            End If
'        Next
'    End If
    
    ' Try to attack a npc
    For i = 1 To MapData(Current_Map(Index)).NpcCount
        ' Can we attack the npc?
        If CanAttackNpc(Index, i) Then
            AttackNpc Index, i
            Exit Sub
        End If
    Next
End Sub

Private Sub HandleUseStatPoint(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim PointType As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    PointType = Buffer.ReadLong
        
    ' Prevent hacking
    If (PointType < 1) Or (PointType > Stats.Stat_Count) Then
        HackingAttempt Index, "Invalid Point Type"
        Exit Sub
    End If
            
    ' Make sure they have Points
    If Current_Points(Index) > 0 Then
        ' Take away a stat point
        Update_Points Index, Current_Points(Index) - 1
                
        Update_BaseStat Index, PointType, Current_BaseStat(Index, PointType) + 1
        SendActionMsg Current_Map(Index), "You have gained " & StatName(PointType) & "!", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index

        SendStats Index
        Update_Vitals Index
    Else
        SendActionMsg Current_Map(Index), "You have no skill Points.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
    End If
End Sub

Private Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String
Dim Msg As String
Dim n As Long, i As Long

    If Current_Access(Index) < ADMIN_MONITER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString

    i = FindPlayer(Name)
    If i > 0 Then
        Msg = "Account: " & Trim$(Player(i).Login) & ", Name: " & Current_Name(i)
        Msg = Msg & vbNewLine & "IP: " & Current_IP(i) & " Map: " & Current_Map(i) & " X: " & Current_X(i) & " Y: " & Current_Y(i)
        Msg = Msg & vbNewLine & "-=- Stats for " & Current_Name(i) & " the " & GetClassName(Current_Class(i)) & " -=-"
        Msg = Msg & vbNewLine & "Level: " & Current_Level(i) & "  Exp: " & Current_Exp(i) & "/" & Current_NextLevel(i)
        Msg = Msg & vbNewLine & "HP: " & Current_BaseVital(i, Vitals.HP) & "/" & Current_MaxVital(i, Vitals.HP) & "  MP: " & Current_BaseVital(i, Vitals.MP) & "/" & Current_MaxVital(i, Vitals.MP) & "  SP: " & Current_BaseVital(i, Vitals.SP) & "/" & Current_MaxVital(i, Vitals.SP)
        
        For n = 1 To Stats.Stat_Count
             Msg = Msg & vbNewLine & StatName(n) & ": " & Current_Stat(i, n) & " (" & Current_BaseStat(i, n) & " + " & Current_ModStat(i, n) & ")"
        Next
        
        Msg = Msg & vbNewLine & "Critical Hit Chance: " & Current_CritChance(i) & "%"
        Msg = Msg & vbNewLine & "Block Chance (Need shield): " & Current_BlockChance(i) & "%"
        Msg = Msg & vbNewLine & "Magic Damage: " & Current_MagicDamage(i)
        SendPlayerMsg Index, Msg, BrightGreen
    Else
        SendActionMsg Current_Map(Index), "Adventurer is not online.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
    End If
End Sub

Private Sub HandleWarpMeTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Name As String
Dim n As Long

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_MAPPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    ' The player
    n = FindPlayer(Name)
    If n <> Index Then
        If n > 0 Then
            PlayerWarp Index, Current_Position(n)
            AddLog Current_Name(Index) & " has warped to " & Current_Name(n) & ", map #" & Current_Map(n) & ".", ADMIN_LOG
        Else
            SendActionMsg Current_Map(Index), "Adventurer is not online.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
        End If
    End If
End Sub

Private Sub HandleWarpToMe(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim Name As String
    
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_MAPPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    ' The player
    n = FindPlayer(Name)
    If n <> Index Then
        If n > 0 Then
            PlayerWarp n, Current_Position(Index)
            AddLog Current_Name(Index) & " has warped " & Current_Name(n) & " to self, map #" & Current_Map(Index) & ".", ADMIN_LOG
        Else
            SendActionMsg Current_Map(Index), "Adventurer is not online.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
        End If
    End If
End Sub

Private Sub HandleWarpTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim NewPosition As PositionRec

     ' Prevent hacking
    If Current_Access(Index) < ADMIN_MAPPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' The map
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        HackingAttempt Index, "Invalid map"
        Exit Sub
    End If
    
    NewPosition.Map = n
    NewPosition.X = Current_X(Index)
    NewPosition.Y = Current_Y(Index)
    
    PlayerWarp Index, NewPosition
    AddLog Current_Name(Index) & " warped to map #" & n & ".", ADMIN_LOG
End Sub

Private Sub HandleSetSprite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim Name As String
Dim SpriteNum As Long

    If Current_Access(Index) < ADMIN_MAPPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    SpriteNum = Buffer.ReadLong
    
    n = FindPlayer(Name)
    If n > 0 Then
        Update_Sprite n, SpriteNum
        
        ' send to the player and map
        SendPlayerData n
        SendDataToMap Current_Map(n), PlayerData(n)
        
        AddLog Current_Name(Index) & " changed " & Current_Name(n) & " to sprite " & SpriteNum, ADMIN_LOG
    End If
End Sub

Private Sub HandleGetStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim Msg As String

    Msg = "-=- Stats for " & Current_Name(Index) & " the " & GetClassName(Current_Class(Index)) & " -=-"
    Msg = Msg & vbNewLine & "Level: " & Current_Level(Index) & "  Exp: " & Current_Exp(Index) & "/" & Current_NextLevel(Index)
    Msg = Msg & vbNewLine & "HP: " & Current_BaseVital(Index, Vitals.HP) & "/" & Current_MaxVital(Index, Vitals.HP) & "  MP: " & Current_BaseVital(Index, Vitals.MP) & "/" & Current_MaxVital(Index, Vitals.MP) & "  SP: " & Current_BaseVital(Index, Vitals.SP) & "/" & Current_MaxVital(Index, Vitals.SP)
    For i = 1 To Stats.Stat_Count
        Msg = Msg & vbNewLine & StatName(i) & ": " & Current_Stat(Index, i) & " (" & Current_BaseStat(Index, i) & " + " & Current_ModStat(Index, i) & ")"
    Next
    Msg = Msg & vbNewLine & "Critical Hit Chance: " & Current_CritChance(Index) & "%"
    Msg = Msg & vbNewLine & "Block Chance (Need shield): " & Current_BlockChance(Index) & "%"
    Msg = Msg & vbNewLine & "Magic Damage: " & Current_MagicDamage(Index)
    SendPlayerMsg Index, Msg, White
End Sub

Private Sub HandleClickWarp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim NewPosition As PositionRec
    
    If Current_Access(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    NewPosition.Map = Current_Map(Index)
    NewPosition.X = Buffer.ReadLong
    NewPosition.Y = Buffer.ReadLong
    
    PlayerWarp Index, NewPosition
End Sub

Private Sub HandleRequestNewMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Dir As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Dir = Buffer.ReadLong
        
    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        HackingAttempt Index, "Invalid Direction"
        Exit Sub
    End If
        
    PlayerMove Index, Dir, 1
End Sub

Private Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim MapNum As Long
Dim i As Long
Dim X As Long
Dim Y As Long
Dim TileSize As Long
Dim TileData() As Byte

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_MAPPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
        
    MapNum = Current_Map(Index)     ' What map are we editing?
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Buffer.DecompressBuffer
    
    With Map(MapNum)
        .Name = Buffer.ReadString
        .Revision = Buffer.ReadLong
        .Moral = Buffer.ReadByte
        .Up = Buffer.ReadInteger
        .Down = Buffer.ReadInteger
        .Left = Buffer.ReadInteger
        .Right = Buffer.ReadInteger
        .Music = Buffer.ReadByte
        .BootMap = Buffer.ReadInteger
        .BootX = Buffer.ReadByte
        .BootY = Buffer.ReadByte
        .TileSet = Buffer.ReadByte
        .MaxX = Buffer.ReadByte
        .MaxY = Buffer.ReadByte
        
        For X = 1 To MAX_MOBS
            .Mobs(X).NpcCount = Buffer.ReadLong
            ReDim .Mobs(X).Npc(.Mobs(X).NpcCount)
            
            If .Mobs(X).NpcCount > 0 Then
                For Y = 1 To .Mobs(X).NpcCount
                    .Mobs(X).Npc(Y) = Buffer.ReadLong
                Next
            End If
        Next
        
        ' set the Tile()
        ReDim .Tile(0 To .MaxX, 0 To .MaxY)
        
        TileSize = LenB(.Tile(0, 0)) * ((UBound(.Tile, 1) + 1) * (UBound(.Tile, 2) + 1))
        ReDim TileData(0 To TileSize - 1)
        TileData = Buffer.ReadBytes(TileSize)
        CopyMemory ByVal VarPtr(.Tile(0, 0)), ByVal VarPtr(TileData(0)), TileSize
    End With
    
    ' Manually setting the revision otherwise you won't download the new data
    Map(MapNum).Revision = Map(MapNum).Revision + 1
    
    ' Update the mapnpcs
    UpdateMapNpc MapNum
    
    For i = 1 To MapData(MapNum).NpcCount
        ClearMapNpc MapNum, i
    Next
    
    SendMapNpcsToMap MapNum
    SpawnMapNpcs MapNum
    
    ' Save the map
    SaveMap MapNum
    
    ' cache the map
    CacheMap MapNum
    
    ClearTempTile MapNum
    
    For i = 1 To MapData(MapNum).MapPlayersCount
        PlayerWarp MapData(MapNum).MapPlayers(i), Current_Position(MapData(MapNum).MapPlayers(i))
    Next
'    For i = 1 To MAX_PLAYERS
'        If IsPlaying(Index) Then
'            If Current_Map(i) = MapNum Then
'                PlayerWarp i, Current_Position(i)
'            End If
'        End If
'    Next
    
    AddLog Current_Name(Index) & " saved map " & MapNum & ".", ADMIN_LOG
End Sub

Private Sub HandleNeedMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Revision As Long
Dim MapNum As Long
Dim X As Long
Dim Y As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Revision = Buffer.ReadLong
            
    MapNum = Current_Map(Index)
            
    If Map(MapNum).Revision <> Revision Then
        SendMap Index, MapNum
    End If

    SendMapItemsTo Index, MapNum
    SendMapNpcsTo Index, MapNum
    SendJoinMap Index
    Player(Index).GettingMap = 0

    SendMapDone Index
    
    ' Check the player quests for any ExploreMap Quest Types
    OnUpdateQuestProgress Index, MapNum, 1, False, QuestTypes.ExploreMap
    
    'TODO : Add a joinmap sub
    ' Will send any open doors
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_KEY Then
                If MapData(MapNum).TempTile.DoorOpen(X, Y) Then
                    SendMapKey MapNum, X, Y, 1
                End If
            End If
        Next
    Next
    
End Sub

Private Sub HandleMapGetItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Doesn't matter if they are dead
    If Current_IsDead(Index) Then Exit Sub
    
    PlayerMapGetItem Index
End Sub

Private Sub HandleMapDropItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim InvNum As Byte
Dim Amount As Long

    ' Doesn't matter if they are dead
    If Current_IsDead(Index) Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    InvNum = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Then
        HackingAttempt Index, "Invalid InvNum"
        Exit Sub
    End If
            
    PlayerMapDropItem Index, InvNum, Amount, True
End Sub

Private Sub HandleMapRespawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
    
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_MAPPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        SpawnItemSlot i, 0, 0, Current_Map(Index), MapData(Current_Map(Index)).MapItem(i).X, MapData(Current_Map(Index)).MapItem(i).Y, 0
        ClearMapItem Current_Map(Index), i
    Next
    
    ' Respawn
    SpawnMapItems Current_Map(Index)
    SpawnMapNpcs Current_Map(Index)
    
    AddLog Current_Name(Index) & " has respawned map #" & Current_Map(Index), ADMIN_LOG
End Sub

Private Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim s As String
Dim i As Long, tMapStart As Long, tMapEnd As Long
    
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_MAPPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If

    s = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1
    
    For i = 1 To MAX_MAPS
        If Trim$(Map(i).Name) = vbNullString Then
            tMapEnd = tMapEnd + 1
        Else
            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(Str$(tMapStart)) & "-" & Trim$(Str$(tMapEnd - 1)) & ", "
            End If
            tMapStart = i + 1
            tMapEnd = i + 1
        End If
    Next
    
    s = s & Trim$(Str$(tMapStart)) & "-" & Trim$(Str$(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    
    SendPlayerMsg Index, s, Brown
End Sub

Private Sub HandleKickPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim Name As String

    ' Prevent hacking
    If Current_Access(Index) <= 0 Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    ' The player index
    n = FindPlayer(Name)
    
    If n <> Index Then
        If n > 0 Then
            If Current_Access(n) <= Current_Access(Index) Then
                SendGlobalMsg "[Realm Event] " & Current_Name(n) & " has been cast out of the realm by " & Current_Name(Index) & "!", ActionColor
                AddLog Current_Name(Index) & " has kicked " & Current_Name(n) & ".", ADMIN_LOG
                SendAlertMsg n, "You have been kicked by " & Current_Name(Index) & "!"
            Else
                SendActionMsg Current_Map(Index), "They have higher access.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
            End If
        Else
            SendActionMsg Current_Map(Index), "Adventurer is not online.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
        End If
    End If
End Sub

Private Sub HandleListBans(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Long, f As Long
Dim s As String, Name As String
    
    ' Prevent hacking
     If Current_Access(Index) < ADMIN_MAPPER Then
         HackingAttempt Index, "Admin Cloning"
         Exit Sub
     End If
    
     n = 1
     f = FreeFile
     Open App.Path & "\Data\banlist.txt" For Input As #f
     Do While Not EOF(f)
         Input #f, s
         Input #f, Name
        
         'sendplayermsg(Index, n & ": Banned IP " & s & " by " & Name, AlertColor)
         SendActionMsg Current_Map(Index), n & ": Banned IP " & s & " by " & Name, BrightBlue, ACTIONMSG_SCREEN, 0, 0, Index
         n = n + 1
     Loop
     Close #f
End Sub

Private Sub HandleBanDestroy(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_CREATOR Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    ZeroBanList
    'sendplayermsg(Index, "Ban list destroyed.", AlertColor)
    SendActionMsg Current_Map(Index), "Ban list destroyed.", BrightBlue, ACTIONMSG_SCREEN, 0, 0, Index
End Sub

Private Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim Name As String

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_MAPPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    ' The player index
    n = FindPlayer(Name)
    
    If n <> Index Then
        If n > 0 Then
            If Current_Access(n) <= Current_Access(Index) Then
                BanIndex n, Index
            Else
                SendActionMsg Current_Map(Index), "They have higher access.", BrightBlue, ACTIONMSG_SCREEN, 0, 0, Index
            End If
        Else
            SendActionMsg Current_Map(Index), "Adventurer is not online.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
        End If
    End If
End Sub

Private Sub HandleRequestEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_MAPPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    SendEditMap Index
End Sub

Private Sub HandleRequestEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    SendItemEditor Index
End Sub

Private Sub HandleEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim ItemNum As Long
    
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' The item #
    ItemNum = Buffer.ReadLong
    
    ' Prevent hacking
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        HackingAttempt Index, "Invalid Item Index"
        Exit Sub
    End If
    
    AddLog Current_Name(Index) & " editing item #" & ItemNum & ".", ADMIN_LOG
    SendEditItemTo Index, ItemNum
End Sub

Private Sub HandleSaveItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim ItemNum As Long

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ItemNum = Buffer.ReadLong
    
    ' Check if a valid item
    If ItemNum <= 0 Then Exit Sub
    If ItemNum > MAX_ITEMS Then Exit Sub
    
    ' Check the length of the buffer, make sure it's the same as the size needed
    If Buffer.Length <> ItemSize Then Exit Sub
    
    ' Set the item from the byte array
    Set_ItemData ItemNum, Buffer.ReadBytes(ItemSize)
    
    ' Save it
    SendUpdateItemToAll ItemNum
    SaveItem ItemNum
    
    ' Check to see if the saved item will affect any players
    For i = 1 To OnlinePlayersCount
        CheckPlayerInventoryItems OnlinePlayers(i)
        CheckEquippedItems OnlinePlayers(i)
        'Update mods
        Update_ModStats OnlinePlayers(i)
        Update_ModVitals OnlinePlayers(i)
    Next
'    For i = 1 To MAX_PLAYERS
'        If IsPlaying(i) Then
'            CheckPlayerInventoryItems (i)
'            CheckEquippedItems (i)
'            'Update mods
'            Update_ModStats i
'            Update_ModVitals i
'        End If
'    Next
    
    AddLog Current_Name(Index) & " saved item #" & ItemNum & ".", ADMIN_LOG
    
    ' Update the cache now
    CacheItems
End Sub

Private Sub HandleRequestEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    SendNpcEditor Index
End Sub

Private Sub HandleEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
    
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' The npc #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_NPCS Then
        HackingAttempt Index, "Invalid NPC Index"
        Exit Sub
    End If
    
    AddLog Current_Name(Index) & " editing npc #" & n & ".", ADMIN_LOG
    SendEditNpcTo Index, n
End Sub

Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim NpcNum As Long

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    NpcNum = Buffer.ReadLong
    
    If NpcNum <= 0 Then Exit Sub
    If NpcNum > MAX_NPCS Then Exit Sub
    
    ' Check the length of the buffer, make sure it's the same as the size needed
    If Buffer.Length <> NpcSize Then Exit Sub
    
    ' Set the Npc from the byte array
    Set_NpcData NpcNum, Buffer.ReadBytes(NpcSize)
    
    ' Save it
    SendUpdateNpcToAll NpcNum
    SaveNpc NpcNum
    AddLog Current_Name(Index) & " saved npc #" & NpcNum & ".", ADMIN_LOG
    
    ' Update the cache now
    CacheNpcs
End Sub

Private Sub HandleRequestEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    SendShopEditor Index
End Sub

Private Sub HandleEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
    
     ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' The shop #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_SHOPS Then
        HackingAttempt Index, "Invalid Shop Index"
        Exit Sub
    End If
    
    AddLog Current_Name(Index) & " editing shop #" & n & ".", ADMIN_LOG
    SendEditShopTo Index, n
End Sub

Private Sub HandleSaveShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim ShopNum As Long

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ShopNum = Buffer.ReadLong
    
    If ShopNum <= 0 Then Exit Sub
    If ShopNum > MAX_SHOPS Then Exit Sub
    
    ' Check the length of the buffer, make sure it's the same as the size needed
    If Buffer.Length <> ShopSize Then Exit Sub
    
    ' Set the Shop from the byte array
    Set_ShopData ShopNum, Buffer.ReadBytes(ShopSize)
    
    ' Save it
    SendUpdateShopToAll ShopNum
    SaveShop ShopNum
    AddLog Current_Name(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG
    
    ' Update the cache now
    CacheShops
End Sub

Private Sub HandleRequestEditSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    SendSpellEditor Index
End Sub

Private Sub HandleEditSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
    
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' The spell #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_SPELLS Then
        HackingAttempt Index, "Invalid Spell Index"
        Exit Sub
    End If
    
    AddLog Current_Name(Index) & " editing spell #" & n & ".", ADMIN_LOG
    SendEditSpellTo Index, n
End Sub

Private Sub HandleSaveSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim SpellNum As Long

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SpellNum = Buffer.ReadLong
    
    If SpellNum <= 0 Then Exit Sub
    If SpellNum > MAX_SPELLS Then Exit Sub
   
    ' Check the length of the buffer, make sure it's the same as the size needed
    If Buffer.Length <> SpellSize Then Exit Sub
    
    ' Set the Spell from the byte array
    Set_SpellData SpellNum, Buffer.ReadBytes(SpellSize)
    
    ' Save it
    SendUpdateSpellToAll SpellNum
    SaveSpell SpellNum
    AddLog Current_Name(Index) & " saving spell #" & SpellNum & ".", ADMIN_LOG
    
    ' Update the cache now
    CacheSpells
End Sub

Private Sub HandleRequestEditAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    SendAnimationEditor Index
End Sub

Private Sub HandleEditAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
    
     ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' The shop #
    n = Buffer.ReadLong
    
    ' Prevent hacking
    If n < 0 Or n > MAX_ANIMATIONS Then
        HackingAttempt Index, "Invalid Animation Index"
        Exit Sub
    End If
    
    AddLog Current_Name(Index) & " editing Animation #" & n & ".", ADMIN_LOG
    SendEditAnimationTo Index, n
End Sub

Private Sub HandleSaveAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim AnimationNum As Long

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    AnimationNum = Buffer.ReadLong
    
    If AnimationNum <= 0 Then Exit Sub
    If AnimationNum > MAX_ANIMATIONS Then Exit Sub
    
    ' Check the length of the buffer, make sure it's the same as the size needed
    If Buffer.Length <> AnimationSize Then Exit Sub
    
    ' Set the animation from the byte array
    Set_AnimationData AnimationNum, Buffer.ReadBytes(AnimationSize)
    
    ' Save it
    SendUpdateAnimationToAll AnimationNum
    SaveAnimation AnimationNum
    AddLog Current_Name(Index) & " saving Animation #" & AnimationNum & ".", ADMIN_LOG
    
    ' Update the cache now
    CacheAnimations
End Sub

Private Sub HandleSetAccess(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim Access As Byte
Dim Name As String

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_CREATOR Then
        HackingAttempt Index, "Trying to use powers not available"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    Access = Buffer.ReadByte
    
    ' Check for invalid access level
    If Access < 0 Then Exit Sub
    If Access > 4 Then Exit Sub
        
    ' The index
    n = FindPlayer(Name)
    If n > 0 Then
        SendGlobalMsg "[Realm Event] The adventurer " & Current_Name(n) & " has been granted a mastership of the realm.", JoinLeftColor
        
        Update_Access n, Access
        SendPlayerData n
        AddLog Current_Name(Index) & " has modified " & Current_Name(n) & "'s access.", ADMIN_LOG
    Else
        'sendplayermsg(Index, "Adventurer is not in the realm.", AlertColor)
        SendActionMsg Current_Map(Index), "Adventurer is not online.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
    End If
End Sub

Private Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendWhosOnline (Index)
End Sub

Private Sub HandleSetMOTD(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim MOTD As String
    
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_MAPPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MOTD = Buffer.ReadString
    
    GameMOTD = MOTD                                                            ' Change the motd in memory
    PutVar App.Path & "\Data\Core Files\Configuration.ini", "message of the day", "MOTD", MOTD
    SendGlobalMsg "[Breaking Realm News!] " & MOTD, Yellow
    AddLog Current_Name(Index) & " changed MOTD to: " & MOTD, ADMIN_LOG
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim ShopNum As Long
Dim TradeSlot As Long
Dim InvSlot As Long
Dim MapNum As Long
Dim NpcNum As Long
Dim MapNpcNum As Byte

    ' Doesn't matter if they are dead
    If Current_IsDead(Index) Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' Trade num
    MapNpcNum = Buffer.ReadLong
    ShopNum = Buffer.ReadLong          ' Shopnum
    TradeSlot = Buffer.ReadLong          ' Trade Item
    
    ' Prevent hacking
    If (TradeSlot <= 0) Or (TradeSlot > MAX_TRADES) Then
        HackingAttempt Index, "Trade Request Modification"
        Exit Sub
    End If
    
    MapNum = Current_Map(Index)
    NpcNum = MapData(MapNum).MapNpc(MapNpcNum).Num
    
    If NpcNum < 0 Then Exit Sub
    If NpcNum > MAX_NPCS Then Exit Sub
    
    ' Make sure it's a shopkeeper
    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then Exit Sub
    
    ' Make sure it's the same shop
    If Npc(NpcNum).Stat(Stats.Wisdom) <> ShopNum Then Exit Sub

    ' Make sure you're in range
    If Not NpcInRange(Index, MapNpcNum, 2) Then Exit Sub
            
    ' Check if inv full
    InvSlot = FindOpenInvSlot(Index, Shop(ShopNum).TradeItem(TradeSlot).GetItem)
    If InvSlot = 0 Then Exit Sub
    
    ' Check if they have the item
    'If HasItem(Index, Shop(i).TradeItem(N).GiveItem) >= Shop(i).TradeItem(N).GiveValue Then
    If CanTakeItem(Index, Shop(ShopNum).TradeItem(TradeSlot).GiveItem, Shop(ShopNum).TradeItem(TradeSlot).GiveValue) Then
        GiveItem Index, Shop(ShopNum).TradeItem(TradeSlot).GetItem, Shop(ShopNum).TradeItem(TradeSlot).GetValue
    Else
        'sendplayermsg(Index, "Trade unsuccessful.", ActionColor)
    End If
End Sub

Private Sub HandleSearch(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim X As Long
Dim Y As Long
Dim MapNum As Long
Dim NpcNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    X = Buffer.ReadLong
    Y = Buffer.ReadLong
    
    ' Prevent subscript out of range
    MapNum = Current_Map(Index)
    If X < 0 Then Exit Sub
    If X > Map(MapNum).MaxX Then Exit Sub
    If Y < 0 Then Exit Sub
    If Y > Map(MapNum).MaxY Then Exit Sub
    
    ' Check for a player
    For i = 1 To MapData(MapNum).MapPlayersCount
        If Current_X(MapData(MapNum).MapPlayers(i)) = X Then
            If Current_Y(MapData(MapNum).MapPlayers(i)) = Y Then
                If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
                    If Player(Index).Target = MapData(MapNum).MapPlayers(i) Then
                        ChangeTarget Index, 0, TARGET_TYPE_NONE
                    Else
                        ChangeTarget Index, MapData(MapNum).MapPlayers(i), TARGET_TYPE_PLAYER
                    End If
                Else
                    ChangeTarget Index, MapData(MapNum).MapPlayers(i), TARGET_TYPE_PLAYER
                End If
                Exit Sub
            End If
        End If
    Next
    
    ' Check for an npc
    For i = MapData(MapNum).NpcCount To 1 Step -1
        NpcNum = MapData(MapNum).MapNpc(i).Num
        If NpcNum > 0 Then
            If MapData(MapNum).MapNpc(i).X = X Then
                If MapData(MapNum).MapNpc(i).Y = Y Then
                    ' Check what type of npc
                    Select Case Npc(NpcNum).Behavior
                        Case NPC_BEHAVIOR_SHOPKEEPER
                            ' Using Magi to hold the shop num
                            If Npc(NpcNum).Stat(Stats.Wisdom) > 0 Then
                                ' Check if you're in range
                                If NpcInRange(Index, i, 2) Then
                                    SendTrade Index, i, Npc(NpcNum).Stat(Stats.Wisdom)
                                End If
                            End If
                            
                        Case NPC_BEHAVIOR_QUEST
                            ' Check if inrange and shit
                            ' send quest
                            If NpcInRange(Index, i, 2) Then
                                SendAvailableQuests Index, NpcNum, i
                            End If
                            
                        Case Else
                            ' Change target
                            If Player(Index).TargetType = TARGET_TYPE_NPC And Player(Index).Target = i Then
                                ChangeTarget Index, 0, TARGET_TYPE_NONE
                            Else
                                ChangeTarget Index, i, TARGET_TYPE_NPC
                            End If
                    End Select
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub

Private Sub HandleParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim Name As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    n = FindPlayer(Name)
    
    If n > 0 Then
        Party_Invite Index, n
    End If
End Sub

Private Sub HandleJoinParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_Join Index
End Sub

Private Sub HandleLeaveParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_Quit Index
End Sub

Private Sub HandleCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim SpellSlot As Long
Dim SpellNum As Long
Dim i As Long
    
    ' doesn't matter if they are already casting something
    If Player(Index).CastingSpell Then Exit Sub
    
    ' Doesn't matter if they are dead either
    If Current_IsDead(Index) Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' Spell slot
    SpellSlot = Buffer.ReadLong
    
    ' Check spellslot is valid
    If SpellSlot <= 0 Then Exit Sub
    If SpellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    ' Get the spell number of the spell you're trying to cast
    SpellNum = Current_Spell(Index, SpellSlot)
    If SpellNum <= 0 Then
        ' add error message
        CancelCastSpell Index
        Exit Sub
    End If
    
    ' first check if it's not on cooldown
    If Current_SpellCooldown(Index, SpellSlot) > 0 Then
        SendPlayerMsg Index, "Spell on cooldown.", White
        CancelCastSpell Index
        Exit Sub
    End If
    
    ' Prelim check for vital required to cast
    For i = 1 To Vitals.Vital_Count
        If Spell(SpellNum).VitalReq(i) > Current_BaseVital(Index, i) Then
            SendPlayerMsg Index, CStr(Spell(SpellNum).VitalReq(i)) & " " & VitalName(i) & " required.", BrightRed
            CancelCastSpell Index
            Exit Sub
        End If
    Next
           
    ' Prelim check for target
    ' Will check the target flags on the spell and make sure you have the appropriate target
    ' If Self cast - doesn't matter if you have a target or not - overrides other flags
    If Not Spell(SpellNum).TargetFlags And Targets.Target_SelfOnly Then
        ' If you have a target
        If Player(Index).Target > 0 Then
            ' If your target is a player check if the spell can be cast on players
            If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
                ' Check if it's hostile
                If (Spell(SpellNum).TargetFlags And Targets.Target_PlayerHostile) = Not Targets.Target_PlayerHostile Then
                    SendPlayerMsg Index, "Can not cast this spell on players.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
                
                ' Check if it's Target_PlayerParty
                ' We do this because Target_PlayerParty overrides the other player target flags
                If Spell(SpellNum).TargetFlags And Targets.Target_PlayerParty Then
                    ' Check if you're in a party
                    If Player(Index).InParty Then
                        ' Check if the target is in your party
                        If Player(Index).PartyIndex <> Player(Player(Index).Target).PartyIndex Then
                            SendPlayerMsg Index, "Can only cast this spell on party members.", BrightRed
                            Exit Sub
                        End If
                    ' Not in a party, can only cast on self then
                    Else
                        If Player(Index).Target <> Index Then
                            SendPlayerMsg Index, "Can only cast this spell on party members.", BrightRed
                            Exit Sub
                        End If
                    End If
                ' Since it's not a PlayerParty spell then we check the other player target flags
                Else
                    ' Check if it's beneficial
                    If (Spell(SpellNum).TargetFlags And Targets.Target_PlayerBeneficial) = Not Targets.Target_PlayerBeneficial Then
                        SendPlayerMsg Index, "Can not cast this spell on players.", BrightRed
                        CancelCastSpell Index
                        Exit Sub
                    End If
                    
                    ' If hostile - check if you can actually attack them
                    If (Spell(SpellNum).TargetFlags And Targets.Target_PlayerHostile) Then
                        If Not CheckAttackPlayer(Index, Player(Index).Target) Then
                            'sendplayermsg Index, "Can not cast this spell on players.", BrightRed
                            CancelCastSpell Index
                            Exit Sub
                        End If
                    End If
                End If
                
                ' Should mean they can cast on player - check their range now
                If Not PlayerInRange(Index, Current_X(Player(Index).Target), Current_Y(Player(Index).Target), Spell(SpellNum).Range) Then
                    SendPlayerMsg Index, "Target not in range.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
                   
                ' Now we check the spell type and if the player needs to be alive
                Select Case Spell(SpellNum).Type
                    ' Revive is the only spell type the target must be dead
                    Case SPELL_TYPE_REVIVE
                        ' Check if they are dead
                        If Not Current_IsDead(Player(Index).Target) Then
                            SendPlayerMsg Index, "Target not dead.", BrightRed
                            CancelCastSpell Index
                            Exit Sub
                        End If
                    ' All other spell types the player must be alive
                    Case Else
                        ' Can't cast on a dead player
                        If Current_IsDead(Player(Index).Target) Then
                            SendPlayerMsg Index, "Target is dead.", BrightRed
                            CancelCastSpell Index
                            Exit Sub
                        End If
                        
                End Select
                
            ' If your target is a npc check if the spell can be cast on npcs
            ElseIf Player(Index).TargetType = TARGET_TYPE_NPC Then
                ' Checks if the spell can be cast on NPCS
                If Not (Spell(SpellNum).TargetFlags And Targets.Target_Npc) = Targets.Target_Npc Then
                    SendPlayerMsg Index, "Can not cast this spell on npcs.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
                
                ' Check if it's a npc you can attack
                If Npc(MapData(Current_Map(Index)).MapNpc(Player(Index).Target).Num).Behavior = NPC_BEHAVIOR_FRIENDLY Then
                    SendPlayerMsg Index, "Can not cast this spell on friendly npcs.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
                
                ' Check if it's a npc you can attack
                If Npc(MapData(Current_Map(Index)).MapNpc(Player(Index).Target).Num).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
                    SendPlayerMsg Index, "Can not cast this spell on friendly npcs.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
                
                ' Check if it's a npc you can attack
                If Npc(MapData(Current_Map(Index)).MapNpc(Player(Index).Target).Num).Behavior = NPC_BEHAVIOR_QUEST Then
                    SendPlayerMsg Index, "Can not cast this spell on friendly npcs.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
                
                ' Should mean they can cast on npc - check their range now
                If Not NpcInRange(Index, Player(Index).Target, Spell(SpellNum).Range) Then
                    SendPlayerMsg Index, "Target not in range.", BrightRed
                    CancelCastSpell Index
                    Exit Sub
                End If
            End If
            
            ' Let's set our target since we can
            Player(Index).CastTarget = Player(Index).Target
            Player(Index).CastTargetType = Player(Index).TargetType
        ' No target
        Else
            ' For now we can't cast if we don't have a target
            SendPlayerMsg Index, "You need a target.", BrightRed
            CancelCastSpell Index
            Exit Sub
        End If
    Else
        ' Set our target to yourself
        Player(Index).CastTarget = Index
        Player(Index).CastTargetType = TARGET_TYPE_PLAYER
    End If
                      
    ' Checking if there is a cast time on a spell
    If Spell(SpellNum).CastTime > 0 Then
        Player(Index).CastingSpell = SpellSlot
        Player(Index).CastTime = GetTickCount + (Spell(SpellNum).CastTime * 1000)
    Else    ' If no cast time it's instant and we just cast the bitch
        OnCastSpell Index, SpellSlot
    End If
    
    ' Send this packet so they can see the person attacking
    SendAttack Index
End Sub

Private Sub HandleRequestLocation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
    If Current_Access(Index) < ADMIN_MAPPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    SendPlayerMsg Index, "[Mapeditor] Map: " & Current_Map(Index) & ", X: " & Current_X(Index) & ", Y: " & Current_Y(Index), AlertColor
End Sub

Private Sub HandleFix(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerWarp Index, Current_Position(Index)
    SendActionMsg Current_Map(Index), "Location request confirmed.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
End Sub

Private Sub HandleChangeInvSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim OldSlot As Long
Dim NewSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    OldSlot = Buffer.ReadLong
    NewSlot = Buffer.ReadLong
    
    PlayerSwitchInvSlots Index, OldSlot, NewSlot
End Sub

Private Sub HandleClearTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ChangeTarget Index, 0, TARGET_TYPE_NONE
End Sub

Private Sub HandleGCreate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long, f As Long, n As Long
Dim GuildName As String, GuildAbbreviation As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    GuildName = Buffer.ReadString
    GuildAbbreviation = Buffer.ReadString
    
    If Current_Guild(Index) > 0 Then
        SendPlayerMsg Index, "[Guild Info] You are in a guild already.", BrightGreen
        Exit Sub
    End If
     
    ' check guild abbreviation size
    If Len(GuildAbbreviation) <= 0 Then
        SendPlayerMsg Index, "[Guild Info] Guild abbreviation must be greater than 0 characters.", BrightGreen
        Exit Sub
    End If
    
    ' check guild abbreviation size
    If Len(GuildAbbreviation) > 3 Then
        SendPlayerMsg Index, "[Guild Info] Guild abbreviation must be 3 characters or less.", BrightGreen
        Exit Sub
    End If
    
    'check guild abbreviation
    If FindGuildAbbreviation(GuildAbbreviation) Then
        SendPlayerMsg Index, "[Guild Info] Guild abbreviation already exists.", BrightGreen
        Exit Sub
    End If
     
    ' check guild name size
    If Len(GuildName) <= 0 Then
        SendPlayerMsg Index, "[Guild Info] Guild name must be greater than 0 characters.", BrightGreen
        Exit Sub
    End If
    
    ' check guild abbreviation size
    If Len(GuildName) > NAME_LENGTH Then
        SendPlayerMsg Index, "[Guild Info] Guild names must be 20 characters or less.", BrightGreen
        Exit Sub
    End If
    
    'check guild name
    If FindGuildName(GuildName) Then
        SendPlayerMsg Index, "[Guild Info] Guild name already exists.", BrightGreen
        Exit Sub
    End If
    
    For i = 1 To MAX_GUILDS
        If Guild(i).Guild = 0 Then
            ' Use TakeCurrency() if you want to take away currency instead
            'If TakeStackedItem(Index, 1, 28) = True Then
            If CanTakeItem(Index, 1, 50) = True Then
                Guild(i).Guild = 1
                Guild(i).GuildName = GuildName
                Guild(i).GuildAbbreviation = GuildAbbreviation
                Guild(i).GMOTD = "Please set a GMOTD."
                Guild(i).Owner = Current_Name(Index)
                ' Append name to file
                f = FreeFile
                Open App.Path & "\Data\Guilds\GuildName.txt" For Append As #f
                    Print #f, GuildName
                Close #f
                
                f = FreeFile
                Open App.Path & "\Data\Guilds\GuildAbbreviation.txt" For Append As #f
                    Print #f, GuildAbbreviation
                Close #f
                
                For n = 1 To MAX_GUILD_RANKS
                    Guild(i).Rank(n) = "Rank" & n
                Next
            
                Update_Guild Index, i
                Update_GuildRank Index, 1
                Update_GuildName Index, GuildName
                SendPlayerGuild (Index)
                SendPlayerData (Index)
                
                SaveGuild (i)
                
                SendPlayerMsg Index, "[Guild Info] You have formed " & GuildName, BrightGreen
                'GlobalMsg("[Guild Info] " & Current_Name(Index) & " has formed the guild: " & GuildName & ".", Green)
                SendPlayerMsg Index, "[" & GuildName & "] " & GetGuildGMOTD(Current_Guild(Index)), BrightGreen
                AddLog Current_Name(Index) & " has created the Guild: " & GuildName & ".", PLAYER_LOG
                Exit Sub
            Else
                ' Put whatever message you want here
                SendPlayerMsg Index, "[Guild Info] You cannot start a guild right now.", BrightGreen
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub HandleSetGMOTD(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim GMOTD As String

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    GMOTD = Buffer.ReadString
    
    If Current_Guild(Index) > 0 Then
        If Current_GuildRank(Index) = 1 Or GetGuildOwner(Current_Guild(Index)) = Current_Name(Index) Then
            Guild(Current_Guild(Index)).GMOTD = GMOTD
            SaveGuild (Current_Guild(Index))
            SendGuildMsg Current_Guild(Index), "[Breaking Guild News!] " & GMOTD, BrightGreen
            AddLog Current_Name(Index) & " changed GMOTD to: " & GMOTD, PLAYER_LOG
            Exit Sub
        Else
            SendPlayerMsg Index, "[Guild Info] You are not high enough rank to do this.", BrightGreen
            Exit Sub
        End If
    Else
        SendPlayerMsg Index, "[Guild Info] You are not in a guild.", BrightGreen
        Exit Sub
    End If
End Sub

Private Sub HandleGQuit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
     If Current_Guild(Index) > 0 Then
        If GetGuildOwner(Current_Guild(Index)) = Current_Name(Index) Then
            SendPlayerMsg Index, "[Guild Info] You can not leave your own Guild. You must delete it first.", BrightGreen
            Exit Sub
        Else
            SendGuildMsg Current_Guild(Index), "[Guild Info] " & Current_Name(Index) & " has left the Guild.", BrightGreen
            AddLog Current_Name(Index) & " has left the guild: " & Current_GuildName(Current_Guild(Index)), ADMIN_LOG
            Update_Guild Index, 0
            Update_GuildRank Index, 0
            Update_GuildName Index, vbNullString
            SendPlayerData Index
            Exit Sub
        End If
    Else
        SendPlayerMsg Index, "[Guild Info] You are not in a guild.", BrightGreen
        Exit Sub
    End If
End Sub

Private Sub HandleGDelete(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim GuildNum As Byte
Dim i As Long
Dim GuildName As String

    GuildNum = Current_Guild(Index)
    
    If GuildNum > 0 Then
        If GetGuildOwner(Current_Guild(Index)) = Current_Name(Index) Then
            For i = 1 To OnlinePlayersCount
                If Current_Guild(OnlinePlayers(i)) = GuildNum Then
                    Update_Guild OnlinePlayers(i), 0
                    Update_GuildRank OnlinePlayers(i), 0
                    Update_GuildName OnlinePlayers(i), vbNullString
                    SendPlayerGuild OnlinePlayers(i)
                    SendPlayerData OnlinePlayers(i)
                End If
            Next
'            For i = 1 To MAX_PLAYERS
'                If IsPlaying(i) Then
'                    If Current_Guild(i) = GuildNum Then
'                        Update_Guild i, 0
'                        Update_GuildRank i, 0
'                        Update_GuildName i, vbNullString
'                        SendPlayerGuild i
'                        SendPlayerData i
'                    End If
'                End If
'            Next
            
            GuildName = GetGuildName(GuildNum)
            
            SendGuildMsg GuildNum, "[Guild Info] The guild has been disbanded.", BrightGreen
            
            DeleteGuildName (GuildName)
            DeleteGuildAbbreviation (GetGuildAbbreviation(GuildNum))
            ClearGuild (GuildNum)
            SaveGuild (GuildNum)
            
            AddLog Current_Name(Index) & " has deleted the guild: " & GuildName & ".", PLAYER_LOG
        Else
            SendPlayerMsg Index, "[Guild Info] You must be the guild owner to do that.", BrightGreen
        End If
    Else
        SendPlayerMsg Index, "[Guild Info] You are not in a guild.", BrightGreen
    End If
End Sub

Private Sub HandleGPromote(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim Name As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    n = FindPlayer(Name)
        
    If n <= 0 Then Exit Sub
        
    ' Prevent promoting self
    If n = Index Then Exit Sub
        
    If Current_Guild(Index) = 0 Then
        SendPlayerMsg Index, "[Guild Info] You are not in a guild.", BrightGreen
        Exit Sub
    End If
    
    If Current_Guild(n) = 0 Or Current_Guild(Index) <> Current_Guild(n) Then
        SendPlayerMsg Index, "[Guild Info] Player is not in your guild.", BrightGreen
        Exit Sub
    End If
    
    If Current_GuildRank(n) = 1 Then
        SendPlayerMsg Index, "[Guild Info] Player is max rank.", BrightGreen
        Exit Sub
    End If
    
    If Current_GuildRank(Index) <= 3 Then
        If Current_GuildRank(Index) > Current_GuildRank(n) Or Current_GuildRank(Index) = Current_GuildRank(n) Then
            SendPlayerMsg Index, "[Guild Info] You can't promote that player.", BrightGreen
            Exit Sub
        Else
            SendGuildMsg Current_Guild(Index), "[Guild Info] " & Current_Name(n) & " has been promoted by " & Current_Name(Index) & ".", BrightGreen
            Update_GuildRank n, Current_GuildRank(n) - 1
            SendPlayerData (n)
            Exit Sub
        End If
    Else
        SendPlayerMsg Index, "[Guild Info] You are not high enough rank to do that.", BrightGreen
        Exit Sub
    End If
End Sub

Private Sub HandleGDemote(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim Name As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    n = FindPlayer(Name)
        
    If n <= 0 Then Exit Sub
        
    ' Prevent demoting self
    If n = Index Then Exit Sub
    
    If Current_Guild(Index) = 0 Then
        SendPlayerMsg Index, "[Guild Info] You are not in a guild.", BrightGreen
        Exit Sub
    End If
    
    If Current_Guild(n) = 0 Or Current_Guild(Index) <> Current_Guild(n) Then
        SendPlayerMsg Index, "[Guild Info] Player is not in your guild.", BrightGreen
        Exit Sub
    End If
    
    If Current_GuildRank(n) = MAX_GUILD_RANKS Then
        SendPlayerMsg Index, "[Guild Info] Player is lowest rank.", BrightGreen
        Exit Sub
    End If
    
    If Current_GuildRank(Index) <= 3 Then
        If GetGuildOwner(Current_Guild(Index)) = Current_Name(Index) Then
            SendGuildMsg Current_Guild(Index), "[Guild Info] " & Current_Name(n) & " has been demoted by " & Current_Name(Index) & ".", BrightGreen
            Update_GuildRank n, Current_GuildRank(n) + 1
            SendPlayerData (n)
            Exit Sub
        Else
            If Current_GuildRank(Index) > Current_GuildRank(n) Or Current_GuildRank(Index) = Current_GuildRank(n) Then
                SendPlayerMsg Index, "[Guild Info] You can't demote that player.", BrightGreen
                Exit Sub
            Else
                SendGuildMsg Current_Guild(Index), "[Guild Info] " & Current_Name(n) & " has been demoted by " & Current_Name(Index) & ".", BrightGreen
                Update_GuildRank n, Current_GuildRank(n) + 1
                SendPlayerData (n)
                Exit Sub
            End If
        End If
    Else
        SendPlayerMsg Index, "[Guild Info] You are not high enough rank to do that.", BrightGreen
        Exit Sub
    End If
End Sub

Private Sub HandleGKick(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim Name As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    n = FindPlayer(Name)
        
    If n <= 0 Then Exit Sub
    
    ' Prevent promoting self
    If n = Index Then Exit Sub
    
    If Current_Guild(Index) = 0 Then
        SendPlayerMsg Index, "[Guild Info] You are not in a guild.", BrightGreen
        Exit Sub
    End If
    
    If Current_Guild(n) = 0 Or Current_Guild(Index) <> Current_Guild(n) Then
        SendPlayerMsg Index, "[Guild Info] Player is not in your guild.", BrightGreen
        Exit Sub
    End If
    
    If Current_GuildRank(Index) <= 3 Then
        If GetGuildOwner(Current_Guild(Index)) = Current_Name(Index) Then
            SendGuildMsg Current_Guild(Index), "[Guild Info] " & Current_Name(n) & " has been kicked out of the Guild.", BrightGreen
            Update_Guild n, 0
            Update_GuildRank n, 0
            Update_GuildName n, vbNullString
            SendPlayerData n
            Exit Sub
        Else
            If Current_GuildRank(Index) > Current_GuildRank(n) Or Current_GuildRank(Index) = Current_GuildRank(n) Then
                SendPlayerMsg Index, "[Guild Info] You can't kick that player.", BrightGreen
                Exit Sub
            Else
                SendGuildMsg Current_Guild(Index), "[Guild Info] " & Current_Name(n) & " has been kicked out of the Guild.", BrightGreen
                Update_Guild n, 0
                Update_GuildRank n, 0
                Update_GuildName n, vbNullString
                SendPlayerData n
                Exit Sub
            End If
        End If
    Else
        SendPlayerMsg Index, "[Guild Info] You are not high enough rank to do that.", BrightGreen
        Exit Sub
    End If
End Sub

Private Sub HandleGInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim Name As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    n = FindPlayer(Name)
    
    If n <= 0 Then Exit Sub
    
    ' Prevent inviting self
    If n = Index Then Exit Sub
    
    If Current_Guild(Index) = 0 Then
        SendPlayerMsg Index, "[Guild Info] You are not in a guild.", BrightGreen
        Exit Sub
    End If
    
    If Current_GuildRank(Index) <= 3 Then
        If n > 0 Then
            'check to see if they are invite to a guild already
            If Player(Index).GuildInvite = 0 Then
                ' Check to see if player is already in a guild
                If Current_Guild(n) = 0 Then
                    SendPlayerMsg Index, "[Guild Info] Guild request has been sent to " & Current_Name(n) & ".", BrightGreen
                    SendPlayerMsg n, "[Guild Info] " & Current_Name(Index) & " wants you to join their guild.  Type /gjoin to join, or /gdecline to decline.", BrightGreen
                
                    Player(n).GuildInvite = Current_Guild(Index)
                    Player(n).GuildInviter = Index
                    Exit Sub
                Else
                    SendPlayerMsg Index, "[Guild Info] Player is already in a guild!", BrightGreen
                    Exit Sub
                End If
            Else
                SendPlayerMsg Index, "[Guild Info] Player has been invited a guild already.", BrightGreen
            End If
        Else
            SendPlayerMsg Index, "[Guild Info] Player is not online.", BrightGreen
        End If
        Exit Sub
    Else
        SendPlayerMsg Index, "[Guild Info] You are not high enough rank to do that.", BrightGreen
        Exit Sub
    End If
End Sub

Private Sub HandleGJoin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Player(Index).GuildInvite > 0 Then
        Update_Guild Index, Player(Index).GuildInvite
        Update_GuildRank Index, MAX_GUILD_RANKS
        Update_GuildName Index, GetGuildName(Player(Index).GuildInvite)
        SendPlayerGuild (Index)
        SendPlayerData (Index)
        
        Player(Index).GuildInvite = 0
        Player(Index).GuildInviter = 0
        
        SendGuildMsg Current_Guild(Index), "[Guild Info] " & Current_Name(Index) & " has joined the Guild.", BrightGreen
        SendPlayerMsg Index, "[Guild News] " & GetGuildGMOTD(Current_Guild(Index)), BrightGreen
        AddLog "[Guild Info] " & Current_Name(Index) & " has joined the Guild.", PLAYER_LOG
    Else
        SendPlayerMsg Index, "[Guild Info] You have not been invited to a guild.", BrightGreen
    End If
End Sub

Private Sub HandleGDecline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim n As Byte

    n = Player(Index).GuildInviter
    If n > 0 Then
         If Player(Index).GuildInvite > 0 Then
            SendPlayerMsg Index, "[Guild Info] You have declined the guild invition.", BrightGreen
            SendPlayerMsg n, "[Guild Info] " & Current_Name(Index) & " has declined the guild invition.", BrightGreen
            Player(Index).GuildInvite = 0
            Player(Index).GuildInviter = 0
        Else
            SendPlayerMsg Index, "[Guild Info] You have not been invited to a guild.", BrightGreen
        End If
    Else
        SendPlayerMsg Index, "[Guild Info] You have not been invited to a guild.", BrightGreen
    End If
End Sub

Private Sub HandleGuildMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String, s As String
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
        
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
            HackingAttempt Index, "Broadcast Text Modification"
            Exit Sub
        End If
    Next
    
    If Current_Guild(Index) > 0 Then
        s = "[Guild] " & Current_Name(Index) & ": " & Msg
        AddLog s, PLAYER_LOG
        SendGuildMsg Current_Guild(Index), s, BrightGreen
    Else
        SendPlayerMsg Index, "[Guild Info] You are not in a guild.", BrightGreen
    End If
End Sub

Private Sub HandlePartyMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Msg As String, s As String
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Msg = Buffer.ReadString
        
    ' Prevent hacking
    For i = 1 To Len(Msg)
        If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then
            HackingAttempt Index, "Party Text Modification"
            Exit Sub
        End If
    Next
    
    If Player(Index).InParty Then
        s = "[Party] " & Current_Name(Index) & ": " & Msg
        AddLog s, PLAYER_LOG
        SendPartyMsg Player(Index).PartyIndex, s, BrightBlue
    Else
        SendPlayerMsg Index, "[Message Info] You are not in a party.", Pink
    End If
End Sub

Private Sub HandleKill(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim Name As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    
    i = FindPlayer(Name)
    If i <= 0 Then
        SendActionMsg Current_Map(Index), "Player not found", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
        Exit Sub
    End If
        
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_MONITER Then
        ' If they try to hack they kill themselves - muhaha
        SendGlobalMsg "[OMG] " & Current_Name(Index) & " stops to smell a flower. Out of nowhere a sword cuts their fucking head off. " & Current_Name(Index) & " has been Owned.", ActionColor
        
        OnDeath Index
        Exit Sub
    End If
    
    ' Don't kill yourself or and admin
    If Current_Name(i) = Current_Name(Index) Or Current_Access(i) > ADMIN_MONITER Then
        Exit Sub
    End If
            
    ' Player is dead
    SendGlobalMsg "[OMG] " & Current_Name(i) & " mysteriously drops dead.", ActionColor
    
    OnDeath i
End Sub

Private Sub HandleSetBound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim MapNum As Long
Dim MapNpcNum As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MapNpcNum = Buffer.ReadLong
    i = Buffer.ReadLong          ' Shopnum
    
    MapNum = Current_Map(Index)
    
    ' Check to make sure they are still in range of npc shop
    If MapData(MapNum).MapNpc(MapNpcNum).Num <= 0 Then Exit Sub
    
    ' Make sure it's a shop npc
    If Npc(MapData(MapNum).MapNpc(MapNpcNum).Num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then Exit Sub
    
    ' another check for the same shop - we don't like hacking :)
    If Npc(MapData(MapNum).MapNpc(MapNpcNum).Num).Stat(Stats.Wisdom) <> i Then Exit Sub
    
    ' Check if you're in range
    If Not NpcInRange(Index, MapNpcNum, 2) Then
        SendActionMsg MapNum, "You are too far away to do that.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
        Exit Sub
    End If
    
    ' Will set the homepoint
    Update_Bound Index, Shop(i).BindPoint
    SendPlayerMsg Index, "Your homepoint has been set.", ActionColor
End Sub

Private Sub HandleCancelSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    CancelCastSpell Index
End Sub

Private Sub HandleRelease(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Make sure you're dead
    If Not Current_IsDead(Index) Then Exit Sub
    
    OnRelease Index
End Sub

Private Sub HandleRevive(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Make sure you're dead
    If Not Current_IsDead(Index) Then Exit Sub
    
    ' Check if someone activated the revivable flag
    If Player(Index).Revivable = 0 Then Exit Sub

    OnRevive Index
End Sub

'
' Quest Packets
'
Private Sub HandleRequestEditQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    SendQuestEditor Index
End Sub

Private Sub HandleEditQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim QuestNum As Long
    
    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' The Quest #
    QuestNum = Buffer.ReadLong
    
    ' Prevent hacking
    If QuestNum < 0 Or QuestNum > MAX_QUESTS Then
        HackingAttempt Index, "Invalid Quest Index"
        Exit Sub
    End If
    
    AddLog Current_Name(Index) & " editing Quest #" & QuestNum & ".", ADMIN_LOG
    SendEditQuestTo Index, QuestNum
End Sub

Private Sub HandleSaveQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim QuestNum As Long

    ' Prevent hacking
    If Current_Access(Index) < ADMIN_DEVELOPER Then
        HackingAttempt Index, "Admin Cloning"
        Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    QuestNum = Buffer.ReadLong
    
    ' Check if a valid Quest
    If QuestNum <= 0 Then Exit Sub
    If QuestNum > MAX_QUESTS Then Exit Sub
    
    ' Check the length of the buffer, make sure it's the same as the size needed
    If Buffer.Length <> QuestSize Then Exit Sub
    
    ' Set the Quest from the byte array
    Set_QuestData QuestNum, Buffer.ReadBytes(QuestSize)
    
    ' Save it
    SendUpdateQuestToAll QuestNum
    SaveQuest QuestNum
    
    ' TODO: Check to see if the saved Quest will affect any players
    For i = 1 To OnlinePlayersCount
        CheckPlayerQuests OnlinePlayers(i)
    Next
    
    AddLog Current_Name(Index) & " saved Quest #" & QuestNum & ".", ADMIN_LOG
    
    ' Update the Npcs Quest List
    Update_Npcs_Quest Quest(QuestNum).StartNPC
    Update_Npcs_Quest Quest(QuestNum).EndNPC
    
    ' Update the cache now
    CacheQuests
End Sub

Private Sub HandleAcceptQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim MapNum As Long
Dim NpcNum As Long
Dim QuestNum As Long
Dim QuestMapNpcNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    QuestNum = Buffer.ReadLong
    QuestMapNpcNum = Buffer.ReadLong
    
    ' Prevent hacking
    If (QuestNum <= 0) Or (QuestNum > MAX_QUESTS) Then
        HackingAttempt Index, "Accept Quest Modification"
        Exit Sub
    End If
    
    MapNum = Current_Map(Index)
    
     ' Prevent hacking
    If QuestMapNpcNum <= 0 Then Exit Sub
    If QuestMapNpcNum > MapData(MapNum).NpcCount Then Exit Sub
    
    NpcNum = MapData(MapNum).MapNpc(QuestMapNpcNum).Num
    
    If NpcNum < 0 Then Exit Sub
    If NpcNum > MAX_NPCS Then Exit Sub
    
    ' Make sure it's a quest giver
    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_QUEST Then Exit Sub
    
    ' Make sure they have the quest
    If Quest(QuestNum).StartNPC <> NpcNum Then Exit Sub

    ' Make sure you're in range
    If Not NpcInRange(Index, QuestMapNpcNum, 2) Then Exit Sub
    
    OnQuestAccept Index, QuestNum
End Sub

Private Sub HandleCompleteQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim MapNum As Long
Dim NpcNum As Long
Dim QuestProgressNum As Long
Dim QuestMapNpcNum As Long
Dim SelectReward As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    QuestProgressNum = Buffer.ReadLong
    QuestMapNpcNum = Buffer.ReadLong
    SelectReward = Buffer.ReadLong
    
    ' Prevent hacking
    If (QuestProgressNum <= 0) Or (QuestProgressNum > MAX_PLAYER_QUESTS) Then
        HackingAttempt Index, "Accept Quest Modification"
        Exit Sub
    End If
    
    MapNum = Current_Map(Index)
    
     ' Prevent hacking
    If QuestMapNpcNum <= 0 Then Exit Sub
    If QuestMapNpcNum > MapData(MapNum).NpcCount Then Exit Sub
    
    NpcNum = MapData(MapNum).MapNpc(QuestMapNpcNum).Num
    
    If NpcNum < 0 Then Exit Sub
    If NpcNum > MAX_NPCS Then Exit Sub
    
    ' Make sure it's a quest giver
    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_QUEST Then Exit Sub
    
    ' Make sure they have the quest
    If Quest(Player(Index).Char.QuestProgress(QuestProgressNum).QuestNum).EndNPC <> NpcNum Then Exit Sub

    ' Make sure you're in range
    If Not NpcInRange(Index, QuestMapNpcNum, 2) Then Exit Sub
    
    OnQuestTurnIn Index, QuestProgressNum, SelectReward
End Sub

Private Sub HandleDropQuest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim MapNum As Long
Dim NpcNum As Long
Dim QuestProgressNum As Long
Dim QuestMapNpcNum As Long
Dim SelectReward As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    QuestProgressNum = Buffer.ReadLong
    
    ' Prevent hacking
    If (QuestProgressNum <= 0) Or (QuestProgressNum > MAX_PLAYER_QUESTS) Then
        HackingAttempt Index, "Drop Quest Modification"
        Exit Sub
    End If
        
    OnQuestDrop Index, QuestProgressNum
End Sub


