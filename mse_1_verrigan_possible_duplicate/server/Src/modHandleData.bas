Attribute VB_Name = "modHandleData"
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/??/2005  Verrigan   Optimized module to handle packets more
'*                        efficiently.
'****************************************************************
Option Explicit
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/??/2005  Verrigan   Added GetAddress() to get the address
'*                        a function. Pass AddressOf <Function>.
'****************************************************************
Public Function GetAddress(FunAddr As Long) As Long
  GetAddress = FunAddr
End Function
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/??/2005  Verrigan   Added InitMessages() to initialize the
'*                        HandleDataSub() array with the
'*                        addresses of each packet function.
'****************************************************************
Public Sub InitMessages()
  HandleDataSub(SMsgGetClasses) = GetAddress(AddressOf HandleGetClasses)
  HandleDataSub(SMsgNewAccount) = GetAddress(AddressOf HandleNewAccount)
  HandleDataSub(SMsgDelAccount) = GetAddress(AddressOf HandleDelAccount)
  HandleDataSub(SMsgLogin) = GetAddress(AddressOf HandleLogin)
  HandleDataSub(SMsgAddChar) = GetAddress(AddressOf HandleAddChar)
  HandleDataSub(SMsgDelChar) = GetAddress(AddressOf HandleDelChar)
  HandleDataSub(SMsgUseChar) = GetAddress(AddressOf HandleUseChar)
  HandleDataSub(SMsgSay) = GetAddress(AddressOf HandleSay)
  HandleDataSub(SMsgEmote) = GetAddress(AddressOf HandleEmote)
  HandleDataSub(SMsgBroadcast) = GetAddress(AddressOf HandleBroadcast)
  HandleDataSub(SMsgGlobal) = GetAddress(AddressOf HandleGlobal)
  HandleDataSub(SMsgAdmin) = GetAddress(AddressOf HandleAdmin)
  HandleDataSub(SMsgPlayer) = GetAddress(AddressOf HandlePlayer)
  HandleDataSub(SMsgPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
  HandleDataSub(SMsgPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
  HandleDataSub(SMsgUseItem) = GetAddress(AddressOf HandleUseItem)
  HandleDataSub(SMsgAttack) = GetAddress(AddressOf HandleAttack)
  HandleDataSub(SMsgUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
  HandleDataSub(SMsgPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
  HandleDataSub(SMsgWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
  HandleDataSub(SMsgWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
  HandleDataSub(SMsgWarpTo) = GetAddress(AddressOf HandleWarpTo)
  HandleDataSub(SMsgSetSprite) = GetAddress(AddressOf HandleSetSprite)
  HandleDataSub(SMsgGetStats) = GetAddress(AddressOf HandleGetStats)
  HandleDataSub(SMsgRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
  HandleDataSub(SMsgMapData) = GetAddress(AddressOf HandleMapData)
  HandleDataSub(SMsgNeedMap) = GetAddress(AddressOf HandleNeedMap)
  HandleDataSub(SMsgMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
  HandleDataSub(SMsgMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
  HandleDataSub(SMsgMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
  HandleDataSub(SMsgMapReport) = GetAddress(AddressOf HandleMapReport)
  HandleDataSub(SMsgKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
  HandleDataSub(SMsgBanList) = GetAddress(AddressOf HandleBanList)
  HandleDataSub(SMsgBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
  HandleDataSub(SMsgBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
  HandleDataSub(SMsgRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
  HandleDataSub(SMsgRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
  HandleDataSub(SMsgEditItem) = GetAddress(AddressOf HandleEditItem)
  HandleDataSub(SMsgSaveItem) = GetAddress(AddressOf HandleSaveItem)
  HandleDataSub(SMsgRequestEditNPC) = GetAddress(AddressOf HandleRequestEditNPC)
  HandleDataSub(SMsgEditNPC) = GetAddress(AddressOf HandleEditNPC)
  HandleDataSub(SMsgSaveNPC) = GetAddress(AddressOf HandleSaveNPC)
  HandleDataSub(SMsgRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
  HandleDataSub(SMsgEditShop) = GetAddress(AddressOf HandleEditShop)
  HandleDataSub(SMsgSaveShop) = GetAddress(AddressOf HandleSaveShop)
  HandleDataSub(SMsgRequestEditSpell) = GetAddress(AddressOf HandleRequestEditSpell)
  HandleDataSub(SMsgEditSpell) = GetAddress(AddressOf HandleEditSpell)
  HandleDataSub(SMsgSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
  HandleDataSub(SMsgSetAccess) = GetAddress(AddressOf HandleSetAccess)
  HandleDataSub(SMsgWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
  HandleDataSub(SMsgSetMOTD) = GetAddress(AddressOf HandleSetMOTD)
  HandleDataSub(SMsgTrade) = GetAddress(AddressOf HandleTrade)
  HandleDataSub(SMsgTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
  HandleDataSub(SMsgFixItem) = GetAddress(AddressOf HandleFixItem)
  HandleDataSub(SMsgSearch) = GetAddress(AddressOf HandleSearch)
  HandleDataSub(SMsgParty) = GetAddress(AddressOf HandleParty)
  HandleDataSub(SMsgJoinParty) = GetAddress(AddressOf HandleJoinParty)
  HandleDataSub(SMsgLeaveParty) = GetAddress(AddressOf HandleLeaveParty)
  HandleDataSub(SMsgSpells) = GetAddress(AddressOf HandleSpells)
  HandleDataSub(SMsgCast) = GetAddress(AddressOf HandleCast)
  HandleDataSub(SMsgRequestLocation) = GetAddress(AddressOf HandleRequestLocation)
End Sub
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/??/2005  Verrigan   Separated the handling of each packet
'*                        into their own subroutines. They are
'*                        still 'basically' handled the same as
'*                        before. :)
'****************************************************************
'Each Handle Sub must be declared with 4 longs.. This is a requirement of
'CallWindowProc. We use these longs to pass the Index of the connection,
'the starting byte address (the 1st byte after the message type), and the
'byte length of the rest of the packet.
Private Sub HandleGetClasses(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  If Not IsPlaying(Index) Then Call SendNewCharClasses(Index)
End Sub
Private Function HasAccess(ByVal Index As Long, ByVal aLevel As Long) As Byte
  HasAccess = 0
  If GetPlayerAccess(Index) < aLevel Then
    Call HackingAttempt(Index, "Admin Cloning")
    Exit Function
  End If
  HasAccess = 1
End Function
Private Function CheckAccountInfo(ByVal Index As Long, ByRef Buffer() As Byte, ByRef Name As String, ByRef Password As String) As Byte
  Dim nBytes() As Byte, pBytes() As Byte
  Dim nLen As Integer, pLen As Integer
  Dim i As Long, n As Byte
  Dim cHV As Integer, cMV As Integer, cLV As Integer
  
  CheckAccountInfo = 0
  If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
    'Get the data.
    nLen = GetIntegerFromBuffer(Buffer, True)
    nBytes = GetFromBuffer(Buffer, nLen, True)
    pLen = GetIntegerFromBuffer(Buffer, True)
    pBytes = GetFromBuffer(Buffer, pLen, True)
    cHV = GetIntegerFromBuffer(Buffer, True)
    cMV = GetIntegerFromBuffer(Buffer, True)
    cLV = GetIntegerFromBuffer(Buffer, True)
    
    'Check versions
    If cHV < CLIENT_MAJOR Or cMV < CLIENT_MINOR Or cLV < CLIENT_REVISION Then
      Call AlertMsg(Index, "Version outdated, please visit " & GAME_WEBSITE)
      Exit Function
    End If
    
    'Prevent hacking
    If nLen < 3 Or pLen < 3 Then
      Call AlertMsg(Index, "Your name and password must be at least three characters in length")
      Exit Function
    End If
    
    For i = 0 To UBound(nBytes)
      n = nBytes(i)
      
      If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
        'All is well. :)
      Else
        Call AlertMsg(Index, "Invalid name. Only letters, numbers, spaces, and _ allowed in names.")
        Exit Function
      End If
    Next
    
    'Everything went well. Convert byte arrays to unicode strings.
    Name = StrConv(nBytes, vbUnicode)
    Password = StrConv(pBytes, vbUnicode)
    CheckAccountInfo = 1
  End If
End Function
Private Sub HandleNewAccount(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Name As String, Password As String
  Dim Buffer() As Byte
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If CheckAccountInfo(Index, Buffer, Name, Password) = 1 Then
    'Check to see if account already exists.
    If Not AccountExist(Name) Then
      Call AddAccount(Index, Name, Password)
      Call TextAdd(frmServer.txtText, "Account " & Name & " has been created.", True)
      Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
      Call AlertMsg(Index, "Your account has been created!")
    Else
      Call AlertMsg(Index, "Sorry, but that account name is already taken!")
    End If
  End If
End Sub
'*******************************************************************
'* CheckAccountPassword() checks for account and matches password. *
'*******************************************************************
Private Function CheckAccountPassword(Index As Long, Name As String, Password As String) As Byte
  CheckAccountPassword = 0
  If Not AccountExist(Name) Or Not PasswordOK(Name, Password) Then
    Call AlertMsg(Index, "You have specified an invalid account or password.")
    Exit Function
  End If
  CheckAccountPassword = 1
End Function
Private Sub HandleDelAccount(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim Name As String, Password As String
  Dim i As Byte
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If CheckAccountInfo(Index, Buffer, Name, Password) = 1 Then
    If CheckAccountPassword(Index, Name, Password) = 0 Then Exit Sub
    
    'Delete names from master name file
    Call LoadPlayer(Index, Name)
    For i = 1 To MAX_CHARS
      If Len(Trim(Player(Index).Char(i).Name)) > 0 Then
        Call DeleteName(Player(Index).Char(i).Name)
      End If
    Next
    Call ClearPlayer(Index)
    
    'Everything went ok
    Call Kill(App.Path & "\accounts\" & Trim(Name) & ".ini")
    Call AddLog("Account " & Trim(Name) & " has been deleted.", PLAYER_LOG)
    Call AlertMsg(Index, "Your account has been deleted.")
  End If
End Sub
Private Sub HandleLogin(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Name As String, Password As String
  Dim Buffer() As Byte
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If CheckAccountInfo(Index, Buffer, Name, Password) = 1 Then
    If CheckAccountPassword(Index, Name, Password) = 0 Then Exit Sub
    
    If IsMultiAccounts(Name) Then
      Call AlertMsg(Index, "Multiple account logins are not authorized.")
      Exit Sub
    End If
    
    'Everything went ok. Load the player.
    Call LoadPlayer(Index, Name)
    Call SendChars(Index)
    
    'Show the player up on the socket status.
    Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
    Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", True)
  End If
End Sub
Private Sub HandleAddChar(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim nBytes() As Byte
  Dim nLen As Byte
  Dim Name As String
  Dim Sex As Byte, Class As Byte, CharNum As Byte
  Dim i As Long, n As Byte
  Dim Buffer() As Byte
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If Not IsPlaying(Index) Then
    nLen = GetIntegerFromBuffer(Buffer, True)
    nBytes = GetFromBuffer(Buffer, nLen, True)
    Sex = GetByteFromBuffer(Buffer, True)
    Class = GetByteFromBuffer(Buffer, True)
    CharNum = GetByteFromBuffer(Buffer, True)
    
    'Who really cares about the name consty? (Removed check to speed up the procedure.)
    
    'Prevent hacking
    If nLen < 3 Then
      Call AlertMsg(Index, "Character name must be at least three characters in length.")
      Exit Sub
    End If
    
    For i = 0 To UBound(nBytes)
      n = nBytes(i)
      
      If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
        'All is well. :)
      Else
        Call AlertMsg(Index, "Invalid name. Only letters, numbers, spaces, and _ allowed in names.")
        Exit Sub
      End If
    Next
    
    If CharNum < 1 Or CharNum > MAX_CHARS Then
      Call HackingAttempt(Index, "Invalid CharNum")
      Exit Sub
    End If
    
    If Sex < SEX_MALE Or Sex > SEX_FEMALE Then
      Call HackingAttempt(Index, "Invalid Sex")
      Exit Sub
    End If
    
    If Class < 0 Or Class > Max_Classes Then
      Call HackingAttempt(Index, "Invalid Class")
      Exit Sub
    End If
    
    'Check if char already exists in slot.
    If CharExist(Index, CharNum) Then
      Call AlertMsg(Index, "Slot already in use. Cannot create char here!")
      Exit Sub
    End If
    
    'Convert byte array to unicode string.
    Name = StrConv(nBytes, vbUnicode)
    
    'Check if name is already in use.
    If FindChar(Name) Then
      Call AlertMsg(Index, "Sorry, but that name is already in use!")
      Exit Sub
    End If
    
    'Everything went ok. Add the character.
    Call AddChar(Index, Name, Sex, Class, CharNum)
    Call SavePlayer(Index)
    Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
    Call AlertMsg(Index, "Character has been created!")
  End If
End Sub
Private Sub HandleDelChar(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim CharNum As Byte
  Dim Buffer() As Byte
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If Not IsPlaying(Index) Then
    CharNum = GetByteFromBuffer(Buffer, True)
    
    'Prevent hacking
    If CharNum < 1 Or CharNum > MAX_CHARS Then
      Call HackingAttempt(Index, "Invalid CharNum")
      Exit Sub
    End If
    
    Call DelChar(Index, CharNum)
    Call AddLog("Character deleted on " & GetPlayerLogin(Index) & "'s account", PLAYER_LOG)
    Call AlertMsg(Index, "Character has been deleted!")
  End If
End Sub
Private Sub HandleUseChar(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim CharNum As Byte
  Dim f As Long
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If Not IsPlaying(Index) Then
    CharNum = GetByteFromBuffer(Buffer, True)
    
    'Prevent hacking
    If CharNum < 1 Or CharNum > MAX_CHARS Then
      Call HackingAttempt(Index, "Invalid CharNum")
      Exit Sub
    End If
    
    'Check to make sure the character exists and if so, set it as its current char
    If CharExist(Index, CharNum) Then
      Player(Index).CharNum = CharNum
      Call JoinGame(Index)
      
      'Why did we ever do CharNum = Player(Index).CharNum here? They were always the same.
      Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
      Call TextAdd(frmServer.txtText, GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", True)
      Call UpdateCaption
      
      'Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
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
End Sub
Private Function ValidateMessage(Index As Long, ByRef Buffer() As Byte) As Byte
  Dim i As Long
  
  ValidateMessage = 0
  For i = 2 To UBound(Buffer)
    If Buffer(i) < 32 Or Buffer(i) > 126 Then
      Call HackingAttempt(Index, "Message Text Modification")
      Exit Function
    End If
  Next i
  ValidateMessage = 1
End Function
Private Sub HandleSay(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim Msg As String
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If ValidateMessage(Index, Buffer) = 1 Then
    Msg = GetPlayerName(Index) & " says, '" & GetStringFromBuffer(Buffer, True) & "'"
    
    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), Msg, SayColor)
  End If
End Sub
Private Sub HandleEmote(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim Msg As String
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If ValidateMessage(Index, Buffer) = 1 Then
    Msg = GetPlayerName(Index) & " " & GetStringFromBuffer(Buffer, True)
    
    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), Msg, EmoteColor)
  End If
End Sub
Private Sub HandleBroadcast(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim Msg As String
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If ValidateMessage(Index, Buffer) = 1 Then
    Msg = GetPlayerName(Index) & ": " & GetStringFromBuffer(Buffer, True)
    
    Call AddLog(Msg, PLAYER_LOG)
    Call GlobalMsg(Msg, BroadcastColor)
    Call TextAdd(frmServer.txtText, Msg, True)
  End If
End Sub
Private Sub HandleGlobal(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim Msg As String
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If GetPlayerAccess(Index) > 0 And ValidateMessage(Index, Buffer) = 1 Then
    Msg = GetPlayerName(Index) & ": " & Msg
    
    Call AddLog(Msg, ADMIN_LOG)
    Call GlobalMsg(Msg, GlobalColor)
    Call TextAdd(frmServer.txtText, Msg, True)
  End If
End Sub
Private Sub HandleAdmin(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim Msg As String
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If GetPlayerAccess(Index) > 0 And ValidateMessage(Index, Buffer) = 1 Then
    Msg = GetPlayerName(Index) & ") " & GetStringFromBuffer(Buffer, True)
    
    Call AddLog("(admin " & Msg, ADMIN_LOG)
    Call AdminMsg("(admin " & Msg, AdminColor)
  End If
End Sub
Private Sub HandlePlayer(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim MsgTo As Long
  Dim Msg As String
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  If ValidateMessage(Index, Buffer) = 1 Then
    MsgTo = FindPlayer(GetStringFromBuffer(Buffer, True))
    
    If MsgTo <> Index Then
      If MsgTo > 0 Then
        Msg = GetStringFromBuffer(Buffer, True)
        
        Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", '" & Msg & "'", PLAYER_LOG)
        Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & Msg & "'", TellColor)
        Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
      Else
        Call PlayerMsg(Index, "Player is not online.", White)
      End If
    Else
      Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " begins to mumble to himself, what a weirdo...", PLAYER_LOG)
      Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " begins to mumble to himself, what a weirdo...", Green)
    End If
  End If
End Sub
Private Function CheckDirection(Index As Long, nDir As Byte) As Byte
  CheckDirection = 0
  
  If nDir < DIR_UP Or nDir > DIR_RIGHT Then
    Call HackingAttempt(Index, "Invalid Direction")
    Exit Function
  End If
  
  CheckDirection = 1
End Function
Private Sub HandlePlayerMove(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim nDir As Byte
  Dim Movement As Byte
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  nDir = GetByteFromBuffer(Buffer, True)
  Movement = GetByteFromBuffer(Buffer, True)
  
  'Prevent hacking
  If CheckDirection(Index, nDir) = 0 Then Exit Sub
  
  If Movement < 1 Or Movement > 2 Then
    Call HackingAttempt(Index, "Invalid Movement")
    Exit Sub
  End If
  
  'Prevent player from moving if they have casted a spell.
  If Player(Index).CastedSpell = YES Then
    'Check if they have already casted a spell, and if so, we can't let them move
    If GetTickCount > Player(Index).AttackTimer + 1000 Then
      Player(Index).CastedSpell = NO
    Else
      Call SendPlayerXY(Index)
      Exit Sub
    End If
  End If
  
  Call PlayerMove(Index, nDir, Movement)
End Sub
Private Sub HandlePlayerDir(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim nDir As Byte
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  nDir = GetByteFromBuffer(Buffer, True)
  
  'Prevent hacking
  If CheckDirection(Index, nDir) = 0 Then Exit Sub
  
  Call SetPlayerDir(Index, nDir)
  Call SendDataToMapBut(Index, GetPlayerMap(Index), "PLAYERDIR" & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & END_CHAR)
End Sub
Private Function CheckInvItem(Index As Long, InvNum As Byte) As Byte
  CheckInvItem = 0
  If InvNum < 1 Or InvNum > MAX_INV Then
    Call HackingAttempt(Index, "Invalid InvNum")
    Exit Function
  End If
  CheckInvItem = 1
End Function
Private Sub HandleUseItem(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim InvNum As Byte
  Dim CharNum As Byte
  Dim i As Long, n As Long
  Dim x As Long, y As Long
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  InvNum = GetByteFromBuffer(Buffer, True)
  CharNum = Player(Index).CharNum
  
  ' Prevent hacking
  If CheckInvItem(Index, InvNum) = 0 Then Exit Sub
  
  If CharNum < 1 Or CharNum > MAX_CHARS Then
    Call HackingAttempt(Index, "Invalid CharNum")
    Exit Sub
  End If
  
  'So far, all is well
  If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
    n = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
    
    ' Find out what kind of item it is
    Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
    Case ITEM_TYPE_ARMOR
      If InvNum <> GetPlayerArmorSlot(Index) Then
        If Int(GetPlayerDEF(Index)) < n Then
          Call PlayerMsg(Index, "Your defense is to low to wear this armor!  Required DEF (" & n * 2 & ")", BrightRed)
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
          Call PlayerMsg(Index, "Your strength is to low to hold this weapon!  Required STR (" & n * 2 & ")", BrightRed)
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
          Call PlayerMsg(Index, "Your speed coordination is to low to wear this helmet!  Required SPEED (" & n * 2 & ")", BrightRed)
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
      If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY Then
        ' Check if the key they are using matches the map key
        If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(x, y).Data1 Then
          TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
          TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                    
          Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
          Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
                   
          ' Check if we are supposed to take away the item
          If Map(GetPlayerMap(Index)).Tile(x, y).Data2 = 1 Then
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
    End Select
  End If
End Sub
Private Sub HandleAttack(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim i As Long, n As Long, Damage As Long
  
  ' Try to attack a player
  For i = 1 To MAX_PLAYERS
    ' Make sure we dont try to attack ourselves
    If i <> Index Then
      ' Can we attack the player?
      If CanAttackPlayer(Index, i) Then
        If Not CanPlayerBlockHit(i) Then
          ' Get the damage we can do
          If Not CanPlayerCriticalHit(Index) Then
            Damage = GetPlayerDamage(Index) - GetPlayerProtection(i)
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
          Call PlayerMsg(Index, GetPlayerName(i) & "'s " & Trim(StrConv(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).ItemName, vbUnicode)) & " has blocked your hit!", BrightCyan)
          Call PlayerMsg(i, "Your " & Trim(StrConv(Item(GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))).ItemName, vbUnicode)) & " has blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan)
        End If
        
        Exit Sub
      End If
    End If
  Next i
  
  ' Try to attack a npc
  For i = 1 To MAX_MAP_NPCS
    ' Can we attack the npc?
    If CanAttackNpc(Index, i) Then
      ' Get the damage we can do
      If Not CanPlayerCriticalHit(Index) Then
        Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
      Else
        n = GetPlayerDamage(Index)
        Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
        Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
      End If
      
      If Damage > 0 Then
        Call AttackNpc(Index, i, Damage)
      Else
        Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
      End If
      Exit Sub
    End If
  Next i
End Sub
Private Sub HandleUseStatPoint(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim PointType As Byte
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  PointType = GetByteFromBuffer(Buffer, True)
  
  'Prevent hacking
  If PointType < 0 Or PointType > 3 Then
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
End Sub
Private Sub HandlePlayerInfoRequest(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim Name As String
  Dim i As Long, n As Long
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  Name = GetStringFromBuffer(Buffer, True)
  i = FindPlayer(Name)
  
  If i > 0 Then
    Call PlayerMsg(Index, "Account: " & Trim(Player(i).Login) & ", Name: " & GetPlayerName(i), BrightGreen)
    If GetPlayerAccess(Index) > ADMIN_MONITER Then
      Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(i) & " -=-", BrightGreen)
      Call PlayerMsg(Index, "Level: " & GetPlayerLevel(i) & "  Exp: " & GetPlayerExp(i) & "/" & GetPlayerNextLevel(i), BrightGreen)
      Call PlayerMsg(Index, "HP: " & GetPlayerHP(i) & "/" & GetPlayerMaxHP(i) & "  MP: " & GetPlayerMP(i) & "/" & GetPlayerMaxMP(i) & "  SP: " & GetPlayerSP(i) & "/" & GetPlayerMaxSP(i), BrightGreen)
      Call PlayerMsg(Index, "STR: " & GetPlayerSTR(i) & "  DEF: " & GetPlayerDEF(i) & "  MAGI: " & GetPlayerMAGI(i) & "  SPEED: " & GetPlayerSPEED(i), BrightGreen)
      n = Int(GetPlayerSTR(i) / 2) + Int(GetPlayerLevel(i) / 2)
      i = Int(GetPlayerDEF(i) / 2) + Int(GetPlayerLevel(i) / 2)
      If n > 100 Then n = 100
      If i > 100 Then i = 100
      Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", BrightGreen)
    End If
  Else
    Call PlayerMsg(Index, "Player is not online.", White)
  End If
End Sub
Private Sub HandleWarpMeTo(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Long
  
  'Prevent hacking
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  n = FindPlayer(GetStringFromBuffer(Buffer, True))
  
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
Private Sub HandleWarpToMe(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Long
  
  'Prevent hacking
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  n = FindPlayer(GetStringFromBuffer(Buffer, True))

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
Private Sub HandleWarpTo(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Integer
  
  'Prevent hacking
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  'The map
  n = GetIntegerFromBuffer(Buffer, True)
  
  ' Prevent hacking
  If n < 0 Or n > MAX_MAPS Then
    Call HackingAttempt(Index, "Invalid map")
    Exit Sub
  End If
  
  Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
  Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
  Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub
Private Sub HandleSetSprite(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Integer
  
  'Prevent hacking
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  'The sprite
  n = GetIntegerFromBuffer(Buffer, True)

  Call SetPlayerSprite(Index, n)
  Call SendPlayerData(Index)
End Sub
Private Sub HandleGetStats(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim i As Long, n As Long
  
  Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(Index) & " -=-", White)
  Call PlayerMsg(Index, "Level: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), White)
  Call PlayerMsg(Index, "HP: " & GetPlayerHP(Index) & "/" & GetPlayerMaxHP(Index) & "  MP: " & GetPlayerMP(Index) & "/" & GetPlayerMaxMP(Index) & "  SP: " & GetPlayerSP(Index) & "/" & GetPlayerMaxSP(Index), White)
  Call PlayerMsg(Index, "STR: " & GetPlayerSTR(Index) & "  DEF: " & GetPlayerDEF(Index) & "  MAGI: " & GetPlayerMAGI(Index) & "  SPEED: " & GetPlayerSPEED(Index), White)
  
  n = Int(GetPlayerSTR(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
  i = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)
  
  If n > 100 Then n = 100
  If i > 100 Then i = 100
  
  Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", White)
End Sub
Private Sub HandleRequestNewMap(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim nDir As Byte
  
  'Prevent hacking
  If CheckDirection(Index, nDir) = 0 Then Exit Sub
  
  Call PlayerMove(Index, Dir, 1)
End Sub
Private Sub HandleMapData(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim i As Long, MapNum As Integer
  Dim nMap As MapRec
  
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  MapNum = GetPlayerMap(Index)
  
  'MsgBox Len(Map(MapNum))
  Call CopyMemory(nMap, Buffer(0), aLen(Buffer))
  Map(MapNum) = nMap
  
  Call SendMapNpcsToMap(MapNum)
  Call SpawnMapNpcs(MapNum)
  
  'Save the map
  Call SaveMap(MapNum)
  
  'Refresh map for everyone online
  For i = 1 To MAX_PLAYERS
    If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
      Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
    End If
  Next
End Sub
Private Sub HandleNeedMap(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim nm As Byte
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  nm = GetByteFromBuffer(Buffer, True)
  
  If nm = YES Then
    Call SendMap(Index, GetPlayerMap(Index))
  End If
  Call SendMapItemsTo(Index, GetPlayerMap(Index))
  Call SendMapNpcsTo(Index, GetPlayerMap(Index))
  Call SendJoinMap(Index)
  Player(Index).GettingMap = NO
  Call SendDataTo(Index, "MAPDONE" & SEP_CHAR & END_CHAR)
End Sub
Private Sub HandleMapGetItem(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Call PlayerMapGetItem(Index)
End Sub
Private Sub HandleMapDropItem(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim InvNum As Byte, Ammount As Long
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  InvNum = GetByteFromBuffer(Buffer, True)
  Ammount = GetLongFromBuffer(Buffer, True)
  
  'Prevent hacking
  If CheckInvItem(Index, InvNum) = 0 Then Exit Sub
  
  If Ammount > GetPlayerInvItemValue(Index, InvNum) Then
    Call HackingAttempt(Index, "Item ammount modification")
    Exit Sub
  End If
  
  If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
    'Check if money and, if it is, we want to make sure that they aren't trying to drop 0 value
    If Ammount <= 0 Then
      Call HackingAttempt(Index, "Trying to drop 0 ammount of currency")
      Exit Sub
    End If
  End If
  
  'No hacking attempt.. Let's drop the item.
  Call PlayerMapDropItem(Index, InvNum, Ammount)
End Sub
Private Sub HandleMapRespawn(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim i As Long
  
  'Prevent hacking
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub
  
  'Clear it all out
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
End Sub
Private Sub HandleMapReport(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim s As String
  Dim tMapStart As Long, tMapEnd As Long, i As Long
  
  'Prevent hacking
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub
  
  s = "Free Maps: "
  tMapStart = 1
  tMapEnd = 1
  
  For i = 1 To MAX_MAPS
    If Trim(StrConv(Map(i).MapName, vbUnicode)) = "" Then
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
  
  Call PlayerMsg(Index, s, Brown)
End Sub
Private Sub HandleKickPlayer(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Long
  
  'Prevent hacking
  If HasAccess(Index, 1) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  'The player index
  n = FindPlayer(GetStringFromBuffer(Buffer, True))
  
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
End Sub
Private Sub HandleBanList(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim f As Long, n As Long
  Dim s As String, Name As String
  
  'Prevent hacking
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub

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
End Sub
Private Sub HandleBanDestroy(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  ' Prevent hacking
  If HasAccess(Index, ADMIN_CREATOR) = 0 Then Exit Sub
  
  Call Kill(App.Path & "\data\banlist.txt")
  Call PlayerMsg(Index, "Ban list destroyed.", White)
End Sub
Private Sub HandleBanPlayer(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Long
  
  ' Prevent hacking
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  ' The player index
  n = FindPlayer(GetStringFromBuffer(Buffer, True))
  
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
End Sub
Private Sub HandleRequestEditMap(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  ' Prevent hacking
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub
  
  Call SendDataTo(Index, "EDITMAP" & SEP_CHAR & END_CHAR)
End Sub
Private Sub HandleRequestEditItem(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  ' Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
  
  Call SendDataTo(Index, "ITEMEDITOR" & SEP_CHAR & END_CHAR)
End Sub
Private Sub HandleEditItem(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Integer
  
  ' Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  ' The item #
  n = GetIntegerFromBuffer(Buffer, True)
  
  ' Prevent hacking
  If n < 0 Or n > MAX_ITEMS Then
      Call HackingAttempt(Index, "Invalid Item Index")
      Exit Sub
  End If
  
  Call AddLog(GetPlayerName(Index) & " editing item #" & n & ".", ADMIN_LOG)
  Call SendEditItemTo(Index, n)
End Sub
Private Sub HandleSaveItem(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Integer
  Dim nMap As MapRec
  
  ' Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  n = GetIntegerFromBuffer(Buffer, True)
  
  If n < 0 Or n > MAX_ITEMS Then
      Call HackingAttempt(Index, "Invalid Item Index")
      Exit Sub
  End If
  
  ' Update the item
  Call CopyMemory(Item(n), Buffer(0), aLen(Buffer))
  
  'Item(n).Name = Parse(2)
  'Item(n).Pic = Val(Parse(3))
  'Item(n).Type = Val(Parse(4))
  'Item(n).Data1 = Val(Parse(5))
  'Item(n).Data2 = Val(Parse(6))
  'Item(n).Data3 = Val(Parse(7))
  
  ' Save it
  Call SendUpdateItemToAll(n)
  Call SaveItem(n)
  Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub
Private Sub HandleRequestEditNPC(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  'Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
      
  Call SendDataTo(Index, "NPCEDITOR" & SEP_CHAR & END_CHAR)
End Sub
Private Sub HandleEditNPC(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Integer
  
  ' Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
      
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  ' The npc #
  n = GetIntegerFromBuffer(Buffer, True)
      
  ' Prevent hacking
  If n < 0 Or n > MAX_NPCS Then
    Call HackingAttempt(Index, "Invalid NPC Index")
    Exit Sub
  End If
  
  Call AddLog(GetPlayerName(Index) & " editing npc #" & n & ".", ADMIN_LOG)
  Call SendEditNpcTo(Index, n)
End Sub
Private Sub HandleSaveNPC(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Integer
  
  ' Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  n = GetIntegerFromBuffer(Buffer, True)
  
  ' Prevent hacking
  If n < 0 Or n > MAX_NPCS Then
    Call HackingAttempt(Index, "Invalid NPC Index")
    Exit Sub
  End If
  
  ' Update the npc
  Call CopyMemory(Npc(n), Buffer(0), aLen(Buffer))
  'Npc(n).Name = Parse(2)
  'Npc(n).AttackSay = Parse(3)
  'Npc(n).Sprite = Val(Parse(4))
  'Npc(n).SpawnSecs = Val(Parse(5))
  'Npc(n).Behavior = Val(Parse(6))
  'Npc(n).Range = Val(Parse(7))
  'Npc(n).DropChance = Val(Parse(8))
  'Npc(n).DropItem = Val(Parse(9))
  'Npc(n).DropItemValue = Val(Parse(10))
  'Npc(n).STR = Val(Parse(11))
  'Npc(n).DEF = Val(Parse(12))
  'Npc(n).SPEED = Val(Parse(13))
  'Npc(n).MAGI = Val(Parse(14))
  
  ' Save it
  Call SendUpdateNpcToAll(n)
  Call SaveNpc(n)
  Call AddLog(GetPlayerName(Index) & " saved npc #" & n & ".", ADMIN_LOG)
End Sub
Private Sub HandleRequestEditShop(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  ' Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
  
  Call SendDataTo(Index, "SHOPEDITOR" & SEP_CHAR & END_CHAR)
End Sub
Private Sub HandleEditShop(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Integer
  
  ' Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  ' The shop #
  n = GetIntegerFromBuffer(Buffer, True)
  
  ' Prevent hacking
  If n < 0 Or n > MAX_SHOPS Then
    Call HackingAttempt(Index, "Invalid Shop Index")
    Exit Sub
  End If
  
  Call AddLog(GetPlayerName(Index) & " editing shop #" & n & ".", ADMIN_LOG)
  Call SendEditShopTo(Index, n)
End Sub
Private Sub HandleSaveShop(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim ShopNum As Integer
  
  ' Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
  
  ShopNum = GetIntegerFromBuffer(Buffer, True)
  
  ' Prevent hacking
  If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
    Call HackingAttempt(Index, "Invalid Shop Index")
    Exit Sub
  End If
  
  Call CopyMemory(Shop(ShopNum), Buffer(0), aLen(Buffer))
  ' Update the shop
  'Shop(ShopNum).Name = Parse(2)
  'Shop(ShopNum).JoinSay = Parse(3)
  'Shop(ShopNum).LeaveSay = Parse(4)
  'Shop(ShopNum).FixesItems = Val(Parse(5))
  
  'n = 6
  'For i = 1 To MAX_TRADES
  '    Shop(ShopNum).TradeItem(i).GiveItem = Val(Parse(n))
  '    Shop(ShopNum).TradeItem(i).GiveValue = Val(Parse(n + 1))
  '    Shop(ShopNum).TradeItem(i).GetItem = Val(Parse(n + 2))
  '    Shop(ShopNum).TradeItem(i).GetValue = Val(Parse(n + 3))
  '    n = n + 4
  'Next i
  
  ' Save it
  Call SendUpdateShopToAll(ShopNum)
  Call SaveShop(ShopNum)
  Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
End Sub
Private Sub HandleRequestEditSpell(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  ' Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
  
  Call SendDataTo(Index, "SPELLEDITOR" & SEP_CHAR & END_CHAR)
End Sub
Private Sub HandleEditSpell(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Integer
  
  ' Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
  
  ' The spell #
  n = GetIntegerFromBuffer(Buffer, True)
  
  ' Prevent hacking
  If n < 0 Or n > MAX_SPELLS Then
    Call HackingAttempt(Index, "Invalid Spell Index")
    Exit Sub
  End If
  
  Call AddLog(GetPlayerName(Index) & " editing spell #" & n & ".", ADMIN_LOG)
  Call SendEditSpellTo(Index, n)
End Sub
Private Sub HandleSaveSpell(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Integer
  
  ' Prevent hacking
  If HasAccess(Index, ADMIN_DEVELOPER) = 0 Then Exit Sub
  
  ' Spell #
  n = GetIntegerFromBuffer(Buffer, True)
  
  ' Prevent hacking
  If n < 0 Or n > MAX_SPELLS Then
    Call HackingAttempt(Index, "Invalid Spell Index")
    Exit Sub
  End If
  
  ' Update the spell
  Call CopyMemory(Spell(n), Buffer(0), ByteLen)
  'Spell(n).Name = Parse(2)
  'Spell(n).ClassReq = Val(Parse(3))
  'Spell(n).LevelReq = Val(Parse(4))
  'Spell(n).Type = Val(Parse(5))
  'Spell(n).Data1 = Val(Parse(6))
  'Spell(n).Data2 = Val(Parse(7))
  'Spell(n).Data3 = Val(Parse(8))
          
  ' Save it
  Call SendUpdateSpellToAll(n)
  Call SaveSpell(n)
  Call AddLog(GetPlayerName(Index) & " saving spell #" & n & ".", ADMIN_LOG)
End Sub
Private Sub HandleSetAccess(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim i As Byte, n As Long
  
  ' Prevent hacking
  If HasAccess(Index, ADMIN_CREATOR) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  ' The index
  n = FindPlayer(GetStringFromBuffer(Buffer, True))
  
  ' The access
  i = GetByteFromBuffer(Buffer, True)
  
  ' Check for invalid access level
  If i >= 0 Or i <= 3 Then
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
    Call PlayerMsg(Index, "Invalid access level.", Red)
  End If
End Sub
Private Sub HandleWhosOnline(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Call SendWhosOnline(Index)
End Sub
Private Sub HandleSetMOTD(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim Msg As String
  
  ' Prevent hacking
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  Msg = GetStringFromBuffer(Buffer, True)
  
  Call PutVar(App.Path & "\data\motd.ini", "MOTD", "Msg", Msg)
  Call GlobalMsg("MOTD changed to: " & Msg, BrightCyan)
  Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Msg, ADMIN_LOG)
End Sub
Private Sub HandleTrade(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  If Map(GetPlayerMap(Index)).Shop > 0 Then
    Call SendTrade(Index, Map(GetPlayerMap(Index)).Shop)
  Else
    Call PlayerMsg(Index, "There is no shop here.", BrightRed)
  End If
End Sub
Private Sub HandleTradeRequest(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim i As Integer, n As Byte, x As Byte
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  ' Trade num
  n = GetByteFromBuffer(Buffer, True)
  
  ' Prevent hacking
  If (n <= 0) Or (n > MAX_TRADES) Then
    Call HackingAttempt(Index, "Trade Request Modification")
    Exit Sub
  End If
  
  ' Index for shop
  i = Map(GetPlayerMap(Index)).Shop
  
  ' Check if inv full
  x = FindOpenInvSlot(Index, Shop(i).TradeItem(n).GetItem)
  If x = 0 Then
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
Private Sub HandleFixItem(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim ItemNum As Integer
  Dim i As Long, n As Integer, DurNeeded As Long, GoldNeeded As Long
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  ' Inv num
  n = GetIntegerFromBuffer(Buffer, True)
  
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
End Sub
Private Sub HandleSearch(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim x As Byte, y As Byte
  Dim i As Integer
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  x = GetByteFromBuffer(Buffer, True)
  y = GetByteFromBuffer(Buffer, True)
  
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
        Call PlayerMsg(Index, "You see a " & Trim(StrConv(Item(MapItem(GetPlayerMap(Index), i).Num).ItemName, vbUnicode)) & ".", Yellow)
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
        Call PlayerMsg(Index, "Your target is now a " & Trim(Npc(MapNpc(GetPlayerMap(Index), i).Num).Name) & ".", Yellow)
        Exit Sub
      End If
    End If
  Next i
End Sub
Private Sub HandleParty(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Integer
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  n = FindPlayer(GetStringFromBuffer(Buffer, True))
  
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
End Sub
Private Sub HandleJoinParty(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim n As Integer
  
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
End Sub
Private Sub HandleLeaveParty(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim n As Long
  
  n = Player(Index).PartyPlayer
  
  If n > 0 Then
    If Player(Index).InParty = YES Then
      Call PlayerMsg(Index, "You have left the party.", Pink)
      Call PlayerMsg(n, GetPlayerName(Index) & " has left the party.", Pink)
    Else
      Call PlayerMsg(Index, "Declined party request.", Pink)
      Call PlayerMsg(n, GetPlayerName(Index) & " declined your request.", Pink)
    End If
    Player(Index).PartyPlayer = 0
    Player(Index).PartyStarter = NO
    Player(Index).InParty = NO
    Player(n).PartyPlayer = 0
    Player(n).PartyStarter = NO
    Player(n).InParty = NO
  Else
    Call PlayerMsg(Index, "You are not in a party!", Pink)
  End If
End Sub
Private Sub HandleSpells(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Call SendPlayerSpells(Index)
End Sub
Private Sub HandleCast(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  Dim Buffer() As Byte
  Dim n As Integer
  
  Buffer = FillBuffer(StartAddr, ByteLen)
  
  ' Spell slot
  n = GetIntegerFromBuffer(Buffer, True)
  
  Call CastSpell(Index, n)
End Sub
Private Sub HandleRequestLocation(ByVal Index As Long, ByVal StartAddr As Long, ByVal ByteLen As Long, ByVal ExtraVar As Long)
  'Prevent hacking
  If HasAccess(Index, ADMIN_MAPPER) = 0 Then Exit Sub
      
  Call PlayerMsg(Index, "Map: " & GetPlayerMap(Index) & ", X: " & GetPlayerX(Index) & ", Y: " & GetPlayerY(Index), Pink)
End Sub
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/??/2005  Verrigan   Rewrote Sub HandleData() to call the
'*                        new packet subroutines by their
'*                        addresses. This is a little more
'*                        efficient than a bunch of "If-Then"s.
'*                        (Not to mention, smaller. :P)
'****************************************************************
Sub HandleData(ByVal Index As Long, ByRef Buffer() As Byte)
On Error Resume Next
  Dim MsgType As Byte
  Dim StartAddr As Long
  
  MsgType = GetByteFromBuffer(Buffer, True)
  StartAddr = 0

  If aLen(Buffer) > 0 Then StartAddr = VarPtr(Buffer(0))
  
  If MsgType > SMSG_COUNT Then
    Call HackingAttempt(Index, "Packet Manipulation")
  Else
    Call CallWindowProc(HandleDataSub(MsgType), Index, StartAddr, aLen(Buffer), 0)
  End If
  If Err.Number <> 0 Then
    Call HackingAttempt(Index, "Packet Manipulation")
  End If
End Sub
