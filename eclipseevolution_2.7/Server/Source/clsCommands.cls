VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function GetVar(file As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    szReturn = vbNullString

    sSpaces = Space(5000)

    file = App.Path & "\" & file

    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), file)

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(file As String, Header As String, Var As String, Value As String)
    On Error GoTo PutVar_Error
    Dim fpath As String
    fpath = App.Path & "\" & file
    Call WritePrivateProfileString(Header, Var, Value, fpath)

    On Error GoTo 0
    Exit Sub

PutVar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutVar of Module modDatabase"
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerAccess(I) > 0 Then
                Call SendDataTo(I, "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
            End If
        End If
    Next I
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToAll("GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataTo(index, "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToMap(MapNum, "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
    Call SendDataTo(index, "ALERTMSG" & SEP_CHAR & Msg & END_CHAR)
    Call CloseSocket(index)
End Sub

Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Function GetPlayerName(ByVal index As Long) As String
    GetPlayerName = Trim$(Player(index).Char(Player(index).CharNum).Name)
End Function

Function GetPlayerGuild(ByVal index As Long) As String
    GetPlayerGuild = Trim$(Player(index).Char(Player(index).CharNum).Guild)
End Function

Function GetPlayerGuildAccess(ByVal index As Long) As Long
    GetPlayerGuildAccess = Player(index).Char(Player(index).CharNum).GuildAccess
End Function

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal GuildAccess As Long)
    Player(index).Char(Player(index).CharNum).GuildAccess = GuildAccess
End Sub

Sub SetPlayerGuild(ByVal index As Long, ByVal Guild As String)
    Player(index).Char(Player(index).CharNum).Guild = Guild
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Char(Player(index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Char(Player(index).CharNum).Class = ClassNum
End Sub

Function GetPlayerClassName(ByVal index As Long) As String
    GetPlayerClassName = ClassData(GetPlayerClass(index)).Name
End Function

Function GetPlayerSprite(ByVal index As Long) As Long
    GetPlayerSprite = Player(index).Char(Player(index).CharNum).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Char(Player(index).CharNum).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    GetPlayerLevel = Player(index).Char(Player(index).CharNum).LEVEL
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal LEVEL As Long)
    Player(index).Char(Player(index).CharNum).LEVEL = LEVEL
End Sub

Function GetPlayerNextLevel(ByVal index As Long) As Long
    If GetPlayerLevel(index) <= MAX_LEVEL Then
        GetPlayerNextLevel = Experience(GetPlayerLevel(index))
    End If
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).Char(Player(index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)
    Player(index).Char(Player(index).CharNum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    GetPlayerAccess = Player(index).Char(Player(index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Char(Player(index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    GetPlayerPK = Player(index).Char(Player(index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).Char(Player(index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal index As Long) As Long
    GetPlayerHP = Player(index).Char(Player(index).CharNum).HP
End Function

Sub SetPlayerHP(ByVal index As Long, ByVal HP As Long)
    Player(index).Char(Player(index).CharNum).HP = HP

    If GetPlayerHP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).HP = 0
    End If

    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then
        Player(index).Char(Player(index).CharNum).HP = GetPlayerMaxHP(index)
    End If

    Call SendHP(index)
End Sub

Function GetPlayerMP(ByVal index As Long) As Long
    GetPlayerMP = Player(index).Char(Player(index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
    Player(index).Char(Player(index).CharNum).MP = MP

    If GetPlayerMP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).MP = 0
    End If

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        Player(index).Char(Player(index).CharNum).MP = GetPlayerMaxMP(index)
    End If

    Call SendMP(index)
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
    GetPlayerSP = Player(index).Char(Player(index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
    Player(index).Char(Player(index).CharNum).SP = SP

    If GetPlayerSP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).SP = 0
    End If

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        Player(index).Char(Player(index).CharNum).SP = GetPlayerMaxSP(index)
    End If

    Call SendSP(index)
End Sub

Function GetPlayerMaxHP(ByVal index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        Add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).addHP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).addHP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).addHP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).addHP
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).addHP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).addHP
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).addHP
    End If

    CharNum = Player(index).CharNum
    ' GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSTR(index) / 2) + ClassData(Player(index).Char(CharNum).Class).STR) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(index) * addHP.LEVEL) + (GetPlayerSTR(index) * addHP.STR) + (GetPlayerDEF(index) * addHP.DEF) + (GetPlayerMAGI(index) * addHP.Magi) + (GetPlayerSPEED(index) * addHP.Speed) + Add
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long

    Add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        Add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).addMP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).addMP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).addMP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).addMP
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).addMP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).addMP
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).addMP
    End If

    CharNum = Player(index).CharNum
    ' GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + ClassData(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(index) * addMP.LEVEL) + (GetPlayerSTR(index) * addMP.STR) + (GetPlayerDEF(index) * addMP.DEF) + (GetPlayerMAGI(index) * addMP.Magi) + (GetPlayerSPEED(index) * addMP.Speed) + Add
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long
    GetPlayerMaxSP = (Player(index).Char(Player(index).CharNum).LEVEL + Int(GetPlayerSPEED(index) / 2) + ClassData(Player(index).Char(Player(index).CharNum).Class).Speed) * 2
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(ClassData(ClassNum).STR / 2) + ClassData(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(ClassData(ClassNum).Magi / 2) + ClassData(ClassNum).Magi) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(ClassData(ClassNum).Speed / 2) + ClassData(ClassNum).Speed) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = ClassData(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = ClassData(ClassNum).DEF
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = ClassData(ClassNum).Speed
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = ClassData(ClassNum).Magi
End Function

Function GetPlayerSTR(ByVal index As Long) As Long
    GetPlayerSTR = Player(index).Char(Player(index).CharNum).STR
End Function

Sub SetPlayerSTR(ByVal index As Long, ByVal STR As Long)
    Player(index).Char(Player(index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal index As Long) As Long
    GetPlayerDEF = Player(index).Char(Player(index).CharNum).DEF
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal DEF As Long)
    Player(index).Char(Player(index).CharNum).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal index As Long) As Long
    GetPlayerSPEED = Player(index).Char(Player(index).CharNum).Speed
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal Speed As Long)
    Player(index).Char(Player(index).CharNum).Speed = Speed
End Sub

Function GetPlayerMAGI(ByVal index As Long) As Long
    GetPlayerMAGI = Player(index).Char(Player(index).CharNum).Magi
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal Magi As Long)
    Player(index).Char(Player(index).CharNum).Magi = Magi
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    GetPlayerPOINTS = Player(index).Char(Player(index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    Player(index).Char(Player(index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    GetPlayerMap = Player(index).Char(Player(index).CharNum).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(index).Char(Player(index).CharNum).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    GetPlayerX = Player(index).Char(Player(index).CharNum).X
End Function

Sub SetPlayerX(ByVal index As Long, ByVal X As Long)
    Player(index).Char(Player(index).CharNum).X = X
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    GetPlayerY = Player(index).Char(Player(index).CharNum).Y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal Y As Long)
    Player(index).Char(Player(index).CharNum).Y = Y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    GetPlayerDir = Player(index).Char(Player(index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Char(Player(index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String
    GetPlayerIP = GameServer.Sockets.Item(index).RemoteAddress
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(index).Char(Player(index).CharNum).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(index).Char(Player(index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(index).Char(Player(index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerArmorSlot(ByVal index As Long) As Long
    GetPlayerArmorSlot = Player(index).Char(Player(index).CharNum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal index As Long) As Long
    GetPlayerWeaponSlot = Player(index).Char(Player(index).CharNum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long
    GetPlayerHelmetSlot = Player(index).Char(Player(index).CharNum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long
    GetPlayerShieldSlot = Player(index).Char(Player(index).CharNum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).ShieldSlot = InvNum
End Sub

Function GetPlayerNecklaceSlot(ByVal index As Long) As Long
    GetPlayerNecklaceSlot = Player(index).Char(Player(index).CharNum).NecklaceSlot
End Function

Sub SetPlayerNecklaceSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).NecklaceSlot = InvNum
End Sub

Function GetPlayerRingSlot(ByVal index As Long) As Long
    GetPlayerRingSlot = Player(index).Char(Player(index).CharNum).RingSlot
End Function

Sub SetPlayerRingSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).RingSlot = InvNum
End Sub

Function GetPlayerLegsSlot(ByVal index As Long) As Long
    GetPlayerLegsSlot = Player(index).Char(Player(index).CharNum).LegsSlot
End Function

Sub SetPlayerLegsSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).LegsSlot = InvNum
End Sub

Function GetMapBootMap(ByVal index As Long) As Long
    GetMapBootMap = Map(GetPlayerMap(index)).BootMap
End Function

Function GetMapBootX(ByVal index As Long) As Long
    GetMapBootX = Map(GetPlayerMap(index)).BootX
End Function

Function GetMapBootY(ByVal index As Long) As Long
    GetMapBootY = Map(GetPlayerMap(index)).BootY
End Function

Sub MutePlayer(ByVal index As Integer)
    Player(index).Mute = True
End Sub

Sub UnMutePlayer(ByVal index As Integer)
    Player(index).Mute = False
End Sub

Public Sub MakeDay()
    GameTime = TIME_DAY

    Hours = 7
    Minutes = 0
    Seconds = 0

    Call SendTimeToAll
End Sub

Public Sub MakeNight()
    GameTime = TIME_NIGHT

    Hours = 21
    Minutes = 0
    Seconds = 0

    Call SendTimeToAll
End Sub

Function IsScrolling() As Boolean
    If IS_SCROLLING = 0 Then
        IsScrolling = False
    Else
        IsScrolling = True
    End If
End Function

Function GetMaxPlayers() As Long
    GetMaxPlayers = MAX_PLAYERS
End Function

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim packet As String
    Dim I As Long

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    I = MapItemSlot

    If I <> 0 Then
        If ItemNum >= 0 Then
            If ItemNum <= MAX_ITEMS Then
                MapItem(MapNum, I).num = ItemNum
                MapItem(MapNum, I).Value = ItemVal
        
                If ItemNum <> 0 Then
                    If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_NECKLACE) Then
                        MapItem(MapNum, I).Dur = ItemDur
                    Else
                        MapItem(MapNum, I).Dur = 0
                    End If
                Else
                    MapItem(MapNum, I).Dur = 0
                End If
        
                MapItem(MapNum, I).X = X
                MapItem(MapNum, I).Y = Y
        
                packet = "SPAWNITEM" & SEP_CHAR & I & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR
                Call SendDataToMap(MapNum, packet)
            End If
        End If
    End If
    
    
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    On Error Resume Next

    If GameServer.Sockets.Item(index).Socket Is Nothing Then
        IsConnected = False
    Else
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If IsConnected(index) Then
        If Player(index).InGame Then
            IsPlaying = True
        Else
            IsPlaying = False
        End If
    Else
        IsPlaying = False
    End If
End Function

Sub SendInventory(ByVal index As Long)
    Dim packet As String
    Dim I As Long

    packet = "PLAYERINV" & SEP_CHAR & index & SEP_CHAR
    For I = 1 To MAX_INV
        packet = packet & GetPlayerInvItemNum(index, I) & SEP_CHAR & GetPlayerInvItemValue(index, I) & SEP_CHAR & GetPlayerInvItemDur(index, I) & SEP_CHAR
    Next I
    packet = packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(index), packet)
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
    Call SendDataToMap(GetPlayerMap(index), "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & index & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & index & END_CHAR)
End Sub

Sub SendIndexInventoryFromMap(ByVal index As Long)
    Dim packet As String
    Dim n As Long
    Dim I As Long
    
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = GetPlayerMap(index) Then
                packet = "PLAYERINV" & SEP_CHAR & I & SEP_CHAR
                For n = 1 To MAX_INV
                    packet = packet & GetPlayerInvItemNum(I, n) & SEP_CHAR & GetPlayerInvItemValue(I, n) & SEP_CHAR & GetPlayerInvItemDur(I, n) & SEP_CHAR
                Next n
                packet = packet & END_CHAR

                Call SendDataTo(index, packet)
            End If
        End If
    Next I
End Sub

Sub SendWornEquipment(ByVal index As Long)
    If IsPlaying(index) Then
        Call SendDataToMap(GetPlayerMap(index), "PLAYERWORNEQ" & SEP_CHAR & index & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & GetPlayerLegsSlot(index) & SEP_CHAR & GetPlayerRingSlot(index) & SEP_CHAR & GetPlayerNecklaceSlot(index) & END_CHAR)
    End If
End Sub

Sub SendHP(ByVal index As Long)
    Call SendDataTo(index, "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & END_CHAR)
End Sub

Sub SendMP(ByVal index As Long)
    Call SendDataTo(index, "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & END_CHAR)
End Sub

Sub SendSP(ByVal index As Long)
    Call SendDataTo(index, "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & END_CHAR)
End Sub

Sub SendPTS(ByVal index As Long)
    Call SendDataTo(index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(index) & END_CHAR)
End Sub

Sub SendStats(ByVal index As Long)
    Call SendDataTo(index, "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & GetPlayerNextLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerLevel(index) & END_CHAR)
End Sub

Sub Flash(ByVal index As Long, ByVal flashfile As String)
    Call SendDataTo(index, "flashevent" & SEP_CHAR & flashfile & END_CHAR)
End Sub

Sub Prompt(ByVal index As Long, ByVal question As String, ByVal Value As Long)
    Call SendDataTo(index, "prompt" & SEP_CHAR & question & SEP_CHAR & Value & END_CHAR)
End Sub

Sub PlaySound(ByVal index As Long, ByVal Sound As String)
    Call SendDataTo(index, "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Sub PlaySoundToMap(ByVal index As Long, ByVal Sound As String)
    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Sub SendPlayerData(ByVal index As Long)
    Dim packet As String
    Dim j As Long

    ' Send index's player data to everyone including himself on th emap
    packet = "PLAYERDATA" & SEP_CHAR
    packet = packet & index & SEP_CHAR
    packet = packet & GetPlayerName(index) & SEP_CHAR
    packet = packet & GetPlayerSprite(index) & SEP_CHAR
    packet = packet & GetPlayerMap(index) & SEP_CHAR
    packet = packet & GetPlayerX(index) & SEP_CHAR
    packet = packet & GetPlayerY(index) & SEP_CHAR
    packet = packet & GetPlayerDir(index) & SEP_CHAR
    packet = packet & GetPlayerAccess(index) & SEP_CHAR
    packet = packet & GetPlayerPK(index) & SEP_CHAR
    packet = packet & GetPlayerGuild(index) & SEP_CHAR
    packet = packet & GetPlayerGuildAccess(index) & SEP_CHAR
    packet = packet & GetPlayerClass(index) & SEP_CHAR
    packet = packet & GetPlayerHead(index) & SEP_CHAR
    packet = packet & GetPlayerBody(index) & SEP_CHAR
    packet = packet & GetPlayerleg(index) & SEP_CHAR
    packet = packet & GetPlayerPaperdoll(index) & SEP_CHAR
    
    packet = packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)
End Sub

Sub SendDataTo(ByVal index As Long, ByVal Data As String)
    Dim dbytes() As Byte

    dbytes = StrConv(Data, vbFromUnicode)
    If IsConnected(index) Then
        GameServer.Sockets.Item(index).WriteBytes dbytes
        DoEvents
    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendDataTo(I, Data)
        End If
    Next I
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If I <> index Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum And I <> index Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Char(Player(index).CharNum).Name = Name
End Sub

Function GetPlayerCharNum(ByVal index As Long) As Long
    GetPlayerCharNum = Player(index).CharNum
End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim I As Long

    Name = LCase$(Name)

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If Len(GetPlayerName(I)) >= Len(Name) Then
                If LCase$(GetPlayerName(I)) = Name Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If
    Next I
End Function

Public Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    On Error GoTo WarpErr

    If Not IsPlaying(index) Then
        Exit Sub
    End If

    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)

    If Not OldMap = MapNum Then
        Call SendLeaveMap(index, OldMap)
    End If

    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, X)
    Call SetPlayerY(index, Y)

    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    PlayersOnMap(MapNum) = YES

    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "warp" & END_CHAR)

    Player(index).GettingMap = YES

    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)

    Call SendInventory(index)
    Call SendIndexInventoryFromMap(index)
    Call SendIndexWornEquipmentFromMap(index)

    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnMapLoad " & index & "," & OldMap & "," & MapNum
    End If
    
    Exit Sub

WarpErr:
    Call AddLog("PlayerWarp error for player index " & index & " on map " & GetPlayerMap(index) & ".", "logs\ErrorLog.txt")
End Sub

Sub JoinWarp(ByVal index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    If Not IsPlaying(index) Then
        Exit Sub
    End If

    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)

    Call SendLeaveMap(index, OldMap)

    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, X)
    Call SetPlayerY(index, Y)

    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    PlayersOnMap(MapNum) = YES

    Player(index).GettingMap = YES

    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)

    Call SendInventory(index)
    Call SendIndexWornEquipmentFromMap(index)
End Sub

Public Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim FileName As String
    Dim FileID As Long

    If ServerLog Then
        FileName = App.Path & "\" & FN

        If FileExists(FN) Then
            FileID = FreeFile
            Open FileName For Output As #FileID
            Print #FileID, Time & ": " & Text
            Close #FileID
        End If
    End If
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
    If index > 0 Then
        If IsPlaying(index) Then
            Call AdminMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", WHITE)
            Call AddLog(GetPlayerName(index) & " was kicked for '" & Reason & "'.", "Logs\HackAttempt.txt")
        End If

        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Sub BattleMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Long)
    Call SendDataTo(index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function
Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Player(index).Char(Player(index).CharNum).Bank(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).num = ItemNum
    Call SendBankUpdate(index, BankSlot)
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Value = ItemValue
    Call SendBankUpdate(index, BankSlot)
End Sub

Function GetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemDur = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Dur = ItemDur
End Sub

Function GetPlayerTarget(ByVal index As Long)
    If Player(index).TargetType = TARGET_TYPE_PLAYER Then
        GetPlayerTarget = Player(index).Target
    Else
        GetPlayerTarget = -1
    End If
End Function
Sub SetTimer(ByVal Name As String, ByVal Interval As Long)
    Call AddNewTimer(Name, Interval)
End Sub

Function GetTimer(ByVal Name As String) As Long
    GetTimer = GetTimeLeft(Name)
End Function

Sub RemoveTimer(ByVal Name As String)
    Call GetRidOfTimer(Name)
End Sub

Sub SetTile(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal xset As Long, ByVal yset As Long, ByVal tileset As Long, ByVal layer As Long)
    Call ScriptSetTile(mapper, X, Y, xset, yset, tileset, layer)
End Sub
Function GetTileX(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)
    Select Case layer
        Case 0
            GetTileX = Map(mapper).Tile(X, Y).Ground - Int(Map(mapper).Tile(X, Y).Ground / 14) * 14
        Case 1
            GetTileX = Map(mapper).Tile(X, Y).Mask - Int(Map(mapper).Tile(X, Y).Mask / 14) * 14
        Case 2
            GetTileX = Map(mapper).Tile(X, Y).Anim - Int(Map(mapper).Tile(X, Y).Anim / 14) * 14
        Case 3
            GetTileX = Map(mapper).Tile(X, Y).Mask2 - Int(Map(mapper).Tile(X, Y).Mask2 / 14) * 14
        Case 4
            GetTileX = Map(mapper).Tile(X, Y).M2Anim - Int(Map(mapper).Tile(X, Y).M2Anim / 14) * 14
        Case 5
            GetTileX = Map(mapper).Tile(X, Y).Fringe - Int(Map(mapper).Tile(X, Y).Fringe / 14) * 14
        Case 6
            GetTileX = Map(mapper).Tile(X, Y).FAnim - Int(Map(mapper).Tile(X, Y).FAnim / 14) * 14
        Case 7
            GetTileX = Map(mapper).Tile(X, Y).Fringe2 - Int(Map(mapper).Tile(X, Y).Fringe2 / 14) * 14
        Case 8
            GetTileX = Map(mapper).Tile(X, Y).F2Anim - Int(Map(mapper).Tile(X, Y).F2Anim / 14) * 14
    End Select
End Function
Function GetTileY(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)
    Select Case layer
        Case 0
            GetTileY = Int(Map(mapper).Tile(X, Y).Ground / 14)
        Case 1
            GetTileY = Int(Map(mapper).Tile(X, Y).Mask / 14)
        Case 2
            GetTileY = Int(Map(mapper).Tile(X, Y).Anim / 14)
        Case 3
            GetTileY = Int(Map(mapper).Tile(X, Y).Mask2 / 14)
        Case 4
            GetTileY = Int(Map(mapper).Tile(X, Y).M2Anim / 14)
        Case 5
            GetTileY = Int(Map(mapper).Tile(X, Y).Fringe / 14)
        Case 6
            GetTileY = Int(Map(mapper).Tile(X, Y).FAnim / 14)
        Case 7
            GetTileY = Int(Map(mapper).Tile(X, Y).Fringe2 / 14)
        Case 8
            GetTileY = Int(Map(mapper).Tile(X, Y).F2Anim / 14)
    End Select
End Function
Function GetTileSet(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)
    Select Case layer
        Case 0
            GetTileSet = Map(mapper).Tile(X, Y).GroundSet
        Case 1
            GetTileSet = Map(mapper).Tile(X, Y).MaskSet
        Case 2
            GetTileSet = Map(mapper).Tile(X, Y).AnimSet
        Case 3
            GetTileSet = Map(mapper).Tile(X, Y).Mask2Set
        Case 4
            GetTileSet = Map(mapper).Tile(X, Y).M2AnimSet
        Case 5
            GetTileSet = Map(mapper).Tile(X, Y).FringeSet
        Case 6
            GetTileSet = Map(mapper).Tile(X, Y).FAnimSet
        Case 7
            GetTileSet = Map(mapper).Tile(X, Y).Fringe2Set
        Case 8
            GetTileSet = Map(mapper).Tile(X, Y).F2AnimSet
    End Select
End Function

Sub SendMap(ByVal mapper As Long)
    Call SendDataToMap(mapper, "CHECKFORMAP" & SEP_CHAR & mapper & SEP_CHAR & (Map(mapper).Revision + 1) & END_CHAR)
End Sub

Sub SpellAnim(ByVal SpellNum As Long, ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    Call SendDataToMap(mapper, "scriptspellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
End Sub
Sub SetAttribute(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)
    Call ScriptSetAttribute(mapper, X, Y, Attrib, Data1, Data2, Data3, String1, String2, String3)
End Sub
Function GetAttribute(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetAttribute = Map(mapper).Tile(X, Y).Type
End Function
Function GetTileData1(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileData1 = Map(mapper).Tile(X, Y).Data1
End Function
Function GetTileData2(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileData2 = Map(mapper).Tile(X, Y).Data2
End Function
Function GetTileData3(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileData3 = Map(mapper).Tile(X, Y).Data3
End Function
Function GetTileString1(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileString1 = Map(mapper).Tile(X, Y).String1
End Function
Function GetTileString2(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileString2 = Map(mapper).Tile(X, Y).String2
End Function
Function GetTileString3(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileString3 = Map(mapper).Tile(X, Y).String3
End Function
Sub PlayerQueryBox(ByVal index As Long, ByVal Message As String, ByVal Script As Long)
    Call SendDataTo(index, "querybox" & SEP_CHAR & Message & SEP_CHAR & Script & END_CHAR)
End Sub
Sub UpdatePaperDoll(ByVal index As Long)
    Call SendDataToMap(GetPlayerMap(index), "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)
End Sub
Sub UpdateSprite(ByVal index As Long)
    Call SendDataToMap(GetPlayerMap(index), "updatesprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index))
End Sub

Sub SetMapNpcNumber(ByVal MapNum As Long, ByVal index As Long, ByVal Number As Long)
    If Number >= 1 Then
        MapNPC(MapNum, index).num = Number
    End If
End Sub

Sub SetMapNpcTarget(ByVal MapNum As Long, ByVal index As Long, ByVal Target As Long)
    MapNPC(MapNum, index).Target = Target
End Sub

Public Sub SetMapName(ByVal MapNum As Long, ByVal MapName As String)
    Dim I As Long

    Map(MapNum).Name = MapName
    Map(MapNum).Revision = Map(MapNum).Revision + 1

    If GetTotalMapPlayers(MapNum) > 0 Then
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) Then
                If GetPlayerMap(I) = MapNum Then
                    Call SendMap(MapNum)
                End If
            End If
        Next I
    End If
End Sub

Sub SetMapNpcDir(ByVal MapNum As Long, ByVal index As Long, ByVal Direction As Long)
    MapNPC(MapNum, index).Dir = Direction
End Sub

Sub SetMapNpcY(ByVal MapNum As Long, ByVal index As Long, ByVal NPC_Y As Long)
    If NPC_Y <= MAX_MAPY And NPC_Y >= 0 Then
        MapNPC(MapNum, index).Y = NPC_Y
    End If
End Sub

Sub SetMapNpcX(ByVal MapNum As Long, ByVal index As Long, ByVal NPC_X As Long)
    If NPC_X <= MAX_MAPX And NPC_X >= 0 Then
        MapNPC(MapNum, index).X = NPC_X
    End If
End Sub

Sub SetMapNpcHP(ByVal MapNum As Long, ByVal index As Long, ByVal HitPoints As Long)
    MapNPC(MapNum, index).HP = HitPoints
End Sub

Sub SendNPC(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    Call SendDataToMap(MapNum, "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).num & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(MapNum, MapNpcNum).num).Big & END_CHAR)
End Sub

Function GetNpcMaxHP(ByVal index As Long) As Long
    GetNpcMaxHP = NPC(index).MAXHP
End Function

Function GetMapNpcNumber(ByVal MapNum As Long, ByVal index As Long) As Long
    GetMapNpcNumber = MapNPC(MapNum, index).num
End Function

Function GetMapNpcHP(ByVal MapNum As Long, ByVal index As Long) As Long
    GetMapNpcHP = MapNPC(MapNum, index).HP
End Function

Function GetNpcName(ByVal Number As Long) As String
    GetNpcName = Trim$(NPC(Number).Name)
End Function

Function GetNpcBehavior(ByVal Number As Long) As Long
    GetNpcBehavior = NPC(Number).Behavior
End Function

Function GetNpcExp(ByVal Number As Long) As Long
    GetNpcExp = NPC(Number).Exp
End Function

Function GetNpcDefense(ByVal Number As Long) As Long
    GetNpcDefense = NPC(Number).DEF
End Function

Function GetNpcStrength(ByVal Number As Long) As Long
    GetNpcStrength = NPC(Number).STR
End Function

Sub SendIndexWornEquipment(ByVal index As Long)
    Dim packet As String
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    Armor = 0
    Helmet = 0
    Shield = 0
    Weapon = 0
    Legs = 0
    Ring = 0
    Necklace = 0

    If GetPlayerArmorSlot(index) > 0 Then
        Armor = GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        Helmet = GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        Shield = GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))
    End If
    If GetPlayerWeaponSlot(index) > 0 Then
        Weapon = GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        Legs = GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))
    End If
    If GetPlayerRingSlot(index) > 0 Then
        Ring = GetPlayerInvItemNum(index, GetPlayerRingSlot(index))
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        Necklace = GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))
    End If

    packet = "itemworn" & SEP_CHAR & index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)
End Sub

Sub SendIndexWornEquipmentFromMap(ByVal index As Long)
    Dim packet As String
    Dim I As Long
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) = True Then
            If GetPlayerMap(index) = GetPlayerMap(I) Then
                Armor = 0
                Helmet = 0
                Shield = 0
                Weapon = 0
                Legs = 0
                Ring = 0
                Necklace = 0

                If GetPlayerArmorSlot(I) > 0 Then
                    Armor = GetPlayerInvItemNum(I, GetPlayerArmorSlot(I))
                End If
                If GetPlayerHelmetSlot(I) > 0 Then
                    Helmet = GetPlayerInvItemNum(I, GetPlayerHelmetSlot(I))
                End If
                If GetPlayerShieldSlot(I) > 0 Then
                    Shield = GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))
                End If
                If GetPlayerWeaponSlot(I) > 0 Then
                    Weapon = GetPlayerInvItemNum(I, GetPlayerWeaponSlot(I))
                End If
                If GetPlayerLegsSlot(I) > 0 Then
                    Legs = GetPlayerInvItemNum(I, GetPlayerLegsSlot(I))
                End If
                If GetPlayerRingSlot(I) > 0 Then
                    Ring = GetPlayerInvItemNum(I, GetPlayerRingSlot(I))
                End If
                If GetPlayerNecklaceSlot(I) > 0 Then
                    Necklace = GetPlayerInvItemNum(I, GetPlayerNecklaceSlot(I))
                End If

                packet = "itemworn" & SEP_CHAR & I & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR
                Call SendDataTo(index, packet)
            End If
        End If
    Next I
End Sub

Function GetPlayersOnMap(ByVal MapNum As Long) As Long
    Dim I As Long
    Dim Count As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                Count = Count + 1
            End If
        End If
    Next I

    GetPlayersOnMap = Count
End Function

Sub ScriptSpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long, ByVal spawn_X As Long, ByVal spawn_Y As Long, ByVal NPCnum As Long)
    ' NPC_index               map_number          X spawn          Y spawn            NPC_number
    Dim packet As String
    Dim I As Long

    ' Check for subscript out of range
    If MapNpcNum < 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    If NPCnum = 0 Then
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        MapNPC(MapNum, MapNpcNum).num = 0
        Map(MapNum).NPC(MapNpcNum) = 0
        MapNPC(MapNum, MapNpcNum).Target = 0
        MapNPC(MapNum, MapNpcNum).HP = 0
        MapNPC(MapNum, MapNpcNum).MP = 0
        MapNPC(MapNum, MapNpcNum).SP = 0
        MapNPC(MapNum, MapNpcNum).Dir = 0
        MapNPC(MapNum, MapNpcNum).X = 0
        MapNPC(MapNum, MapNpcNum).Y = 0

        ' Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(mapnum, MapNpcNum).num & SEP_CHAR & MapNpc(mapnum, MapNpcNum).X & SEP_CHAR & MapNpc(mapnum, MapNpcNum).Y & SEP_CHAR & MapNpc(mapnum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(mapnum, MapNpcNum).num).Big & END_CHAR
        ' Call SendDataToMap(mapnum, Packet)
        Call SaveMap(MapNum)
    End If

' MapNpc(mapnum, MapNpcNum).num = 0
' MapNpc(mapnum, MapNpcNum).SpawnWait = GetTickCount
' MapNpc(mapnum, MapNpcNum).HP = 0
' Call SendDataToMap(mapnum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)


    Map(MapNum).Revision = Map(MapNum).Revision + 1

    MapNPC(MapNum, MapNpcNum).num = NPCnum
    Map(MapNum).NPC(MapNpcNum) = NPCnum

    MapNPC(MapNum, MapNpcNum).Target = 0

    MapNPC(MapNum, MapNpcNum).HP = GetNpcMaxHP(NPCnum)
    MapNPC(MapNum, MapNpcNum).MP = GetNpcMaxMP(NPCnum)
    MapNPC(MapNum, MapNpcNum).SP = GetNpcMaxSP(NPCnum)

    MapNPC(MapNum, MapNpcNum).Dir = Int(Rnd * 4)

    MapNPC(MapNum, MapNpcNum).X = spawn_X
    MapNPC(MapNum, MapNpcNum).Y = spawn_Y

    packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).num & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(MapNum, MapNpcNum).num).Big & END_CHAR
    Call SendDataToMap(MapNum, packet)

    Call SaveMap(MapNum)

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
            Call SendDataTo(I, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(I) & SEP_CHAR & Map(GetPlayerMap(I)).Revision & END_CHAR)
        End If
    Next I

End Sub

Sub SpawnItemOnMap(ByVal Map As Long, ByVal X As Long, ByVal Y As Long, ByVal ItemNum As Long, ByVal Amount As Long, ByVal durability As Long)
    Dim I As Long

    If ItemNum > MAX_ITEMS Or ItemNum = 0 Then
        Exit Sub
    End If

    If Map < 0 Or Map > MAX_MAPS Then
        Exit Sub
    End If

    If Amount < 0 Then
        Exit Sub
    End If

    I = FindOpenMapItemSlot(Map)

    If I <> 0 Then

        Call SpawnItemSlot(I, ItemNum, Amount, durability, Map, X, Y)
    Else
        I = MAX_MAP_ITEMS
        Call SpawnItemSlot(I, ItemNum, Amount, durability, Map, X, Y)
    End If
End Sub

Function GetItemName(ByVal Number As Long)
    GetItemName = Item(Number).Name
End Function

Sub ClearItemSlot(ByVal Map As Long, ByVal ItemIndex As Long)
    Call SpawnItemSlot(ItemIndex, 0, 0, 0, Map, MapItem(Map, ItemIndex).X, MapItem(Map, ItemIndex).Y)
End Sub

Sub SendWhosOnline(ByVal index As Integer)
    Dim S As String
    Dim n As Long, I As Long

    S = vbNullString
    n = 0
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And I <> index Then
            S = S & GetPlayerName(I) & ", "
            n = n + 1
        End If
    Next I

    If n = 0 Then
        S = "There are no other players online."
    Else
        S = Mid$(S, 1, Len(S) - 2)
        S = "There are " & n & " other players online: " & S & "."
    End If

    Call PlayerMsg(index, S, WhoColor)
End Sub

Sub GoShopping(ByVal index As Long, ByVal ShopNum As Long)
    Call SendTrade(index, ShopNum)
End Sub

Sub LockPlayer(ByVal index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(index).Locked = False
        Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(index).Locked = True
        Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)
    End If
End Sub

Function IsPlayerLocked(ByVal index As Long) As Boolean
    IsPlayerLocked = Player(index).Locked
End Function

Public Sub SetServerTime(ByVal Time_Hour As Long, ByVal Time_Minute As Long, ByVal Time_Second As Long)
    If Time_Hour <> -1 Then
        If Time_Hour > -1 And Time_Hour < 24 Then
            Hours = Time_Hour
        End If
    End If

    If Time_Minute <> -1 Then
        If Time_Minute > -1 And Time_Minute < 60 Then
            Minutes = Time_Minute
        End If
    End If

    If Time_Second <> -1 Then
        If Time_Second > -1 And Time_Second < 60 Then
            Seconds = Time_Second
        End If
    End If
End Sub

Function GetServerTime(ByVal TimeType As Long) As Long
    Select Case TimeType
        Case 0
            GetServerTime = Hours

        Case 1
            GetServerTime = Minutes

        Case 2
            GetServerTime = Seconds
    End Select
End Function

Function GetInvItemslot(index, F_item_num, F_Min_value, F_Min_dur)
    Dim F_slot

    F_slot = 1
    GetInvItemslot = 0

    Do While F_slot < 25
        If GetPlayerInvItemNum(index, F_slot) = F_item_num Then
            If GetPlayerInvItemValue(index, F_slot) >= F_Min_value Then
                If GetPlayerInvItemDur(index, F_slot) >= F_Min_dur Then

                    GetInvItemslot = F_slot
                    F_slot = 25

                End If
            End If
        End If
        F_slot = F_slot + 1
    Loop

End Function

Sub giveplayeritemtoslot(index, F_slot, F_num, F_Val, F_dur)
    Dim packet

    If F_num < 0 Or F_num > MAX_ITEMS Or F_Val < 0 Or F_slot < 1 Or F_slot > 24 Then
        Exit Sub
    End If

    Player(index).Char(Player(index).CharNum).Inv(F_slot).num = F_num
    Player(index).Char(Player(index).CharNum).Inv(F_slot).Value = F_Val
    Player(index).Char(Player(index).CharNum).Inv(F_slot).Dur = F_dur

    packet = "PLAYERINVUPDATE" & SEP_CHAR & F_slot & SEP_CHAR & index & SEP_CHAR & GetPlayerInvItemNum(index, F_slot) & SEP_CHAR & GetPlayerInvItemValue(index, F_slot) & SEP_CHAR & GetPlayerInvItemDur(index, F_slot) & SEP_CHAR & index & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)

End Sub

Function giveplayeritem(index, F_num, F_Val, F_dur)
    Dim n

    If F_num < 0 Or F_num > MAX_ITEMS Or F_Val < 0 Then
        Exit Function
    End If

    giveplayeritem = 0
    n = FindOpenInvSlot(index, F_num)

    If n <> 0 Then
        giveplayeritem = 1
        ' Set item in players inventor
        Call SetPlayerInvItemNum(index, n, F_num)
        If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, n)).Stackable = 1 Then
            Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + F_Val)
        Else
            Call SetPlayerInvItemValue(index, n, 0)
        End If
        Call SetPlayerInvItemDur(index, n, F_dur)
        Call SendInventoryUpdate(index, n)
        Exit Function
    Else
        giveplayeritem = 0
        Exit Function
    End If
End Function

Sub textbubble(player_index, bubble_index, F_text, F_map, F_X, F_Y, F_colour)
    Dim packet

    If player_index < 1 Or player_index > MAX_PLAYERS Or bubble_index < 1 Or bubble_index > 20 Or F_map < 1 Or F_map > MAX_MAPS Or F_colour < 0 Or F_colour > 15 Then
        Exit Sub
    End If

    packet = "scriptbubble" & SEP_CHAR & bubble_index & SEP_CHAR & F_text & SEP_CHAR & F_map & SEP_CHAR & F_X & SEP_CHAR & F_Y & SEP_CHAR & F_colour & END_CHAR
    Call SendDataTo(player_index, packet)

End Sub

Sub maptextbubble(bubble_index, F_text, F_map, F_X, F_Y, F_colour)
    Dim packet
    Dim I

    If bubble_index < 1 Or bubble_index > 20 Or F_map < 1 Or F_map > MAX_MAPS Or F_colour < 0 Or F_colour > 15 Then
        Exit Sub
    End If

    packet = "scriptbubble" & SEP_CHAR & bubble_index & SEP_CHAR & F_text & SEP_CHAR & F_map & SEP_CHAR & F_X & SEP_CHAR & F_Y & SEP_CHAR & F_colour & END_CHAR

    For I = 1 To MAX_PLAYERS
        If GetPlayerMap(I) = F_map Then
            Call SendDataTo(I, packet)
        End If
    Next I

End Sub

Function GetMapName(F_map)
    GetMapName = Map(F_map).Name
End Function

Function GetMapUp(F_map)
    GetMapUp = Map(F_map).Up
End Function

Function GetMapDown(F_map)
    GetMapDown = Map(F_map).Down
End Function

Function GetMapLeft(F_map)
    GetMapLeft = Map(F_map).Left
End Function

Function GetMapRight(F_map)
    GetMapRight = Map(F_map).Right
End Function

Sub CustomMenuShow(ByVal index As Long, ByVal Title As String, ByVal FileName As String, ByVal closable As Long)
    Dim packet As String

    If index <= 0 Or index >= MAX_PLAYERS Or closable < 0 Or closable > 1 Then
        Exit Sub
    End If

    packet = "showcustommenu" & SEP_CHAR & Title & SEP_CHAR & FileName & SEP_CHAR & closable & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub CustomMenuClose(ByVal index As Long)
    Dim packet As String

    packet = "closecustommenu" & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub CustomMenuPicture(ByVal player_index As Long, ByVal picture_index As Long, ByVal FileName As String, ByVal Left As Long, ByVal top As Long)
    Dim packet As String

    If picture_index < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Sub
    End If

    packet = "loadpiccustommenu" & SEP_CHAR & picture_index & SEP_CHAR & FileName & SEP_CHAR & Left & SEP_CHAR & top & END_CHAR
    Call SendDataTo(player_index, packet)

End Sub

Sub CustomMenuLabel(ByVal player_index As Long, ByVal picture_index As Long, ByVal caption As String, ByVal Left As Long, ByVal top As Long, ByVal customsize As Long, ByVal customcolour As Long, ByVal alignment As Long, ByVal width As Long, ByVal height As Long)
    Dim packet As String

    If alignment < 0 Or alignment > 2 Or width < 0 Or height < 0 Or customcolour > 15 Or customcolour < 0 Or picture_index < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Sub
    End If

    packet = "loadlabelcustommenu" & SEP_CHAR & picture_index & SEP_CHAR & caption & SEP_CHAR & Left & SEP_CHAR & top & SEP_CHAR & customsize & SEP_CHAR & customcolour & SEP_CHAR & alignment & SEP_CHAR & width & SEP_CHAR & height & END_CHAR
    Call SendDataTo(player_index, packet)

End Sub

Sub CustomMenuTextBox(ByVal player_index As Long, ByVal customIndex As Long, ByVal width As Long, ByVal Left As Long, ByVal top As Long, ByVal Text As String)
    Dim packet As String

    If customIndex < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Sub
    End If

    packet = "loadtextboxcustommenu" & SEP_CHAR & customIndex & SEP_CHAR & Left & SEP_CHAR & width & SEP_CHAR & top & SEP_CHAR & Text & END_CHAR
    Call SendDataTo(player_index, packet)

End Sub

Function getplayermenuclicktitle(ByVal player_index As Long)
    Dim Msg As String

    If player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Function
    End If

    Msg = Player(player_index).CustomTitle

    getplayermenuclicktitle = Trim$(Msg)

End Function

Function getplayermenuclickmsg(ByVal player_index As Long)
    Dim Msg As String

    If player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Function
    End If

    Msg = Player(player_index).CustomMsg

    getplayermenuclickmsg = Trim$(Msg)
End Function

Sub Loadinternet(ByVal index As Long, ByVal address As String)
    Dim packet As String

    If index <= 0 Or index > MAX_PLAYERS Then
        Exit Sub
    End If

    packet = "loadinternetwindow" & SEP_CHAR & address & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub updateplayermenutext(ByVal player_index As Long, ByVal box_index As Long)
    Dim packet As String

    packet = "returncustomboxmsg" & SEP_CHAR & box_index & END_CHAR
    Call SendDataTo(player_index, packet)

End Sub

Function getplayermenutext(ByVal player_index As Long)

    getplayermenutext = Player(player_index).CustomMsg

End Function

Sub npcmoving(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Direction As Long, ByVal Speed As Long)

    If CanNpcMove(MapNum, MapNpcNum, Direction) Then
        Call NpcMove(MapNum, MapNpcNum, Direction, Speed)
    End If

End Sub


Function GetPlayerDirX(index)

    If GetPlayerDir(index) = 1 Then
        GetPlayerDirX = GetPlayerX(index)
    ElseIf GetPlayerDir(index) = 2 Then
        GetPlayerDirX = GetPlayerX(index) - 1
    ElseIf GetPlayerDir(index) = 3 Then
        GetPlayerDirX = GetPlayerX(index) + 1
    ElseIf GetPlayerDir(index) = 0 Then
        GetPlayerDirX = GetPlayerX(index)
    End If

End Function


Function GetPlayerDirY(index)

    If GetPlayerDir(index) = 1 Then
        GetPlayerDirY = GetPlayerY(index) + 1
    ElseIf GetPlayerDir(index) = 2 Then
        GetPlayerDirY = GetPlayerY(index)
    ElseIf GetPlayerDir(index) = 3 Then
        GetPlayerDirY = GetPlayerY(index)
    ElseIf GetPlayerDir(index) = 0 Then
        GetPlayerDirY = GetPlayerY(index) - 1
    End If

End Function


Sub SetSpeed(ByVal index As Long, ByVal Movement As String, ByVal Speed As Long)
    If Speed < 1 Then
        Call PlayerMsg(index, "Automated message: script error, value " & Speed & " at SetSpeed is too low - must be above 1.", RED)
    ElseIf Speed > 32 Then
        Call PlayerMsg(index, "Automated message: script error, value " & Speed & " at SetSpeed is to high - must be below 32.", RED)
    Else
        If Int(Log(Speed) / Log(2)) <> Log(Speed) / Log(2) Then
            Exit Sub
        End If

        Call SendDataTo(index, "setspeed" & SEP_CHAR & Movement & SEP_CHAR & Speed & END_CHAR)
    End If
End Sub


Function GetNpcX(ByVal MapNum As Long, ByVal MapNpcNum As Long)

    If MapNpcNum < 1 Or MapNpcNum > 25 Then
    Else
        GetNpcX = MapNPC(MapNum, MapNpcNum).X
    End If

End Function


Function GetPlayerTargetNpc(ByVal index As Long)

    If index > 0 Then
        If Player(index).TargetType = TARGET_TYPE_NPC Then
            GetPlayerTargetNpc = Player(index).TargetNPC
        Else
            GetPlayerTargetNpc = -1
        End If
    End If

End Function


Function GetNpcY(ByVal MapNum As Long, ByVal MapNpcNum As Long)

    If MapNpcNum > 0 Then
        GetNpcY = MapNPC(MapNum, MapNpcNum).Y
    End If

End Function

Sub SetWeather(ByVal MapNum As Long, ByVal Weather As Long, ByVal Interval As Long)
    Call SendDataToMap(MapNum, "mapweather" & SEP_CHAR & MapNum & SEP_CHAR & Weather & SEP_CHAR & Interval & END_CHAR)
End Sub

Function GetWeather(ByVal MapNum As Long)
    GetWeather = Map(MapNum).Weather
End Function

Sub Image(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal Tile As Long, ByVal top As Long, ByVal height As Long, ByVal Left As Long, ByVal width As Long)
    Dim packet As String
    packet = "fog" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Tile & SEP_CHAR & top & SEP_CHAR & height & SEP_CHAR & Left & SEP_CHAR & width & END_CHAR
    Call SendDataTo(index, packet)
End Sub

Sub SetPlayerNameColor(ByVal index As Long, ByVal Color As Long)
    Call SendDataTo(index, "namecolor" & SEP_CHAR & Color & END_CHAR)
End Sub

Sub LockSpells(ByVal index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(index).LockedSpells = False
        Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(index).LockedSpells = True
        Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)
    End If
End Sub

Sub LockItems(ByVal index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(index).LockedItems = False
        Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(index).LockedItems = True
        Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)
    End If
End Sub

Sub LockAttack(ByVal index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(index).LockedAttack = False
        Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(index).LockedAttack = True
        Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)
    End If
End Sub

Function GetIndexPlayerOnMap(ByVal Map As Long, ByVal X As Long, ByVal Y As Long)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(I) = Map Then
            If GetPlayerX(I) = X And GetPlayerY(I) = Y Then
                GetIndexPlayerOnMap = I
                Exit Function
            End If
        End If
    Next I
End Function

Function GetPlayerHead(ByVal index As Long)
    GetPlayerHead = Player(index).Char(Player(index).CharNum).Head
End Function

Public Sub SetPlayerHead(ByVal index As Long, ByVal Head As Long)
    Player(index).Char(Player(index).CharNum).Head = Head
End Sub

Function GetPlayerBody(ByVal index As Long)
    GetPlayerBody = Player(index).Char(Player(index).CharNum).Body
End Function

Public Sub SetPlayerBody(ByVal index As Long, ByVal Body As Long)
    Player(index).Char(Player(index).CharNum).Body = Body
End Sub

Function GetPlayerleg(ByVal index As Long)
    GetPlayerleg = Player(index).Char(Player(index).CharNum).Leg
End Function

Public Sub SetPlayerLeg(ByVal index As Long, ByVal Leg As Long)
    Player(index).Char(Player(index).CharNum).Leg = Leg
End Sub

Public Sub MovePlayer(ByVal index As Long, ByVal Direction As Long, ByVal Movement As Long)
    Call PlayerMove(index, Direction, Movement)
End Sub

Public Sub SavePlayer(ByVal index As Long)
    Call SavePlayer(index)
End Sub

Function GetPlayerGender(ByVal index As Long) As Long
    GetPlayerGender = Player(index).Char(Player(index).CharNum).Sex
End Function

Function GetPlayerPaperdoll(ByVal index As Long) As Byte
    If Player(index).InGame Then
        GetPlayerPaperdoll = Player(index).Char(GetPlayerCharNum(index)).PAPERDOLL
    End If
End Function

Public Sub SetPlayerPaperdoll(ByVal index As Long, ByVal Mode As Byte)
    If Mode = 0 Or Mode = 1 Then
        If Player(index).InGame Then
            Player(index).Char(GetPlayerCharNum(index)).PAPERDOLL = Mode
        End If
    End If
End Sub

Public Sub ChangeMapName(ByVal index As Long, ByVal Name As String)
    Map(index).Name = Name
    Call SaveMap(index)
    Call SendMap(index)
End Sub

' The subs below are added by DFA
Public Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Call modGameLogic.PlayerMapDropItem(index, InvNum, Amount)
End Sub

Public Function ScriptGetTickCount() As Long
    ScriptGetTickCount = GetTickCount
End Function


' Eclipse 2.7 Scripting Commands

Public Function GetMaxMapX() As Byte
    GetMaxMapX = MAX_MAPX
End Function

Public Function GetMaxMapY() As Byte
    GetMaxMapY = MAX_MAPY
End Function

Public Sub SetMapMoral(ByVal MapNum As Long, ByVal Moral As Long)
    Map(MapNum).Moral = Moral
End Sub

Public Function GetMapMoral(ByVal MapNum As Long) As Byte
    GetMapMoral = Map(MapNum).Moral
End Function

Public Sub ShopReload(ByVal ShopNum As Long)
    Dim FileID As Integer
    Dim FileName As String
    
    FileName = App.Path & "\Shops\Shop" & ShopNum & ".dat"

    FileID = FreeFile

    Open FileName For Binary As #FileID
        Get #FileID, , Shop(ShopNum)
    Close #FileID
End Sub

Sub DamageNPC(ByVal index As Long, ByVal NPCnum As Long, ByVal Damage As Long)
    Call AttackNpc(index, NPCnum, Damage)
    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & NPCnum & END_CHAR)
End Sub

Sub DamagePlayer(ByVal index As Long, ByVal PIndex As Long, ByVal Damage As Long)
    Call AttackPlayer(index, PIndex, Damage)
End Sub

Sub NPCAttack(ByVal X As Long, ByVal Target As Long, ByVal Damage As Long)
    Call NpcAttackPlayer(X, Target, Damage)
End Sub

Sub FireArrow(ByVal index As Long, ByVal arrow As Long, ByVal Dir As Long)
    Call SendDataToMap(GetPlayerMap(index), "checkarrows" & SEP_CHAR & index & SEP_CHAR & arrow & SEP_CHAR & Dir & END_CHAR)
End Sub
