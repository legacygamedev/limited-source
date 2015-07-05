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

    sSpaces = Space$(5000)

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

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataTo(Index, "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToMap(MapNum, "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Call SendDataTo(Index, "ALERTMSG" & SEP_CHAR & Msg & END_CHAR)
    Call CloseSocket(Index)
End Sub

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Char(Player(Index).CharNum).GuildAccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal GuildAccess As Long)
    Player(Index).Char(Player(Index).CharNum).GuildAccess = GuildAccess
End Sub

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Char(Player(Index).CharNum).Guild = Guild
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

Function GetPlayerClassName(ByVal Index As Long) As String
    GetPlayerClassName = ClassData(GetPlayerClass(Index)).Name
End Function

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Char(Player(Index).CharNum).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).LEVEL
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal LEVEL As Long)
    Player(Index).Char(Player(Index).CharNum).LEVEL = LEVEL
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    If GetPlayerLevel(Index) <= MAX_LEVEL Then
        GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
    End If
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    If GetPlayerLevel(Index) < MAX_LEVEL Then
        Player(Index).Char(Player(Index).CharNum).Exp = Exp
    Else
        Call SetPlayerLevel(Index, MAX_LEVEL)
    End If
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(Player(Index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(Player(Index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).Char(Player(Index).CharNum).HP = HP

    If GetPlayerHP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).HP = 0
    End If

    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    End If

    Call SendHP(Index)
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If GetPlayerMP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).MP = 0
    End If

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    End If

    Call SendMP(Index)
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).SP = 0
    End If

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    End If

    Call SendSP(Index)
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).addHP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).addHP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).addHP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).addHP
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).addHP
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).addHP
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).addHP
    End If

    CharNum = Player(Index).CharNum
    ' GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSTR(index) / 2) + ClassData(Player(index).Char(CharNum).Class).STR) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(Index) * addHP.LEVEL) + (GetPlayerSTR(Index) * addHP.STR) + (GetPlayerDEF(Index) * addHP.DEF) + (GetPlayerMAGI(Index) * addHP.Magi) + (GetPlayerSPEED(Index) * addHP.Speed) + Add
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long

    Add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).addMP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).addMP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).addMP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).addMP
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).addMP
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).addMP
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).addMP
    End If

    CharNum = Player(Index).CharNum
    ' GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + ClassData(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(Index) * addMP.LEVEL) + (GetPlayerSTR(Index) * addMP.STR) + (GetPlayerDEF(Index) * addMP.DEF) + (GetPlayerMAGI(Index) * addMP.Magi) + (GetPlayerSPEED(Index) * addMP.Speed) + Add
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long
    'nullifies add
    Add = 0
    
    'Items SP bonus'
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).addSP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).addSP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).addSP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).addSP
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).addSP
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).addSP
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).addSP
    End If

    
    CharNum = Player(Index).CharNum
    GetPlayerMaxSP = (GetPlayerLevel(Index) * addSP.LEVEL) + (GetPlayerSTR(Index) * addSP.STR) + (GetPlayerDEF(Index) * addSP.DEF) + (GetPlayerMAGI(Index) * addSP.Magi) + (GetPlayerSPEED(Index) * addSP.Speed) + Add
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

Function GetPlayerSTR(ByVal Index As Long) As Long
    GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).STR
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).Speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Char(Player(Index).CharNum).Speed = Speed
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).Magi
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal Magi As Long)
    Player(Index).Char(Player(Index).CharNum).Magi = Magi
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(Player(Index).CharNum).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).Char(Player(Index).CharNum).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Char(Player(Index).CharNum).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = GameServer.Sockets.Item(Index).RemoteAddress
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).Char(Player(Index).CharNum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).Char(Player(Index).CharNum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).Char(Player(Index).CharNum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).Char(Player(Index).CharNum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ShieldSlot = InvNum
End Sub

Function GetPlayerNecklaceSlot(ByVal Index As Long) As Long
    GetPlayerNecklaceSlot = Player(Index).Char(Player(Index).CharNum).NecklaceSlot
End Function

Sub SetPlayerNecklaceSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).NecklaceSlot = InvNum
End Sub

Function GetPlayerRingSlot(ByVal Index As Long) As Long
    GetPlayerRingSlot = Player(Index).Char(Player(Index).CharNum).RingSlot
End Function

Sub SetPlayerRingSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).RingSlot = InvNum
End Sub

Function GetPlayerLegsSlot(ByVal Index As Long) As Long
    GetPlayerLegsSlot = Player(Index).Char(Player(Index).CharNum).LegsSlot
End Function

Sub SetPlayerLegsSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).LegsSlot = InvNum
End Sub

Function GetMapBootMap(ByVal Index As Long) As Long
    GetMapBootMap = Map(GetPlayerMap(Index)).BootMap
End Function

Function GetMapBootX(ByVal Index As Long) As Long
    GetMapBootX = Map(GetPlayerMap(Index)).BootX
End Function

Function GetMapBootY(ByVal Index As Long) As Long
    GetMapBootY = Map(GetPlayerMap(Index)).BootY
End Function

Sub MutePlayer(ByVal Index As Integer)
    Player(Index).Mute = True
End Sub

Sub UnMutePlayer(ByVal Index As Integer)
    Player(Index).Mute = False
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

Function IsConnected(ByVal Index As Long) As Boolean
    On Error Resume Next

    If GameServer.Sockets.Item(Index).Socket Is Nothing Then
        IsConnected = False
    Else
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If Player(Index).InGame Then
            IsPlaying = True
        Else
            IsPlaying = False
        End If
    Else
        IsPlaying = False
    End If
End Function

Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim I As Long

    packet = "PLAYERINV" & SEP_CHAR & Index & SEP_CHAR
    For I = 1 To MAX_INV
        packet = packet & GetPlayerInvItemNum(Index, I) & SEP_CHAR & GetPlayerInvItemValue(Index, I) & SEP_CHAR & GetPlayerInvItemDur(Index, I) & SEP_CHAR
    Next I
    packet = packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Call SendDataToMap(GetPlayerMap(Index), "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & Index & END_CHAR)
End Sub

Sub SendIndexInventoryFromMap(ByVal Index As Long)
    Dim packet As String
    Dim n As Long
    Dim I As Long
    
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = GetPlayerMap(Index) Then
                packet = "PLAYERINV" & SEP_CHAR & I & SEP_CHAR
                For n = 1 To MAX_INV
                    packet = packet & GetPlayerInvItemNum(I, n) & SEP_CHAR & GetPlayerInvItemValue(I, n) & SEP_CHAR & GetPlayerInvItemDur(I, n) & SEP_CHAR
                Next n
                packet = packet & END_CHAR

                Call SendDataTo(Index, packet)
            End If
        End If
    Next I
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    If IsPlaying(Index) Then
        Call SendDataToMap(GetPlayerMap(Index), "PLAYERWORNEQ" & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & GetPlayerLegsSlot(Index) & SEP_CHAR & GetPlayerRingSlot(Index) & SEP_CHAR & GetPlayerNecklaceSlot(Index) & END_CHAR)
    End If
End Sub

Sub SendHP(ByVal Index As Long)
   Call SendDataToMap(GetPlayerMap(Index), "PLAYERHP" & SEP_CHAR & Index & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR)
End Sub

Sub SendMP(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & END_CHAR)
End Sub

Sub SendSP(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & END_CHAR)
End Sub

Sub SendPTS(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR)
End Sub

Sub SendStats(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerSTR(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & END_CHAR)
End Sub

Sub Flash(ByVal Index As Long, ByVal flashfile As String)
    Call SendDataTo(Index, "flashevent" & SEP_CHAR & flashfile & END_CHAR)
End Sub

Sub Prompt(ByVal Index As Long, ByVal question As String, ByVal Value As Long)
    Call SendDataTo(Index, "prompt" & SEP_CHAR & question & SEP_CHAR & Value & END_CHAR)
End Sub

Sub PlaySound(ByVal Index As Long, ByVal Sound As String)
    Call SendDataTo(Index, "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Sub PlaySoundToMap(ByVal Index As Long, ByVal Sound As String)
    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim packet As String
    Dim j As Long

    ' Send index's player data to everyone including himself on the emap
    packet = "PLAYERDATA" & SEP_CHAR
    packet = packet & Index & SEP_CHAR
    packet = packet & GetPlayerName(Index) & SEP_CHAR
    packet = packet & GetPlayerSprite(Index) & SEP_CHAR
    packet = packet & GetPlayerMap(Index) & SEP_CHAR
    packet = packet & GetPlayerX(Index) & SEP_CHAR
    packet = packet & GetPlayerY(Index) & SEP_CHAR
    packet = packet & GetPlayerDir(Index) & SEP_CHAR
    packet = packet & GetPlayerAccess(Index) & SEP_CHAR
    packet = packet & GetPlayerPK(Index) & SEP_CHAR
    packet = packet & GetPlayerGuild(Index) & SEP_CHAR
    packet = packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    packet = packet & GetPlayerClass(Index) & SEP_CHAR
    packet = packet & GetPlayerHead(Index) & SEP_CHAR
    packet = packet & GetPlayerBody(Index) & SEP_CHAR
    packet = packet & GetPlayerleg(Index) & SEP_CHAR
    packet = packet & GetPlayerPaperdoll(Index) & SEP_CHAR
    packet = packet & GetPlayerLevel(Index) & SEP_CHAR
    
    packet = packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
    Dim dbytes() As Byte

    dbytes = StrConv(Data, vbFromUnicode)
    If IsConnected(Index) Then
        GameServer.Sockets.Item(Index).WriteBytes dbytes
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

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If I <> Index Then
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

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum And I <> Index Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
End Sub

Function GetPlayerCharNum(ByVal Index As Long) As Long
    GetPlayerCharNum = Player(Index).CharNum
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

Public Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    On Error GoTo WarpErr

    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If Not OldMap = MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    PlayersOnMap(MapNum) = YES

    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "warp" & END_CHAR)

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexInventoryFromMap(Index)
    Call SendIndexWornEquipmentFromMap(Index)

    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\main.ess", "OnMapLoad " & Index & "," & OldMap & "," & MapNum
    End If
    
    Exit Sub

WarpErr:
    Call AddLog("PlayerWarp error for player index " & Index & " on map " & GetPlayerMap(Index) & ".", "logs\ErrorLog.txt")
End Sub

Sub JoinWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    Call SendLeaveMap(Index, OldMap)

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    PlayersOnMap(MapNum) = YES

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexWornEquipmentFromMap(Index)
End Sub

Public Sub AddLog(ByVal text As String, ByVal FN As String)
    Dim FileName As String
    Dim FileID As Long

    If ServerLog Then
        FileName = App.Path & "\" & FN

        If FileExists(FN) Then
            FileID = FreeFile
            Open FileName For Output As #FileID
            Print #FileID, Time & ": " & text
            Close #FileID
        End If
    End If
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call AdminMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", WHITE)
            Call AddLog(GetPlayerName(Index) & " was kicked for '" & Reason & "'.", "Logs\HackAttempt.txt")
        End If

        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Sub BattleMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Long)
    Call SendDataTo(Index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function
Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).num = ItemNum
    Call SendBankUpdate(Index, BankSlot)
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Value = ItemValue
    Call SendBankUpdate(Index, BankSlot)
End Sub

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemDur = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Dur = ItemDur
End Sub

Function GetPlayerTarget(ByVal Index As Long)
    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
        GetPlayerTarget = Player(Index).Target
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
Sub PlayerQueryBox(ByVal Index As Long, ByVal Message As String, ByVal Script As Long)
    Call SendDataTo(Index, "querybox" & SEP_CHAR & Message & SEP_CHAR & Script & END_CHAR)
End Sub
Sub UpdatePaperDoll(ByVal Index As Long)
    Call SendDataToMap(GetPlayerMap(Index), "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
End Sub
Sub UpdateSprite(ByVal Index As Long)
    Call SendDataToMap(GetPlayerMap(Index), "updatesprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index))
End Sub

Sub SetMapNpcNumber(ByVal MapNum As Long, ByVal Index As Long, ByVal Number As Long)
    If Number >= 1 Then
        MapNPC(MapNum, Index).num = Number
    End If
End Sub

Sub SetMapNpcTarget(ByVal MapNum As Long, ByVal Index As Long, ByVal Target As Long)
    MapNPC(MapNum, Index).Target = Target
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

Sub SetMapNpcDir(ByVal MapNum As Long, ByVal Index As Long, ByVal Direction As Long)
    MapNPC(MapNum, Index).Dir = Direction
End Sub

Sub SetMapNpcY(ByVal MapNum As Long, ByVal Index As Long, ByVal NPC_Y As Long)
    If NPC_Y <= MAX_MAPY And NPC_Y >= 0 Then
        MapNPC(MapNum, Index).Y = NPC_Y
    End If
End Sub

Sub SetMapNpcX(ByVal MapNum As Long, ByVal Index As Long, ByVal NPC_X As Long)
    If NPC_X <= MAX_MAPX And NPC_X >= 0 Then
        MapNPC(MapNum, Index).X = NPC_X
    End If
End Sub

Sub SetMapNpcHP(ByVal MapNum As Long, ByVal Index As Long, ByVal HitPoints As Long)
    MapNPC(MapNum, Index).HP = HitPoints
End Sub

Sub SendNPC(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    Call SendDataToMap(MapNum, "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).num & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(MapNum, MapNpcNum).num).Big & END_CHAR)
End Sub

Function GetNpcMaxHP(ByVal Index As Long) As Long
    GetNpcMaxHP = NPC(Index).MAXHP
End Function

Function GetMapNpcNumber(ByVal MapNum As Long, ByVal Index As Long) As Long
    GetMapNpcNumber = MapNPC(MapNum, Index).num
End Function

Function GetMapNpcHP(ByVal MapNum As Long, ByVal Index As Long) As Long
    GetMapNpcHP = MapNPC(MapNum, Index).HP
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

Sub SendIndexWornEquipment(ByVal Index As Long)
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

    If GetPlayerArmorSlot(Index) > 0 Then
        Armor = GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Helmet = GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Shield = GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))
    End If
    If GetPlayerWeaponSlot(Index) > 0 Then
        Weapon = GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Legs = GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Ring = GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Necklace = GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))
    End If

    packet = "itemworn" & SEP_CHAR & Index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Sub SendIndexWornEquipmentFromMap(ByVal Index As Long)
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
            If GetPlayerMap(Index) = GetPlayerMap(I) Then
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
                Call SendDataTo(Index, packet)
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

Sub SendWhosOnline(ByVal Index As Integer)
    Dim S As String
    Dim n As Long, I As Long

    S = vbNullString
    n = 0
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And I <> Index Then
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

    Call PlayerMsg(Index, S, WhoColor)
End Sub

Sub GoShopping(ByVal Index As Long, ByVal ShopNum As Long)
    Call SendTrade(Index, ShopNum)
End Sub

Sub LockPlayer(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).Locked = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).Locked = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Function IsPlayerLocked(ByVal Index As Long) As Boolean
    IsPlayerLocked = Player(Index).Locked
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

Function GetInvItemslot(Index, F_item_num, F_Min_value, F_Min_dur)
    Dim F_slot

    F_slot = 1
    GetInvItemslot = 0

    Do While F_slot < 25
        If GetPlayerInvItemNum(Index, F_slot) = F_item_num Then
            If GetPlayerInvItemValue(Index, F_slot) >= F_Min_value Then
                If GetPlayerInvItemDur(Index, F_slot) >= F_Min_dur Then

                    GetInvItemslot = F_slot
                    F_slot = 25

                End If
            End If
        End If
        F_slot = F_slot + 1
    Loop

End Function

Sub giveplayeritemtoslot(Index, F_slot, F_num, F_Val, F_dur)
    Dim packet

    If F_num < 0 Or F_num > MAX_ITEMS Or F_Val < 0 Or F_slot < 1 Or F_slot > 24 Then
        Exit Sub
    End If

    Player(Index).Char(Player(Index).CharNum).Inv(F_slot).num = F_num
    Player(Index).Char(Player(Index).CharNum).Inv(F_slot).Value = F_Val
    Player(Index).Char(Player(Index).CharNum).Inv(F_slot).Dur = F_dur

    packet = "PLAYERINVUPDATE" & SEP_CHAR & F_slot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, F_slot) & SEP_CHAR & GetPlayerInvItemValue(Index, F_slot) & SEP_CHAR & GetPlayerInvItemDur(Index, F_slot) & SEP_CHAR & Index & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), packet)

End Sub

Function giveplayeritem(Index, F_num, F_Val, F_dur)
    Dim n

    If F_num < 0 Or F_num > MAX_ITEMS Or F_Val < 0 Then
        Exit Function
    End If

    giveplayeritem = 0
    n = FindOpenInvSlot(Index, F_num)

    If n <> 0 Then
        giveplayeritem = 1
        ' Set item in players inventor
        Call SetPlayerInvItemNum(Index, n, F_num)
        If Item(GetPlayerInvItemNum(Index, n)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, n)).Stackable = 1 Then
            Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + F_Val)
        Else
            Call SetPlayerInvItemValue(Index, n, 0)
        End If
        Call SetPlayerInvItemDur(Index, n, F_dur)
        Call SendInventoryUpdate(Index, n)
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

Sub CustomMenuShow(ByVal Index As Long, ByVal Title As String, ByVal FileName As String, ByVal closable As Long)
    Dim packet As String

    If Index <= 0 Or Index >= MAX_PLAYERS Or closable < 0 Or closable > 1 Then
        Exit Sub
    End If

    packet = "showcustommenu" & SEP_CHAR & Title & SEP_CHAR & FileName & SEP_CHAR & closable & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub CustomMenuClose(ByVal Index As Long)
    Dim packet As String

    packet = "closecustommenu" & END_CHAR
    Call SendDataTo(Index, packet)

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

Sub CustomMenuTextBox(ByVal player_index As Long, ByVal customIndex As Long, ByVal width As Long, ByVal Left As Long, ByVal top As Long, ByVal text As String)
    Dim packet As String

    If customIndex < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Sub
    End If

    packet = "loadtextboxcustommenu" & SEP_CHAR & customIndex & SEP_CHAR & Left & SEP_CHAR & width & SEP_CHAR & top & SEP_CHAR & text & END_CHAR
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

Sub Loadinternet(ByVal Index As Long, ByVal address As String)
    Dim packet As String

    If Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Sub
    End If

    packet = "loadinternetwindow" & SEP_CHAR & address & END_CHAR
    Call SendDataTo(Index, packet)

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


Function GetPlayerDirX(Index)

    If GetPlayerDir(Index) = 1 Then
        GetPlayerDirX = GetPlayerX(Index)
    ElseIf GetPlayerDir(Index) = 2 Then
        GetPlayerDirX = GetPlayerX(Index) - 1
    ElseIf GetPlayerDir(Index) = 3 Then
        GetPlayerDirX = GetPlayerX(Index) + 1
    ElseIf GetPlayerDir(Index) = 0 Then
        GetPlayerDirX = GetPlayerX(Index)
    End If

End Function


Function GetPlayerDirY(Index)

    If GetPlayerDir(Index) = 1 Then
        GetPlayerDirY = GetPlayerY(Index) + 1
    ElseIf GetPlayerDir(Index) = 2 Then
        GetPlayerDirY = GetPlayerY(Index)
    ElseIf GetPlayerDir(Index) = 3 Then
        GetPlayerDirY = GetPlayerY(Index)
    ElseIf GetPlayerDir(Index) = 0 Then
        GetPlayerDirY = GetPlayerY(Index) - 1
    End If

End Function


Sub SetSpeed(ByVal Index As Long, ByVal Movement As String, ByVal Speed As Long)
    If Speed < 1 Then
        Call PlayerMsg(Index, "Automated message: script error, value " & Speed & " at SetSpeed is too low - must be above 1.", RED)
    ElseIf Speed > 32 Then
        Call PlayerMsg(Index, "Automated message: script error, value " & Speed & " at SetSpeed is to high - must be below 32.", RED)
    Else
        If Int(Log(Speed) / Log(2)) <> Log(Speed) / Log(2) Then
            Exit Sub
        End If

        Call SendDataTo(Index, "setspeed" & SEP_CHAR & Movement & SEP_CHAR & Speed & END_CHAR)
    End If
End Sub


Function GetNpcX(ByVal MapNum As Long, ByVal MapNpcNum As Long)

    If MapNpcNum < 1 Or MapNpcNum > 25 Then
    Else
        GetNpcX = MapNPC(MapNum, MapNpcNum).X
    End If

End Function


Function GetPlayerTargetNpc(ByVal Index As Long)

    If Index > 0 Then
        If Player(Index).TargetType = TARGET_TYPE_NPC Then
            GetPlayerTargetNpc = Player(Index).TargetNPC
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

Sub Image(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal Tile As Long, ByVal top As Long, ByVal height As Long, ByVal Left As Long, ByVal width As Long)
    Dim packet As String
    packet = "fog" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Tile & SEP_CHAR & top & SEP_CHAR & height & SEP_CHAR & Left & SEP_CHAR & width & END_CHAR
    Call SendDataTo(Index, packet)
End Sub

Sub SetPlayerNameColor(ByVal Index As Long, ByVal Color As Long)
    Call SendDataTo(Index, "namecolor" & SEP_CHAR & Color & END_CHAR)
End Sub

Sub LockSpells(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedSpells = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedSpells = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Sub LockItems(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedItems = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedItems = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Sub LockAttack(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedAttack = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedAttack = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
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

Function GetPlayerHead(ByVal Index As Long)
    GetPlayerHead = Player(Index).Char(Player(Index).CharNum).Head
End Function

Public Sub SetPlayerHead(ByVal Index As Long, ByVal Head As Long)
    Player(Index).Char(Player(Index).CharNum).Head = Head
End Sub

Function GetPlayerBody(ByVal Index As Long)
    GetPlayerBody = Player(Index).Char(Player(Index).CharNum).Body
End Function

Public Sub SetPlayerBody(ByVal Index As Long, ByVal Body As Long)
    Player(Index).Char(Player(Index).CharNum).Body = Body
End Sub

Function GetPlayerleg(ByVal Index As Long)
    GetPlayerleg = Player(Index).Char(Player(Index).CharNum).Leg
End Function

Public Sub SetPlayerLeg(ByVal Index As Long, ByVal Leg As Long)
    Player(Index).Char(Player(Index).CharNum).Leg = Leg
End Sub

Public Sub MovePlayer(ByVal Index As Long, ByVal Direction As Long, ByVal Movement As Long)
     Dim X As Integer
     Dim Y As Integer
       
     Y = GetPlayerY(Index)
     X = GetPlayerX(Index)
       
     Select Case Direction
         Case DIR_UP
             Y = Y - 1
         Case DIR_DOWN
             Y = Y + 1
         Case DIR_LEFT
             X = X - 1
         Case DIR_RIGHT
             X = X + 1
    End Select
      
     Call PlayerMove(Index, Direction, Movement, X, Y)
     Call SendPlayerNewXY(Index)
End Sub

Public Sub SavePlayer(ByVal Index As Long)
    Call SavePlayer(Index)
End Sub

Function GetPlayerGender(ByVal Index As Long) As Long
    GetPlayerGender = Player(Index).Char(Player(Index).CharNum).Sex
End Function

Function GetPlayerPaperdoll(ByVal Index As Long) As Byte
    If Player(Index).InGame Then
        GetPlayerPaperdoll = Player(Index).Char(GetPlayerCharNum(Index)).PAPERDOLL
    End If
End Function

Public Sub SetPlayerPaperdoll(ByVal Index As Long, ByVal Mode As Byte)
    If Mode = 0 Or Mode = 1 Then
        If Player(Index).InGame Then
            Player(Index).Char(GetPlayerCharNum(Index)).PAPERDOLL = Mode
        End If
    End If
End Sub

Public Sub ChangeMapName(ByVal Index As Long, ByVal Name As String)
    Map(Index).Name = Name
    Call SaveMap(Index)
    Call SendMap(Index)
End Sub

' The subs below are added by DFA
Public Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Call modGameLogic.PlayerMapDropItem(Index, InvNum, Amount)
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

Sub DamageNPC(ByVal Index As Long, ByVal NPCnum As Long, ByVal Damage As Long)
    Call AttackNpc(Index, NPCnum, Damage)
    Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & NPCnum & END_CHAR)
End Sub

Sub DamagePlayer(ByVal Index As Long, ByVal PIndex As Long, ByVal Damage As Long)
    Call AttackPlayer(Index, PIndex, Damage)
End Sub

Sub NPCAttack(ByVal X As Long, ByVal Target As Long, ByVal Damage As Long)
    Call NpcAttackPlayer(X, Target, Damage)
End Sub

Sub FireArrow(ByVal Index As Long, ByVal arrow As Long, ByVal Dir As Long)
    Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & arrow & SEP_CHAR & Dir & END_CHAR)
End Sub

' Eclipse Stable Scripting Commands

Function GetTargetSTR(ByVal Index As Long) As Long
    Dim Target
    
    Target = GetPlayerTarget(Index)
    
    If Target > 0 Then
        GetTargetSTR = Player(Target).Char(Player(Target).CharNum).STR
    Else
        GetTargetSTR = 0
    End If
End Function

Function GetTargetDEF(ByVal Index As Long) As Long
    Dim Target
    
    Target = GetPlayerTarget(Index)

    If Target > 0 Then
        GetTargetDEF = Player(Target).Char(Player(Target).CharNum).DEF
    Else
        GetTargetDEF = 0
    End If
End Function

Function GetTargetMAGI(ByVal Index As Long) As Long
    Dim Target
    
    Target = GetPlayerTarget(Index)
    
    If Target > 0 Then
        GetTargetMAGI = Player(Target).Char(Player(Target).CharNum).Magi
    Else
        GetTargetMAGI = 0
    End If
End Function

Function GetTargetSPEED(ByVal Index As Long) As Long
    Dim Target
    
    Target = GetPlayerTarget(Index)
    
    If Target > 0 Then
        GetTargetSPEED = Player(Target).Char(Player(Target).CharNum).Speed
    Else
        GetTargetSPEED = 0
    End If
End Function

Function GetTargetJob(ByVal Index As Long) As String
    Dim Target
    
    Target = GetPlayerTarget(Index)
    
    If Target > 0 Then
        GetTargetJob = Trim$(ClassData(GetPlayerClass(Target)).Name)
    Else
        GetTargetJob = 0
    End If
End Function

Function GetTargetName(ByVal Index As Long) As String
    Dim Target
    
    Target = GetPlayerTarget(Index)
    
    If Target > 0 Then
        GetTargetName = Trim$(Player(Target).Char(Player(Target).CharNum).Name)
    Else
        GetTargetName = 0
    End If
End Function

Function GetTargetGuild(ByVal Index As Long) As String
    Dim Target
    
    Target = GetPlayerTarget(Index)
    
    If Target > 0 Then
        GetTargetGuild = Trim$(Player(Target).Char(Player(Target).CharNum).Guild)
    Else
        GetTargetGuild = 0
    End If
End Function

Function StopBGM(ByVal Index As Long)
    Call SendDataTo(Index, "StopBGM" & SEP_CHAR & END_CHAR)
    'Call PlayerMsg(index, "StopBGM function called, packet send...", RED)
End Function

Function PlayBGM(ByVal Index As Long)
    Call SendDataTo(Index, "PlayBGM" & SEP_CHAR & END_CHAR)
    'Call PlayerMsg(index, "playBGM function called, packet send...", RED)
End Function

Function StopSound(ByVal Index As Long)
    Call SendDataTo(Index, "StopSound" & SEP_CHAR & END_CHAR)
    'Call PlayerMsg(index, "StopSound function called, packet send...", RED)
End Function

Function PlayBGS(ByVal Index As Long, ByVal music As String)
    Call SendDataTo(Index, "bkgsound" & SEP_CHAR & music & SEP_CHAR & END_CHAR)
End Function

Sub SellHouse(Index)
    Call SendDataTo(Index, "HOUSESELL" & END_CHAR)
End Sub
