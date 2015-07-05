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
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
    
    File = App.Path & "\" & File
    
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = left(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, Header As String, Var As String, Value As String)
    File = App.Path & "\" & File
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Public Sub GlobalMsg(ByVal msg As String, ByVal Color As Byte)
Dim Packet As String
    
    Packet = "GLOBALMSG" & SEP_CHAR & msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Public Sub AdminMsg(ByVal msg As String, ByVal Color As Byte)
Dim Packet As String
Dim I As Long

    Packet = "ADMINMSG" & SEP_CHAR & msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerAccess(I) > 0 Then
            Call SendDataTo(I, Packet)
        End If
    Next I
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal msg As String, ByVal Color As Byte)
Dim Packet As String
Dim text As String

    Packet = "MAPMSG" & SEP_CHAR & msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
    Call CloseSocket(index)
End Sub



' :::::::::::::::::::::::::::::::::::::::::::
' :: Functions/Subs Available To The Owner ::
' :::::::::::::::::::::::::::::::::::::::::::
Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim(Player(index).Login)
End Function

Function GetPlayerName(ByVal index As Long) As String
    GetPlayerName = Trim(Player(index).Char(Player(index).CharNum).Name)
End Function

Function GetPlayerGuild(ByVal index As Long) As String
    GetPlayerGuild = Trim(Player(index).Char(Player(index).CharNum).Guild)
End Function

Function GetPlayerGuildAccess(ByVal index As Long) As Long
    GetPlayerGuildAccess = Player(index).Char(Player(index).CharNum).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal Guildaccess As Long)
    Player(index).Char(Player(index).CharNum).Guildaccess = Guildaccess
End Sub

Sub setplayerguildname(ByVal index As Long, ByVal guildname As String)
        Player(index).Char(Player(index).CharNum).Guild = Trim(guildname)
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Char(Player(index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Char(Player(index).CharNum).Class = ClassNum
End Sub

Function GetPlayerClassName(ByVal index As Long) As String
    GetPlayerClassName = GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "CLASS", "Name")
End Function

Function GetPlayerSprite(ByVal index As Long) As Long
    GetPlayerSprite = Player(index).Char(Player(index).CharNum).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Char(Player(index).CharNum).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    GetPlayerLevel = Player(index).Char(Player(index).CharNum).Level
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
    Player(index).Char(Player(index).CharNum).Level = Level
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

If index <= 0 Or index >= MAX_PLAYERS Then
    Exit Function
End If

    GetPlayerAccess = Player(index).Char(Player(index).CharNum).access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal access As Long)
    Player(index).Char(Player(index).CharNum).access = access
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
    
    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then
        Player(index).Char(Player(index).CharNum).HP = GetPlayerMaxHP(index)
    End If
    If GetPlayerHP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).HP = 0
    End If
    Call SendStats(index)
End Sub

Function GetPlayerMP(ByVal index As Long) As Long
    GetPlayerMP = Player(index).Char(Player(index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
    Player(index).Char(Player(index).CharNum).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        Player(index).Char(Player(index).CharNum).MP = GetPlayerMaxMP(index)
    End If
    If GetPlayerMP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).MP = 0
    End If
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
    GetPlayerSP = Player(index).Char(Player(index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
    Player(index).Char(Player(index).CharNum).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        Player(index).Char(Player(index).CharNum).SP = GetPlayerMaxSP(index)
    End If
    If GetPlayerSP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).SP = 0
    End If
End Sub

Function GetPlayerMaxHP(ByVal index As Long) As Long
Dim CharNum As Long
Dim I As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddHP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddHP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddHP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddHP
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).AddHP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddHP
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).AddHP
    End If

    CharNum = Player(index).CharNum
    'GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSTR(index) / 2) + Class(Player(index).Char(CharNum).Class).STR) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(index) * AddHP.Level) + (GetPlayerSTR(index) * AddHP.STR) + (GetPlayerDEF(index) * AddHP.DEF) + (GetPlayerMAGI(index) * AddHP.Magi) + (GetPlayerSPEED(index) * AddHP.Speed) + add
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
Dim CharNum As Long
Dim add As Long

add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddMP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddMP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddMP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddMP
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).AddMP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddMP
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).AddMP
    End If

    CharNum = Player(index).CharNum
    'GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(index) * AddMP.Level) + (GetPlayerSTR(index) * AddMP.STR) + (GetPlayerDEF(index) * AddMP.DEF) + (GetPlayerMAGI(index) * AddMP.Magi) + (GetPlayerSPEED(index) * AddMP.Speed) + add
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long
Dim CharNum As Long

    CharNum = Player(index).CharNum
    GetPlayerMaxSP = (Player(index).Char(CharNum).Level + Int(GetPlayerSPEED(index) / 2) + Class(Player(index).Char(CharNum).Class).Speed) * 2
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(Class(ClassNum).STR / 2) + Class(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(Class(ClassNum).Magi / 2) + Class(ClassNum).Magi) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(Class(ClassNum).Speed / 2) + Class(ClassNum).Speed) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = Class(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = Class(ClassNum).DEF
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = Class(ClassNum).Speed
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = Class(ClassNum).Magi
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
    GetPlayerMap = Player(index).Char(Player(index).CharNum).map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(index).Char(Player(index).CharNum).map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    GetPlayerX = Player(index).Char(Player(index).CharNum).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    Player(index).Char(Player(index).CharNum).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    GetPlayerY = Player(index).Char(Player(index).CharNum).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    Player(index).Char(Player(index).CharNum).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    GetPlayerDir = Player(index).Char(Player(index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Char(Player(index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
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

Function GetBootMap(ByVal index As Long) As Long
    GetBootMap = map(GetPlayerMap(index)).BootMap
End Function

Function GetBootX(ByVal index As Long) As Long
    GetBootX = map(GetPlayerMap(index)).BootX
End Function

Function GetBootY(ByVal index As Long) As Long
    GetBootY = map(GetPlayerMap(index)).BootY
End Function

Function IsScrolling()

If IS_SCROLLING = 1 Then
IsScrolling = 1
Else
IsScrolling = 0
End If

End Function

Function GetMaxPlayers()

GetMaxPlayers = MAX_PLAYERS

End Function

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim Packet As String
Dim I As Long
    
    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    I = MapItemSlot
    
    If I <> 0 And ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
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
        
        MapItem(MapNum, I).x = x
        MapItem(MapNum, I).y = y
            
        Packet = "SPAWNITEM" & SEP_CHAR & I & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
If index < 1 Then Exit Function
    If IsConnected(index) And Player(index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Sub SendInventory(ByVal index As Long)
Dim Packet As String
Dim I As Long

    Packet = "PLAYERINV" & SEP_CHAR & index & SEP_CHAR
    For I = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(index, I) & SEP_CHAR & GetPlayerInvItemValue(index, I) & SEP_CHAR & GetPlayerInvItemDur(index, I) & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & index & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & index & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendWornEquipment(ByVal index As Long)

Dim Packet As String
    
    If IsPlaying(index) Then
        Packet = "PLAYERWORNEQ" & SEP_CHAR & index & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & GetPlayerLegsSlot(index) & SEP_CHAR & GetPlayerRingSlot(index) & SEP_CHAR & GetPlayerNecklaceSlot(index) & SEP_CHAR & END_CHAR
        Call SendDataToMap(GetPlayerMap(index), Packet)
        Call SendIndexWornEquipment(index)
    End If
End Sub

Sub SendHP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
    
    Packet = "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendMP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendSP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendStats(ByVal index As Long)
Dim Packet As String
    
    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & GetPlayerNextLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub Flash(ByVal index As Long, ByVal flashfile As String)
    Call SendDataTo(index, "flashevent" & SEP_CHAR & flashfile & SEP_CHAR & END_CHAR)
End Sub

Sub Prompt(ByVal index As Long, ByVal question As String, ByVal Value As Long)
    Call SendDataTo(index, "prompt" & SEP_CHAR & question & SEP_CHAR & Value & SEP_CHAR & END_CHAR)
End Sub

Sub PlaySound(ByVal index As Long, ByVal Sound As String)
    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & SEP_CHAR & END_CHAR)
End Sub

Sub SendPlayerData(ByVal index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & index & SEP_CHAR
    Packet = Packet & GetPlayerName(index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(index) & SEP_CHAR
    Packet = Packet & GetPlayerX(index) & SEP_CHAR
    Packet = Packet & GetPlayerY(index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendDataTo(ByVal index As Long, ByVal Data As String)
Dim I As Long, n As Long, startc As Long

    If IsConnected(index) Then
        frmServer.Socket(index).SendData Data
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
        If IsPlaying(I) And I <> index Then
            Call SendDataTo(I, Data)
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

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(I)) >= Len(Trim(Name)) Then
                If UCase(Mid(GetPlayerName(I), 1, Len(Trim(Name)))) = UCase(Trim(Name)) Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If
    Next I
    
    FindPlayer = 0
End Function

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim Packet As String
Dim OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Check if there was an npc on the map the player is leaving, and if so say goodbye
    'If Trim(Shop(ShopNum).LeaveSay) <> "" Then
        'Call PlayerMsg(Index, Trim(Shop(ShopNum).Name) & " : " & Trim(Shop(ShopNum).LeaveSay) & "", SayColor)
    'End If
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)
    Call SendLeaveMap(index, OldMap)
    
    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
                
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES

    Player(index).GettingMap = YES
    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "warp" & SEP_CHAR & END_CHAR)
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & map(MapNum).Revision & SEP_CHAR & END_CHAR)
    
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendIndexWornEquipmentFromMap(index)
    
    If Scripting = 1 Then
    MyScript.ExecuteStatement "Scripts\Main.txt", "OnMapLoad " & index
    End If
    
    Packet = "forcehouseclose" & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
    
End Sub

Sub AddLog(ByVal text As String, ByVal FN As String)
Dim FileName As String
Dim f As Long

    If ServerLog = True Then
        FileName = App.Path & "\" & FN
    
        If Not FileExist(FN) Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If
    
        'f = FreeFile
        'Open FileName For Append As #f
            'Print #f, Time & ": " & text
        'Close #f
    End If
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If
    
        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Sub BattleMsg(ByVal index As Long, ByVal msg As String, ByVal Color As Byte, ByVal Side As Long)
    Call SendDataTo(index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR)
End Sub

Function Rand(ByVal High As Long, ByVal Low As Long)
Randomize
High = High + 1
Do Until Rand >= Low
    Rand = Int(Rnd * High)
Loop
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

Sub SetTile(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal xset As Long, ByVal yset As Long, ByVal tileset As Long, ByVal layer As Long)
Call ScriptSetTile(mapper, x, y, xset, yset, tileset, layer)
End Sub
Function GetTileX(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal layer As Long)
Select Case layer
Case 0
    GetTileX = map(mapper).Tile(x, y).Ground - Int(map(mapper).Tile(x, y).Ground / 14) * 14
Case 1
    GetTileX = map(mapper).Tile(x, y).Mask - Int(map(mapper).Tile(x, y).Mask / 14) * 14
Case 2
    GetTileX = map(mapper).Tile(x, y).Anim - Int(map(mapper).Tile(x, y).Anim / 14) * 14
Case 3
    GetTileX = map(mapper).Tile(x, y).Mask2 - Int(map(mapper).Tile(x, y).Mask2 / 14) * 14
Case 4
    GetTileX = map(mapper).Tile(x, y).M2Anim - Int(map(mapper).Tile(x, y).M2Anim / 14) * 14
Case 5
    GetTileX = map(mapper).Tile(x, y).Fringe - Int(map(mapper).Tile(x, y).Fringe / 14) * 14
Case 6
    GetTileX = map(mapper).Tile(x, y).FAnim - Int(map(mapper).Tile(x, y).FAnim / 14) * 14
Case 7
    GetTileX = map(mapper).Tile(x, y).Fringe2 - Int(map(mapper).Tile(x, y).Fringe2 / 14) * 14
Case 8
    GetTileX = map(mapper).Tile(x, y).F2Anim - Int(map(mapper).Tile(x, y).F2Anim / 14) * 14
End Select
End Function
Function GetTileY(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal layer As Long)
Select Case layer
Case 0
    GetTileY = Int(map(mapper).Tile(x, y).Ground / 14)
Case 1
    GetTileY = Int(map(mapper).Tile(x, y).Mask / 14)
Case 2
    GetTileY = Int(map(mapper).Tile(x, y).Anim / 14)
Case 3
    GetTileY = Int(map(mapper).Tile(x, y).Mask2 / 14)
Case 4
    GetTileY = Int(map(mapper).Tile(x, y).M2Anim / 14)
Case 5
    GetTileY = Int(map(mapper).Tile(x, y).Fringe / 14)
Case 6
    GetTileY = Int(map(mapper).Tile(x, y).FAnim / 14)
Case 7
    GetTileY = Int(map(mapper).Tile(x, y).Fringe2 / 14)
Case 8
    GetTileY = Int(map(mapper).Tile(x, y).F2Anim / 14)
End Select
End Function
Function GetTileSet(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal layer As Long)
Select Case layer
Case 0
    GetTileSet = map(mapper).Tile(x, y).GroundSet
Case 1
    GetTileSet = map(mapper).Tile(x, y).MaskSet
Case 2
    GetTileSet = map(mapper).Tile(x, y).AnimSet
Case 3
    GetTileSet = map(mapper).Tile(x, y).Mask2Set
Case 4
    GetTileSet = map(mapper).Tile(x, y).M2AnimSet
Case 5
    GetTileSet = map(mapper).Tile(x, y).FringeSet
Case 6
    GetTileSet = map(mapper).Tile(x, y).FAnimSet
Case 7
    GetTileSet = map(mapper).Tile(x, y).Fringe2Set
Case 8
    GetTileSet = map(mapper).Tile(x, y).F2AnimSet
End Select
End Function

Sub SendMap(ByVal mapper As Long)
Call SendDataToMap(mapper, "CHECKFORMAP" & SEP_CHAR & mapper & SEP_CHAR & (map(mapper).Revision + 1) & SEP_CHAR & END_CHAR)
End Sub

Sub SpellAnim(ByVal SpellNum As Long, ByVal mapper As Long, ByVal x As Long, ByVal y As Long)
Call SendDataToMap(mapper, "scriptspellanim" & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & END_CHAR)
End Sub
Sub SetAttribute(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)
Call ScriptSetAttribute(mapper, x, y, Attrib, Data1, Data2, Data3, String1, String2, String3)
End Sub
Function GetAttribute(ByVal mapper As Long, ByVal x As Long, ByVal y As Long)
    GetAttribute = map(mapper).Tile(x, y).Type
End Function
Function GetTileData1(ByVal mapper As Long, ByVal x As Long, ByVal y As Long)
    GetTileData1 = map(mapper).Tile(x, y).Data1
End Function
Function GetTileData2(ByVal mapper As Long, ByVal x As Long, ByVal y As Long)
    GetTileData2 = map(mapper).Tile(x, y).Data2
End Function
Function GetTileData3(ByVal mapper As Long, ByVal x As Long, ByVal y As Long)
    GetTileData3 = map(mapper).Tile(x, y).Data3
End Function
Function GetTileString1(ByVal mapper As Long, ByVal x As Long, ByVal y As Long)
    GetTileString1 = map(mapper).Tile(x, y).String1
End Function
Function GetTileString2(ByVal mapper As Long, ByVal x As Long, ByVal y As Long)
    GetTileString2 = map(mapper).Tile(x, y).String2
End Function
Function GetTileString3(ByVal mapper As Long, ByVal x As Long, ByVal y As Long)
    GetTileString3 = map(mapper).Tile(x, y).String3
End Function
Sub PlayerQueryBox(ByVal index As Long, ByVal Message As String, ByVal Script As Long)
    Call SendDataTo(index, "querybox" & SEP_CHAR & Message & SEP_CHAR & Script & SEP_CHAR & END_CHAR)
End Sub
Sub UpdatePaperDoll(ByVal index As Long)
    Call SendDataToMap(GetPlayerMap(index), "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
End Sub
Sub UpdateSprite(ByVal index As Long)
    Call SendDataToMap(GetPlayerMap(index), "updatesprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index))
End Sub

Sub SetMapNpcNumber(ByVal MapNum As Long, ByVal index As Long, ByVal number As Long)
    If number >= 1 Then
    MapNpc(MapNum, index).num = number
    End If
End Sub

Sub SetMapNpcTARGET(ByVal MapNum As Long, ByVal index As Long, ByVal Target As Long)
    MapNpc(MapNum, index).num = Target
End Sub

Sub SetMapNpcdir(ByVal MapNum As Long, ByVal index As Long, ByVal direction As Long)
    MapNpc(MapNum, index).Dir = direction
End Sub

Sub SetMapNpcY(ByVal MapNum As Long, ByVal index As Long, ByVal NPC_Y As Long)
    If NPC_Y <= MAX_MAPY And NPC_Y >= 0 Then
    MapNpc(MapNum, index).y = NPC_Y
    End If
End Sub

Sub SetMapNpcX(ByVal MapNum As Long, ByVal index As Long, ByVal NPC_X As Long)
    If NPC_X <= MAX_MAPX And NPC_X >= 0 Then
    MapNpc(MapNum, index).x = NPC_X
    End If
End Sub

Sub SetMapNpcHP(ByVal MapNum As Long, ByVal index As Long, ByVal HitPoints As Long)
    MapNpc(MapNum, index).HP = HitPoints
End Sub

Sub sendNPC(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim Packet
    Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub
   
Function GetNpcMaxhp(ByVal index As Long)
    GetNpcMaxhp = Npc(index).MaxHp
End Function

Function GetMapNPCnumber(ByVal MapNum As Long, ByVal index As Long)
   GetMapNPCnumber = MapNpc(MapNum, index).num
End Function

Function GetMapNpcHP(ByVal MapNum As Long, ByVal index As Long)
    GetMapNpcHP = MapNpc(MapNum, index).HP
End Function

Function GetNpcName(ByVal number As Long)
    GetNpcName = Npc(number).Name
End Function

Function GetNpcBehavior(ByVal number As Long)
    GetNpcBehavior = Npc(number).Behavior
End Function

Function GetNpcExp(ByVal number As Long)
GetNpcExp = Npc(number).Exp
End Function

Function GetNpcDefence(ByVal number As Long)
GetNpcDefence = Npc(number).DEF
End Function

Function GetNpcStrength(ByVal number As Long)
GetNpcStrength = Npc(number).STR
End Function

Sub SendIndexWornEquipment(ByVal index As Long)
Dim Packet As String
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

    If GetPlayerArmorSlot(index) > 0 Then Armor = GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))
    If GetPlayerHelmetSlot(index) > 0 Then Helmet = GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))
    If GetPlayerShieldSlot(index) > 0 Then Shield = GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))
    If GetPlayerWeaponSlot(index) > 0 Then Weapon = GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))
    If GetPlayerLegsSlot(index) > 0 Then Legs = GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))
    If GetPlayerRingSlot(index) > 0 Then Ring = GetPlayerInvItemNum(index, GetPlayerRingSlot(index))
    If GetPlayerNecklaceSlot(index) > 0 Then Necklace = GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))
   
    Packet = "itemworn" & SEP_CHAR & index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendIndexWornEquipmentFromMap(ByVal index As Long)
Dim Packet As String
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
           
                If GetPlayerArmorSlot(I) > 0 Then Armor = GetPlayerInvItemNum(I, GetPlayerArmorSlot(I))
                If GetPlayerHelmetSlot(I) > 0 Then Helmet = GetPlayerInvItemNum(I, GetPlayerHelmetSlot(I))
                If GetPlayerShieldSlot(I) > 0 Then Shield = GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))
                If GetPlayerWeaponSlot(I) > 0 Then Weapon = GetPlayerInvItemNum(I, GetPlayerWeaponSlot(I))
                If GetPlayerLegsSlot(I) > 0 Then Legs = GetPlayerInvItemNum(I, GetPlayerLegsSlot(I))
                If GetPlayerRingSlot(I) > 0 Then Ring = GetPlayerInvItemNum(I, GetPlayerRingSlot(I))
                If GetPlayerNecklaceSlot(I) > 0 Then Necklace = GetPlayerInvItemNum(I, GetPlayerNecklaceSlot(I))
               
                Packet = "itemworn" & SEP_CHAR & I & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & SEP_CHAR & END_CHAR
                Call SendDataTo(index, Packet)
                'Call SendDataTo(I, Packet)
            End If
        End If
    Next I
End Sub

Function getPlayersOnMap(ByVal MapNum As Long)
Dim I As Long, n As Long

    n = 0
    
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
            n = n + 1
        End If
    Next I
    
    getPlayersOnMap = n
End Function

Sub ScriptSpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long, ByVal spawn_x As Long, ByVal spawn_y As Long, ByVal NpcNum As Long)
'                         NPC_index               map_number          X spawn          y spawn            NPC_number
Dim Packet As String
Dim I As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum < 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    If NpcNum = 0 Then
        map(MapNum).Revision = map(MapNum).Revision + 1
        MapNpc(MapNum, MapNpcNum).num = 0
        map(MapNum).Npc(MapNpcNum) = 0
        MapNpc(MapNum, MapNpcNum).Target = 0
        MapNpc(MapNum, MapNpcNum).HP = 0
        MapNpc(MapNum, MapNpcNum).MP = 0
        MapNpc(MapNum, MapNpcNum).SP = 0
        MapNpc(MapNum, MapNpcNum).Dir = 0
        MapNpc(MapNum, MapNpcNum).x = 0
        MapNpc(MapNum, MapNpcNum).y = 0
             
        'Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(mapnum, MapNpcNum).num & SEP_CHAR & MapNpc(mapnum, MapNpcNum).x & SEP_CHAR & MapNpc(mapnum, MapNpcNum).y & SEP_CHAR & MapNpc(mapnum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(mapnum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
        'Call SendDataToMap(mapnum, Packet)
        Call SaveMap(MapNum)
    End If
    
    'MapNpc(mapnum, MapNpcNum).num = 0
    'MapNpc(mapnum, MapNpcNum).SpawnWait = GetTickCount
    'MapNpc(mapnum, MapNpcNum).HP = 0
    'Call SendDataToMap(mapnum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
    
    
        map(MapNum).Revision = map(MapNum).Revision + 1
    
        MapNpc(MapNum, MapNpcNum).num = NpcNum
        map(MapNum).Npc(MapNpcNum) = NpcNum
        
        MapNpc(MapNum, MapNpcNum).Target = 0
        
        MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxhp(NpcNum)
        MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
        MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)
                
        MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)
        
        MapNpc(MapNum, MapNpcNum).x = spawn_x
        MapNpc(MapNum, MapNpcNum).y = spawn_y
             
        Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).x & SEP_CHAR & MapNpc(MapNum, MapNpcNum).y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, Packet)
        
        Call SaveMap(MapNum)
        
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) And GetPlayerMap(I) = MapNum Then
                Call SendDataTo(I, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(I) & SEP_CHAR & map(GetPlayerMap(I)).Revision & SEP_CHAR & END_CHAR)
            End If
        Next I
        
End Sub

Sub SpawnItemOnMap(ByVal map As Long, ByVal x As Long, ByVal y As Long, ByVal ItemNum As Long, ByVal Amount As Long, ByVal durability As Long)
Dim I As Long

    If ItemNum > MAX_ITEMS Or ItemNum = 0 Then
            Exit Sub
    End If
    
    If map < 0 Or map > MAX_MAPS Then
        Exit Sub
    End If
    
    If Amount < 0 Then
        Exit Sub
    End If
    
    I = FindOpenMapItemSlot(map)
        
    If I <> 0 Then

            Call SpawnItemSlot(I, ItemNum, Amount, durability, map, x, y)
    Else
    I = MAX_MAP_ITEMS
    Call SpawnItemSlot(I, ItemNum, Amount, durability, map, x, y)
    End If
End Sub

Function getitemname(ByVal number As Long)
getitemname = Item(number).Name
End Function

Sub clearitemslot(ByVal map As Long, ByVal item_index As Long)
Dim I
I = item_index
        Call SpawnItemSlot(I, 0, 0, 0, map, MapItem(map, I).x, MapItem(map, I).y)
        'Call ClearMapItem(I, map)
        'Call SaveMap(map)
            'Packet = "MAPITEMDATA" & SEP_CHAR
    'For I = 1 To MAX_MAP_ITEMS
        'Packet = Packet & MapItem(MapNum, I).num & SEP_CHAR & MapItem(MapNum, I).Value & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & MapItem(MapNum, I).x & SEP_CHAR & MapItem(MapNum, I).y & SEP_CHAR
    'Next I
    'Packet = Packet & END_CHAR
    
    'Call SendDataToMap(MapNum, Packet)
End Sub

Sub goshopping(ByVal index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim I As Long, x As Long, y As Long, z As Long, XX As Long

    z = 0
    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For I = 1 To 7
        For XX = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(I).Value(XX).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(XX).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(XX).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(XX).GetValue & SEP_CHAR
            
            ' Item #
            x = Shop(ShopNum).TradeItem(I).Value(XX).GetItem
            
            If Item(x).Type = ITEM_TYPE_SPELL Then
                ' Spell class requirement
                y = Spell(Item(x).Data1).ClassReq
                
                If y = 0 Then
                    Call PlayerMsg(index, Trim(Item(x).Name) & " can be used by all classes.", Yellow)
                Else
                    Call PlayerMsg(index, Trim(Item(x).Name) & " can only be used by a " & GetClassName(y - 1) & ".", Yellow)
                End If
            End If
            If x < 1 Then
                z = z + 1
            End If
        Next XX
    Next I
    Packet = Packet & END_CHAR
    
    If z = (MAX_TRADES * 6) Then
        Call PlayerMsg(index, "This shop has nothing to sell!", BrightRed)
    Else
        Call SendDataTo(index, Packet)
    End If
End Sub

Sub lockplayer(ByVal index As Long, ByVal locked As Long)

If locked > 1 Or locked < 0 Then
Exit Sub
End If

If locked = 0 Then
Player(index).locked = False
Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
End If

If locked = 1 Then
Player(index).locked = True
Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
End If

End Sub

Function getplayerlocked(ByVal index As Long)

If Player(index).locked Then
getplayerlocked = 1
Else
getplayerlocked = 0
End If

End Function

Sub setservertimehour(ByVal script_hour As Long)

If script_hour >= 0 And script_hour < 24 Then
Hours = script_hour
End If

End Sub

Sub setservertimeminute(ByVal script_minute As Long)

If script_minute >= 0 And script_minute < 60 Then
Minutes = script_minute
End If

End Sub

Sub setservertimesecond(ByVal script_second As Long)

If script_second >= 0 And script_second < 60 Then
Seconds = script_second
End If

End Sub

Function getservertimehour()
getservertimehour = Hours
End Function

Function getservertimeminute()
getservertimeminute = Minutes
End Function

Function getservertimesecond()
getservertimesecond = Seconds
End Function

Function getInvItemslot(index, F_item_num, F_Min_value, F_Min_dur)
Dim F_slot

F_slot = 1
getInvItemslot = 0

Do While F_slot < 25
    If GetPlayerInvItemNum(index, F_slot) = F_item_num Then
        If GetPlayerInvItemValue(index, F_slot) >= F_Min_value Then
            If GetPlayerInvItemDur(index, F_slot) >= F_Min_dur Then
            
            getInvItemslot = F_slot
            F_slot = 25
        
            End If
        End If
    End If
F_slot = F_slot + 1
Loop

End Function

Sub giveplayeritemtoslot(index, F_slot, F_num, F_Val, F_dur)
Dim Packet

If F_num < 0 Or F_num > MAX_ITEMS Or F_Val < 0 Or F_slot < 1 Or F_slot > 24 Then
Exit Sub
End If

Player(index).Char(Player(index).CharNum).Inv(F_slot).num = F_num
Player(index).Char(Player(index).CharNum).Inv(F_slot).Value = F_Val
Player(index).Char(Player(index).CharNum).Inv(F_slot).Dur = F_dur

Packet = "PLAYERINVUPDATE" & SEP_CHAR & F_slot & SEP_CHAR & index & SEP_CHAR & GetPlayerInvItemNum(index, F_slot) & SEP_CHAR & GetPlayerInvItemValue(index, F_slot) & SEP_CHAR & GetPlayerInvItemDur(index, F_slot) & SEP_CHAR & index & SEP_CHAR & END_CHAR
Call SendDataToMap(GetPlayerMap(index), Packet)

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

Sub textbubble(player_index, bubble_index, F_text, F_map, F_x, F_y, F_colour)
Dim Packet
    
If player_index < 1 Or player_index > MAX_PLAYERS Or bubble_index < 1 Or bubble_index > 20 Or F_map < 1 Or F_map > MAX_MAPS Or F_colour < 0 Or F_colour > 15 Then
Exit Sub
End If
    
    Packet = "scriptbubble" & SEP_CHAR & bubble_index & SEP_CHAR & F_text & SEP_CHAR & F_map & SEP_CHAR & F_x & SEP_CHAR & F_y & SEP_CHAR & F_colour & SEP_CHAR & END_CHAR
    Call SendDataTo(player_index, Packet)
    
End Sub

Sub maptextbubble(bubble_index, F_text, F_map, F_x, F_y, F_colour)
Dim Packet
Dim I

If bubble_index < 1 Or bubble_index > 20 Or F_map < 1 Or F_map > MAX_MAPS Or F_colour < 0 Or F_colour > 15 Then
Exit Sub
End If
    
    Packet = "scriptbubble" & SEP_CHAR & bubble_index & SEP_CHAR & F_text & SEP_CHAR & F_map & SEP_CHAR & F_x & SEP_CHAR & F_y & SEP_CHAR & F_colour & SEP_CHAR & END_CHAR
    
For I = 1 To MAX_PLAYERS
If GetPlayerMap(I) = F_map Then
Call SendDataTo(I, Packet)
End If
Next I
    
End Sub

Function getmapname(F_map)
getmapname = map(F_map).Name
End Function

Function getmapup(F_map)
getmapup = map(F_map).Up
End Function

Function getmapdown(F_map)
getmapdown = map(F_map).Down
End Function

Function getmapleft(F_map)
getmapleft = map(F_map).left
End Function

Function getmapright(F_map)
getmapright = map(F_map).Right
End Function

Sub CustomMenuShow(ByVal index As Long, ByVal Title As String, ByVal FileName As String, ByVal closable As Long)
Dim Packet As String

If index <= 0 Or index >= MAX_PLAYERS Or closable < 0 Or closable > 1 Then
Exit Sub
End If

Packet = "showcustommenu" & SEP_CHAR & Title & SEP_CHAR & FileName & SEP_CHAR & closable & SEP_CHAR & END_CHAR
Call SendDataTo(index, Packet)
End Sub

Sub CustomMenuClose(ByVal index As Long)
Dim Packet As String

Packet = "closecustommenu" & SEP_CHAR & END_CHAR
Call SendDataTo(index, Packet)

End Sub

Sub CustomMenuPicture(ByVal player_index As Long, ByVal picture_index As Long, ByVal FileName As String, ByVal left As Long, ByVal top As Long)
Dim Packet As String

If picture_index > 39 Or picture_index < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
Exit Sub
End If

Packet = "loadpiccustommenu" & SEP_CHAR & picture_index & SEP_CHAR & FileName & SEP_CHAR & left & SEP_CHAR & top & SEP_CHAR & END_CHAR
Call SendDataTo(player_index, Packet)

End Sub

Sub CustomMenuLabel(ByVal player_index As Long, ByVal picture_index As Long, ByVal caption As String, ByVal left As Long, ByVal top As Long, ByVal customsize As Long, ByVal customcolour As Long, ByVal alignment As Long, ByVal width As Long, ByVal height As Long)
Dim Packet As String

If alignment < 0 Or alignment > 2 Or width < 0 Or height < 0 Or customcolour > 15 Or customcolour < 0 Or picture_index > 19 Or picture_index < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
Exit Sub
End If

Packet = "loadlabelcustommenu" & SEP_CHAR & picture_index & SEP_CHAR & caption & SEP_CHAR & left & SEP_CHAR & top & SEP_CHAR & customsize & SEP_CHAR & customcolour & SEP_CHAR & alignment & SEP_CHAR & width & SEP_CHAR & height & SEP_CHAR & END_CHAR
Call SendDataTo(player_index, Packet)

End Sub

Sub CustomMenuTextBox(ByVal player_index As Long, ByVal customIndex As Long, ByVal width As Long, ByVal left As Long, ByVal top As Long, ByVal text As String)
Dim Packet As String

If customIndex > 4 Or customIndex < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
Exit Sub
End If
        
Packet = "loadtextboxcustommenu" & SEP_CHAR & customIndex & SEP_CHAR & left & SEP_CHAR & width & SEP_CHAR & top & SEP_CHAR & text & SEP_CHAR & END_CHAR
Call SendDataTo(player_index, Packet)

End Sub

Function getplayermenuclicktitle(ByVal player_index As Long)
Dim msg As String

If player_index <= 0 Or player_index > MAX_PLAYERS Then
Exit Function
End If

msg = Player(player_index).custom_title

getplayermenuclicktitle = Trim(msg)

End Function

Function getplayermenuclickmsg(ByVal player_index As Long)
Dim msg As String

If player_index <= 0 Or player_index > MAX_PLAYERS Then
Exit Function
End If

msg = Player(player_index).custom_msg

getplayermenuclickmsg = Trim(msg)
End Function

Sub Loadinternet(ByVal index As Long, ByVal address As String)
Dim Packet As String

If index <= 0 Or index > MAX_PLAYERS Then
Exit Sub
End If

Packet = "loadinternetwindow" & SEP_CHAR & address & SEP_CHAR & END_CHAR
Call SendDataTo(index, Packet)

End Sub

Sub updateplayermenutext(ByVal player_index As Long, ByVal box_index As Long)
Dim Packet As String
        
        Packet = "returncustomboxmsg" & SEP_CHAR & box_index & SEP_CHAR & END_CHAR
        Call SendDataTo(player_index, Packet)
    
End Sub

Function getplayermenutext(ByVal player_index As Long)

getplayermenutext = Player(player_index).custom_msg

End Function


