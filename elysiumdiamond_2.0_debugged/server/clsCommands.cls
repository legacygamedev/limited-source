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
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, Header As String, Var As String, Value As String)
    File = App.Path & "\" & File
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
    
    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim I As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerAccess(I) > 0 Then
            Call SendDataTo(I, Packet)
        End If
    Next I
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
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
    GetPlayerNextLevel = Experience(GetPlayerLevel(index))
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

    CharNum = Player(index).CharNum
    GetPlayerMaxHP = (Player(index).Char(CharNum).Level + Int(GetPlayerSTR(index) / 2) + Class(Player(index).Char(CharNum).Class).STR) * 2
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
Dim CharNum As Long

    CharNum = Player(index).CharNum
    GetPlayerMaxMP = (Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).Magi) * 2
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
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Num = ItemNum
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

Function GetBootMap(ByVal index As Long) As Long
    GetBootMap = Map(GetPlayerMap(index)).BootMap
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long)
Dim I As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(index, InvNum) > 0) And (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
        I = FindOpenMapItemSlot(GetPlayerMap(index))
        
        If I <> 0 Then
            MapItem(GetPlayerMap(index), I).Dur = 0
            
            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerArmorSlot(index) Then
                        Call SetPlayerArmorSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(index) Then
                        Call SetPlayerWeaponSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(index) Then
                        Call SetPlayerHelmetSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(index) Then
                        Call SetPlayerShieldSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), I).Dur = GetPlayerInvItemDur(index, InvNum)
            End Select
                                
            MapItem(GetPlayerMap(index), I).Num = GetPlayerInvItemNum(index, InvNum)
            MapItem(GetPlayerMap(index), I).X = GetPlayerX(index)
            MapItem(GetPlayerMap(index), I).Y = GetPlayerY(index)
                        
            If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), I).Value = GetPlayerInvItemValue(index, InvNum)
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(index), I).Value = Amount
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Amount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(index), I).Value = 0
                If Item(GetPlayerInvItemNum(index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                    If Item(GetPlayerInvItemNum(index, InvNum)).Data1 <= -1 Then
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & " - Ind.", Yellow)
                    Else
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & " - " & GetPlayerInvItemDur(index, InvNum) & "/" & Item(GetPlayerInvItemNum(index, InvNum)).Data1 & ".", Yellow)
                    End If
                Else
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                End If
                
                Call SetPlayerInvItemNum(index, InvNum, 0)
                Call SetPlayerInvItemValue(index, InvNum, 0)
                Call SetPlayerInvItemDur(index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(I, MapItem(GetPlayerMap(index), I).Num, Amount, MapItem(GetPlayerMap(index), I).Dur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Else
            Call PlayerMsg(index, "To many items already on the ground.", BrightRed)
        End If
    End If
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
Dim Packet As String
Dim I As Long
    
    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    I = MapItemSlot
    
    If I <> 0 And ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
        MapItem(MapNum, I).Num = ItemNum
        MapItem(MapNum, I).Value = ItemVal
        
        If ItemNum <> 0 Then
            If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Then
                MapItem(MapNum, I).Dur = ItemDur
            Else
                MapItem(MapNum, I).Dur = 0
            End If
        Else
            MapItem(MapNum, I).Dur = 0
        End If
        
        MapItem(MapNum, I).X = X
        MapItem(MapNum, I).Y = Y
            
        Packet = "SPAWNITEM" & SEP_CHAR & I & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR
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
        Packet = "PLAYERWORNEQ" & SEP_CHAR & index & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & END_CHAR
        Call SendDataToMap(GetPlayerMap(index), Packet)
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

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
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
    Call SetPlayerX(index, X)
    Call SetPlayerY(index, Y)
                
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES

    Player(index).GettingMap = YES
    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "warp" & SEP_CHAR & END_CHAR)
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & END_CHAR)
    
    Call SendInventory(index)
    Call SendWornEquipment(index)
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

Sub BattleMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Long)
    Call SendDataTo(index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR)
End Sub

Function Rand(ByVal High As Long, ByVal Low As Long)
Randomize
High = High + 1
Do Until Rand >= Low
    Rand = Int(Rnd * High)
Loop
End Function
