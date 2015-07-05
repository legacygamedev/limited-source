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

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.
Option Explicit

Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Private Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)

Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    szReturn = vbNullString
    sSpaces = Space(5000)
    File = App.Path & "\" & File
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, Header As String, Var As String, Value As String)
    File = App.Path & "\" & File
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Packet As String

    Packet = GLOBALMSG_CHAR & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Packet As String
    Dim i As Long

    Packet = ADMINMSG_CHAR & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, Packet)
        End If

    Next i

End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Packet As String

    Packet = PLAYERMSG_CHAR & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Packet As String
    Dim text As String

    Packet = MAPMSG_CHAR & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    Call SendDataToMap(MapNum, Packet)
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Dim Packet As String

    Packet = ALERTMSG_CHAR & SEP_CHAR & Msg & END_CHAR
    Call SendDataTo(Index, Packet)
    Call CloseSocket(Index)
End Sub

' :::::::::::::::::::::::::::::::::::::::::::
' :: Functions/Subs Available To The Owner ::
' :::::::::::::::::::::::::::::::::::::::::::
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
    GetPlayerGuildAccess = Player(Index).Char(Player(Index).CharNum).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Guildaccess As Long)
    Player(Index).Char(Player(Index).CharNum).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

Function GetPlayerClassName(ByVal Index As Long) As String
    GetPlayerClassName = GetVar("Classes\Class" & GetPlayerClass(Index) & ".ini", "CLASS", "Name")
End Function

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Char(Player(Index).CharNum).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Char(Player(Index).CharNum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Char(Player(Index).CharNum).Exp = Exp
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

    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    End If

    If GetPlayerHP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).HP = 0
    End If

    Call SendStats(Index)
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    End If

    If GetPlayerMP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).MP = 0
    End If

End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    End If

    If GetPlayerSP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).SP = 0
    End If

End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim i As Long

    CharNum = Player(Index).CharNum
    GetPlayerMaxHP = (Player(Index).Char(CharNum).Level + Int(GetPlayerstr(Index) / 2) + Class(Player(Index).Char(CharNum).Class).STR) * 2
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    Dim CharNum As Long

    CharNum = Player(Index).CharNum
    GetPlayerMaxMP = (Player(Index).Char(CharNum).Level + Int(GetPlayerMAGI(Index) / 2) + Class(Player(Index).Char(CharNum).Class).Magi) * 2
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    Dim CharNum As Long

    CharNum = Player(Index).CharNum
    GetPlayerMaxSP = (Player(Index).Char(CharNum).Level + Int(GetPlayerSPEED(Index) / 2) + Class(Player(Index).Char(CharNum).Class).Speed) * 2
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

Function GetClassstr(ByVal ClassNum As Long) As Long
    GetClassstr = Class(ClassNum).STR
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

Function GetPlayerstr(ByVal Index As Long) As Long
    GetPlayerstr = Player(Index).Char(Player(Index).CharNum).STR
End Function

Sub SetPlayerstr(ByVal Index As Long, ByVal STR As Long)
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
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
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

Function GetBootMap(ByVal Index As Long) As Long
    GetBootMap = Map(GetPlayerMap(Index)).BootMap
End Function

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        i = FindOpenMapItemSlot(GetPlayerMap(Index))

        If i <> 0 Then
            MapItem(GetPlayerMap(Index), i).Dur = 0

            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type

                Case ITEM_TYPE_ARMOR

                    If InvNum = GetPlayerArmorSlot(Index) Then
                        Call SetPlayerArmorSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If

                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)

                Case ITEM_TYPE_WEAPON

                    If InvNum = GetPlayerWeaponSlot(Index) Then
                        Call SetPlayerWeaponSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If

                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)

                Case ITEM_TYPE_HELMET

                    If InvNum = GetPlayerHelmetSlot(Index) Then
                        Call SetPlayerHelmetSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If

                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)

                Case ITEM_TYPE_SHIELD

                    If InvNum = GetPlayerShieldSlot(Index) Then
                        Call SetPlayerShieldSlot(Index, 0)
                        Call SendWornEquipment(Index)
                    End If

                    MapItem(GetPlayerMap(Index), i).Dur = GetPlayerInvItemDur(Index, InvNum)
            End Select

            MapItem(GetPlayerMap(Index), i).num = GetPlayerInvItemNum(Index, InvNum)
            MapItem(GetPlayerMap(Index), i).X = GetPlayerX(Index)
            MapItem(GetPlayerMap(Index), i).Y = GetPlayerY(Index)

            If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then

                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                    MapItem(GetPlayerMap(Index), i).Value = GetPlayerInvItemValue(Index, InvNum)
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemDur(Index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(Index), i).Value = Amount
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
                End If

            Else

                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(Index), i).Value = 0

                If Item(GetPlayerInvItemNum(Index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(Index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                    If Item(GetPlayerInvItemNum(Index, InvNum)).Data1 <= -1 Then
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - Ind.", Yellow)
                    Else
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & " - " & GetPlayerInvItemDur(Index, InvNum) & "/" & Item(GetPlayerInvItemNum(Index, InvNum)).Data1 & ".", Yellow)
                    End If

                Else
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops a " & Trim$(Item(GetPlayerInvItemNum(Index, InvNum)).Name) & ".", Yellow)
                End If

                Call SetPlayerInvItemNum(Index, InvNum, 0)
                Call SetPlayerInvItemValue(Index, InvNum, 0)
                Call SetPlayerInvItemDur(Index, InvNum, 0)
            End If

            ' Send inventory update
            Call SendInventoryUpdate(Index, InvNum)

            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).num, Amount, MapItem(GetPlayerMap(Index), i).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
        Else
            Call PlayerMsg(Index, "To many items already on the ground.", BrightRed)
        End If
    End If

End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim Packet As String
    Dim i As Long

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 And ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
        MapItem(MapNum, i).num = ItemNum
        MapItem(MapNum, i).Value = ItemVal

        If ItemNum <> 0 Then
            If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Then
                MapItem(MapNum, i).Dur = ItemDur
            Else
                MapItem(MapNum, i).Dur = 0
            End If

        Else
            MapItem(MapNum, i).Dur = 0
        End If

        MapItem(MapNum, i).X = X
        MapItem(MapNum, i).Y = Y
        Packet = SPAWNITEM_CHAR & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If

End Sub

Function IsConnected(ByVal Index As Long) As Boolean

    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

    If Index < 1 Then Exit Function
    If IsConnected(Index) And Player(Index).InGame = "YES" Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If

End Function

Sub SendInventory(ByVal Index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = PLAYERINV_CHAR & SEP_CHAR & Index & SEP_CHAR

    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, i) & SEP_CHAR & GetPlayerInvItemValue(Index, i) & SEP_CHAR & GetPlayerInvItemDur(Index, i) & SEP_CHAR
    Next i

    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Dim Packet As String

    Packet = PLAYERINVUPDATE_CHAR & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & Index & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim Packet As String

    If IsPlaying(Index) Then
        Packet = PLAYERWORNEQ_CHAR & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & END_CHAR
        Call SendDataToMap(GetPlayerMap(Index), Packet)
    End If

End Sub

Sub SendHP(ByVal Index As Long)
    Dim Packet As String

    Packet = PLAYERHP_CHAR & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
    Packet = PLAYERPOINTS_CHAR & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMP(ByVal Index As Long)
    Dim Packet As String

    Packet = PLAYERMP_CHAR & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendSP(ByVal Index As Long)
    Dim Packet As String

    Packet = PLAYERSP_CHAR & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendStats(ByVal Index As Long)
    Dim Packet As String

    Packet = PLAYERSTATSPACKET_CHAR & SEP_CHAR & GetPlayerstr(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub Flash(ByVal Index As Long, ByVal flashfile As String)
    Call SendDataTo(Index, FLASHEVENT_CHAR & SEP_CHAR & flashfile & END_CHAR)
End Sub

Sub Prompt(ByVal Index As Long, ByVal question As String, ByVal Value As Long)
    Call SendDataTo(Index, PROMPT_CHAR & SEP_CHAR & question & SEP_CHAR & Value & END_CHAR)
End Sub

Sub PlaySound(ByVal Index As Long, ByVal sound As String)
    Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & sound & END_CHAR)
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim Packet As String

    ' Send index's player data to everyone including himself on the map
    Packet = PLAYERDATA_CHAR & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(Index) & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
    Dim i As Long, N As Long, startc As Long

    If IsConnected(Index) Then
        frmServer.Socket(Index).SendData Data

        DoEvents
    End If

End Sub

Sub SendDataToAll(ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If

    Next i

End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> Index Then
            Call SendDataTo(i, Data)
        End If

    Next i

End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next i

End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next i

End Sub

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
End Sub

Function GetPlayerCharNum(ByVal Index As Long) As Long
    GetPlayerCharNum = Player(Index).CharNum
End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase(Mid(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next i

    FindPlayer = 0
End Function

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, Optional sound As Boolean = True)
    Dim Packet As String
    Dim OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if there was an npc on the map the player is leaving, and if so say goodbye
    'If Trim$(Shop(ShopNum).LeaveSay) <> vbNullString Then
    'Call PlayerMsg(Index, Trim$(Shop(ShopNum).Name) & ": " & Trim$(Shop(ShopNum).LeaveSay) & vbNullString, SayColor)
    'End If
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)
    Call SendLeaveMap(Index, OldMap)
    Call UpdateGrid(OldMap, GetPlayerX(Index), GetPlayerY(Index), MapNum, X, Y)
    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    If Player(Index).Pet.Alive = YES Then
        Player(Index).Pet.MapToGo = -1
        Player(Index).Pet.XToGo = -1
        Player(Index).Pet.YToGo = -1
        Player(Index).Pet.Map = MapNum
        Player(Index).Pet.X = X
        Player(Index).Pet.Y = Y
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    Player(Index).GettingMap = YES

    'If sound Then Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & "Warp" & END_CHAR)
    If sound Then Call SendSound(Index, WARP_SOUND, SDTM)
    Call SendDataTo(Index, CHECKFORMAP_CHAR & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
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

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)

    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & ".")
        Call AddLog(GetPlayerName(Index) & " booted for " & Reason & ".", "errorlist.txt")
    Else
        Call AddLog("Avoided error in Sub HackingAttempt(" & Index & "," & Reason & ")", "errorlist.txt")
    End If

End Sub

Sub BattleMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Long)
    Call SendDataTo(Index, DAMAGEDISPLAY_CHAR & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Function Rand(ByVal High As Long, ByVal Low As Long)
    Randomize
    High = High + 1

    Do Until Rand >= Low
        Rand = Int(Rnd * High)
    Loop

End Function

Sub AttributeCreator(ByVal MapNum As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal AtType As Long, _
   ByVal Data1 As Long, _
   ByVal Data2 As Long, _
   ByVal Data3 As Long, _
   ByVal String1 As String, _
   ByVal String2 As String, _
   ByVal String3 As String, _
   ByVal Permanent As Byte, _
   ByVal AllMap As Byte, _
   ByVal Player As Long)

    'Map X Y AtType Data1 Data2 Data3 string1 string2 string3 Perminent AllMap Player
    Dim i As Long

    If Permanent <> 0 Then
        Map(MapNum).Tile(X, Y).Type = AtType
        Map(MapNum).Tile(X, Y).Data1 = Data1
        Map(MapNum).Tile(X, Y).Data2 = Data2
        Map(MapNum).Tile(X, Y).Data3 = Data3
        Map(MapNum).Tile(X, Y).String1 = String1
        Map(MapNum).Tile(X, Y).String2 = String2
        Map(MapNum).Tile(X, Y).String3 = String3
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        Call SaveMap(MapNum)

        For i = 1 To MAX_PLAYERS

            If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, CHECKFORMAP_CHAR & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & END_CHAR)
            End If

        Next

    Else

        If AllMap <> 0 Then
            Call SendDataToMap(MapNum, TEMPATTRIBUTE_CHAR & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & AtType & SEP_CHAR & Data1 & SEP_CHAR & Data2 & SEP_CHAR & Data3 & SEP_CHAR & String1 & SEP_CHAR & String2 & SEP_CHAR & String3 & END_CHAR)
        Else
            Call SendDataTo(Player, TEMPATTRIBUTE_CHAR & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & AtType & SEP_CHAR & Data1 & SEP_CHAR & Data2 & SEP_CHAR & Data3 & SEP_CHAR & String1 & SEP_CHAR & String2 & SEP_CHAR & String3 & END_CHAR)
        End If
    End If

End Sub

Sub ClearTempTiles(ByVal Index As Long)
    Call SendDataTo(Index, CLEARTEMPTILE_CHAR & END_CHAR)
End Sub

Sub CreateText(ByVal FileName As String)
    Dim f As Long

    f = FreeFile
    Open FileName For Output As #f
    Close #f
End Sub

Function ReadText(ByVal FileName As String) As String
    Dim f As Long

    f = FreeFile
    Open FileName For Input As #f
    ReadText = Input$(LOF(f), f)
    Close #f
End Function

Sub TileCreator(ByVal MapNum As Long, _
   ByVal X As Long, _
   ByVal Y As Long, _
   ByVal Pic As Long, _
   ByVal Layer As String, _
   ByVal LayerSet As Long, _
   ByVal DeleteRest As Byte, _
   ByVal Permanent As Byte, _
   ByVal AllMap As Byte, _
   ByVal Player As Long)

    'map x y pic layer layerset deleterest permanent allmap player
    Dim i As Long

    If Permanent <> 0 Then
        If DeleteRest <> 0 Then
            Map(MapNum).Tile(X, Y).Ground = 0
            Map(MapNum).Tile(X, Y).Mask = 0
            Map(MapNum).Tile(X, Y).Anim = 0
            Map(MapNum).Tile(X, Y).Mask2 = 0
            Map(MapNum).Tile(X, Y).M2Anim = 0
            Map(MapNum).Tile(X, Y).Fringe = 0
            Map(MapNum).Tile(X, Y).FAnim = 0
            Map(MapNum).Tile(X, Y).Fringe2 = 0
            Map(MapNum).Tile(X, Y).F2Anim = 0
            Map(MapNum).Tile(X, Y).GroundSet = 0
            Map(MapNum).Tile(X, Y).MaskSet = 0
            Map(MapNum).Tile(X, Y).AnimSet = 0
            Map(MapNum).Tile(X, Y).Mask2Set = 0
            Map(MapNum).Tile(X, Y).M2AnimSet = 0
            Map(MapNum).Tile(X, Y).FringeSet = 0
            Map(MapNum).Tile(X, Y).FAnimSet = 0
            Map(MapNum).Tile(X, Y).Fringe2Set = 0
            Map(MapNum).Tile(X, Y).F2AnimSet = 0
        End If

        Layer = Trim$(LCase$(Layer))

        If Layer = "ground" Then
            Map(MapNum).Tile(X, Y).Ground = Pic
            Map(MapNum).Tile(X, Y).GroundSet = LayerSet
        End If

        If Layer = "mask" Then
            Map(MapNum).Tile(X, Y).Mask = Pic
            Map(MapNum).Tile(X, Y).MaskSet = LayerSet
        End If

        If Layer = "anim" Then
            Map(MapNum).Tile(X, Y).Anim = Pic
            Map(MapNum).Tile(X, Y).AnimSet = LayerSet
        End If

        If Layer = "mask2" Then
            Map(MapNum).Tile(X, Y).Mask2 = Pic
            Map(MapNum).Tile(X, Y).Mask2Set = LayerSet
        End If

        If Layer = "m2anim" Then
            Map(MapNum).Tile(X, Y).M2Anim = Pic
            Map(MapNum).Tile(X, Y).M2AnimSet = LayerSet
        End If

        If Layer = "fringe" Then
            Map(MapNum).Tile(X, Y).Fringe = Pic
            Map(MapNum).Tile(X, Y).FringeSet = LayerSet
        End If

        If Layer = "fanim" Then
            Map(MapNum).Tile(X, Y).FAnim = Pic
            Map(MapNum).Tile(X, Y).FAnimSet = LayerSet
        End If

        If Layer = "fringe2" Then
            Map(MapNum).Tile(X, Y).Fringe2 = Pic
            Map(MapNum).Tile(X, Y).Fringe2Set = LayerSet
        End If

        If Layer = "f2anim" Then
            Map(MapNum).Tile(X, Y).Fringe2 = Pic
            Map(MapNum).Tile(X, Y).Fringe2Set = LayerSet
        End If

        Map(MapNum).Revision = Map(MapNum).Revision + 1
        Call SaveMap(MapNum)
    Else

        If AllMap <> 0 Then
            Call SendDataToMap(MapNum, TEMPTILE_CHAR & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Pic & SEP_CHAR & Layer & SEP_CHAR & DeleteRest & END_CHAR)
        Else
            Call SendDataTo(Player, TEMPTILE_CHAR & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Pic & SEP_CHAR & Layer & SEP_CHAR & DeleteRest & END_CHAR)
        End If
    End If

End Sub

Sub WriteText(ByVal FileName As String, _
   ByVal Data As String)
    Dim f As Long

    f = FreeFile
    Open FileName For Append As #f
    Print #f, Data
    Close #f
End Sub

Sub ReloadMap(ByVal MapNum As Long)
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i), False)
        End If

    Next

End Sub

Sub StartSpeech(ByVal Index As Long, ByVal SpeechNum As Long, ByVal Number As Long, ByVal NpcNum As Long)
    Call SendDataTo(Index, STARTSPEECH_CHAR & SEP_CHAR & SpeechNum & SEP_CHAR & Number & SEP_CHAR & NpcNum & END_CHAR)
End Sub
