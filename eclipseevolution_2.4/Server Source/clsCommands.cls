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
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Sub AddLog(ByVal text As String, ByVal FN As String)

    On Error Resume Next
  Dim FileName As String
  Dim f As Long

    If ServerLog = True Then
        FileName = App.Path & "\" & FN

        If FileExist(FN) Then
            f = FreeFile
            Open FileName For Output As #f
            Print #f, Time & ": " & text
            Close #f
        End If

        'f = FreeFile
        'Open FileName For Append As #f
        'Print #f, Time & ": " & text
        'Close #f
    End If

End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)

  Dim packet As String
  Dim i As Long

    packet = PacketID.AdminMsg & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, packet)
        End If

    Next i

End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)

  Dim packet As String

    packet = PacketID.AlertMsg & SEP_CHAR & Msg & SEP_CHAR & END_CHAR

    Call SendDataTo(index, packet)
    Call CloseSocket(index)

End Sub

Sub BattleMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Long)

    Call SendDataTo(index, PacketID.DamageDisplay & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR)

End Sub

Sub clearitemslot(ByVal Map As Long, ByVal item_index As Long)

  Dim i

    i = item_index
    Call SpawnItemSlot(i, 0, 0, 0, Map, MapItem(Map, i).X, MapItem(Map, i).Y)
    'Call ClearMapItem(I, map)
    'Call SaveMap(map)
    'Packet = PacketID.MapItemData & SEP_CHAR
    'For I = 1 To MAX_MAP_ITEMS
    'Packet = Packet & MapItem(MapNum, I).num & SEP_CHAR & MapItem(MapNum, I).Value & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & MapItem(MapNum, I).x & SEP_CHAR & MapItem(MapNum, I).y & SEP_CHAR
    'Next I
    'Packet = Packet & END_CHAR

    'Call SendDataToMap(MapNum, Packet)

End Sub

Sub ColorPlayerName(ByVal index As Long, ByVal Color As Long)

  Dim packet As String

    packet = PacketID.NameColor & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub CustomMenuClose(ByVal index As Long)

  Dim packet As String

    packet = PacketID.CloseCustomMenu & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub CustomMenuLabel(ByVal player_index As Long, ByVal picture_index As Long, ByVal caption As String, ByVal left As Long, ByVal top As Long, ByVal customsize As Long, ByVal customcolour As Long, ByVal alignment As Long, ByVal width As Long, ByVal height As Long)

  Dim packet As String

    If alignment < 0 Or alignment > 2 Or width < 0 Or height < 0 Or customcolour > 15 Or customcolour < 0 Or picture_index > 19 Or picture_index < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Sub
    End If

    packet = PacketID.LoadLabelCustomMenu & SEP_CHAR & picture_index & SEP_CHAR & caption & SEP_CHAR & left & SEP_CHAR & top & SEP_CHAR & customsize & SEP_CHAR & customcolour & SEP_CHAR & alignment & SEP_CHAR & width & SEP_CHAR & height & SEP_CHAR & END_CHAR
    Call SendDataTo(player_index, packet)

End Sub

Sub CustomMenuPicture(ByVal player_index As Long, ByVal picture_index As Long, ByVal FileName As String, ByVal left As Long, ByVal top As Long)

  Dim packet As String

    If picture_index > 39 Or picture_index < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Sub
    End If

    packet = PacketID.LoadPicCustomMenu & SEP_CHAR & picture_index & SEP_CHAR & FileName & SEP_CHAR & left & SEP_CHAR & top & SEP_CHAR & END_CHAR
    Call SendDataTo(player_index, packet)

End Sub

Sub CustomMenuShow(ByVal index As Long, ByVal Title As String, ByVal FileName As String, ByVal closable As Long)

  Dim packet As String

    If index <= 0 Or index >= MAX_PLAYERS Or closable < 0 Or closable > 1 Then
        Exit Sub
    End If

    packet = PacketID.ShowCustomMenu & SEP_CHAR & Title & SEP_CHAR & FileName & SEP_CHAR & closable & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub CustomMenuTextBox(ByVal player_index As Long, ByVal customIndex As Long, ByVal width As Long, ByVal left As Long, ByVal top As Long, ByVal text As String)

  Dim packet As String

    If customIndex > 4 Or customIndex < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Sub
    End If

    packet = PacketID.LoadTextBoxCustomMenu & SEP_CHAR & customIndex & SEP_CHAR & left & SEP_CHAR & width & SEP_CHAR & top & SEP_CHAR & text & SEP_CHAR & END_CHAR
    Call SendDataTo(player_index, packet)

End Sub

Function FindPlayer(ByVal Name As String) As Long

  Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small

            If Len(GetPlayerName(i)) >= Len(Trim(Name)) Then
                If UCase(Mid(GetPlayerName(i), 1, Len(Trim(Name)))) = UCase(Trim(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If

            End If
        End If
    Next i

    FindPlayer = 0

End Function

Sub Flash(ByVal index As Long, ByVal flashfile As String)

    Call SendDataTo(index, PacketID.FlashEvent & SEP_CHAR & flashfile & SEP_CHAR & END_CHAR)

End Sub

Function GetAttribute(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)

    GetAttribute = Map(mapper).tile(X, Y).Type

End Function

Function GetBootMap(ByVal index As Long) As Long

    GetBootMap = Map(GetPlayerMap(index)).BootMap

End Function

Function GetBootX(ByVal index As Long) As Long

    GetBootX = Map(GetPlayerMap(index)).BootX

End Function

Function GetBootY(ByVal index As Long) As Long

    GetBootY = Map(GetPlayerMap(index)).BootY

End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long

    GetClassDEF = Class(ClassNum).DEF

End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long

    GetClassMAGI = Class(ClassNum).Magi

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

Function GetClassSPEED(ByVal ClassNum As Long) As Long

    GetClassSPEED = Class(ClassNum).Speed

End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long

    GetClassSTR = Class(ClassNum).STR

End Function

Function GetIndexPlayerOnMap(ByVal Map As Long, ByVal X As Long, ByVal Y As Long)

  Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerMap(i) = Map Then

            If GetPlayerX(i) = X And GetPlayerY(i) = Y Then
                GetIndexPlayerOnMap = i
                Exit Function
            End If

        End If

    Next i

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

Function getitemname(ByVal number As Long)

    getitemname = Item(number).Name

End Function

Function getmapdown(F_map)

    getmapdown = Map(F_map).Down

End Function

Function getmapleft(F_map)

    getmapleft = Map(F_map).left

End Function

Function getmapname(F_map)

    getmapname = Map(F_map).Name

End Function

Function GetMapNpcHP(ByVal MapNum As Long, ByVal index As Long)

    GetMapNpcHP = MapNpc(MapNum, index).HP

End Function

Function GetMapNPCnumber(ByVal MapNum As Long, ByVal index As Long)

    GetMapNPCnumber = MapNpc(MapNum, index).num

End Function

Function getmapright(F_map)

    getmapright = Map(F_map).Right

End Function

Function getmapup(F_map)

    getmapup = Map(F_map).Up

End Function

Function GetMaxPlayers()

    GetMaxPlayers = MAX_PLAYERS

End Function

Function GetNpcBehavior(ByVal number As Long)

    GetNpcBehavior = Npc(number).Behavior

End Function

Function GetNpcDefence(ByVal number As Long)

    GetNpcDefence = Npc(number).DEF

End Function

Function GetNpcExp(ByVal number As Long)

    GetNpcExp = Npc(number).Exp

End Function

Function GetNpcMaxhp(ByVal index As Long)

    GetNpcMaxhp = Npc(index).MaxHp

End Function

Function GetNpcName(ByVal number As Long)

    GetNpcName = Npc(number).Name

End Function

Function GetNpcStrength(ByVal number As Long)

    GetNpcStrength = Npc(number).STR

End Function

Function GetNpcX(ByVal MapNum As Long, ByVal MapNpcNum As Long)

    If MapNpcNum < 1 Or MapNpcNum > 25 Then
     Else
        GetNpcX = MapNpc(MapNum, MapNpcNum).X
    End If

End Function

Function GetNpcY(ByVal MapNum As Long, ByVal MapNpcNum As Long)

    If MapNpcNum > 0 Then
        GetNpcY = MapNpc(MapNum, MapNpcNum).Y
    End If

End Function

Function GetPlayerAccess(ByVal index As Long) As Long

    If index <= 0 Or index >= MAX_PLAYERS Then
        Exit Function
    End If

    GetPlayerAccess = Player(index).Char(Player(index).CharNum).access

End Function

Function GetPlayerArmorSlot(ByVal index As Long) As Long

    GetPlayerArmorSlot = Player(index).Char(Player(index).CharNum).ArmorSlot

End Function

Function GetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long) As Long

    GetPlayerBankItemDur = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Dur

End Function

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long

    GetPlayerBankItemNum = Player(index).Char(Player(index).CharNum).Bank(BankSlot).num

End Function

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long

    GetPlayerBankItemValue = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Value

End Function

Function GetPlayerBody(ByVal index As Long)

    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerBody = Player(index).Char(Player(index).CharNum).body
    End If

End Function

Function GetPlayerCharNum(ByVal index As Long) As Long

    GetPlayerCharNum = Player(index).CharNum

End Function

Function GetPlayerClass(ByVal index As Long) As Long

    GetPlayerClass = Player(index).Char(Player(index).CharNum).Class

End Function

Function GetPlayerClassName(ByVal index As Long) As String

    GetPlayerClassName = GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "CLASS", "Name")

End Function

Function GetPlayerDEF(ByVal index As Long) As Long

    GetPlayerDEF = Player(index).Char(Player(index).CharNum).DEF

End Function

Function GetPlayerDir(ByVal index As Long) As Long

    GetPlayerDir = Player(index).Char(Player(index).CharNum).Dir

End Function

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

Function GetPlayerExp(ByVal index As Long) As Long

    GetPlayerExp = Player(index).Char(Player(index).CharNum).Exp

End Function

Function GetPlayerGender(ByVal index As Long) As Long

    If index <> 0 And index < MAX_PLAYERS Then
        GetPlayerGender = Val(GetVar("accounts\" & GetPlayerLogin(index) & ".ini", "CHAR" & GetPlayerCharNum(index), "Sex"))
     Else
        GetPlayerGender = 0
    End If

End Function

Function GetPlayerGuild(ByVal index As Long) As String

    GetPlayerGuild = Trim(Player(index).Char(Player(index).CharNum).Guild)

End Function

Function GetPlayerGuildAccess(ByVal index As Long) As Long

    GetPlayerGuildAccess = Player(index).Char(Player(index).CharNum).Guildaccess

End Function

Function GetPlayerHead(ByVal index As Long)

    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerHead = Player(index).Char(Player(index).CharNum).head
    End If

End Function

Function GetPlayerHelmetSlot(ByVal index As Long) As Long

    GetPlayerHelmetSlot = Player(index).Char(Player(index).CharNum).HelmetSlot

End Function

Function GetPlayerHP(ByVal index As Long) As Long

    GetPlayerHP = Player(index).Char(Player(index).CharNum).HP

End Function

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long

    GetPlayerInvItemDur = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur

End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long

    GetPlayerInvItemNum = Player(index).Char(Player(index).CharNum).Inv(InvSlot).num

End Function

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long

    GetPlayerInvItemValue = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value

End Function

Function GetPlayerIP(ByVal index As Long) As String

    GetPlayerIP = frmServer.Socket(index).RemoteHostIP

End Function

Function GetPlayerleg(ByVal index As Long)

    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerleg = Player(index).Char(Player(index).CharNum).leg
    End If

End Function

Function GetPlayerLegsSlot(ByVal index As Long) As Long

    GetPlayerLegsSlot = Player(index).Char(Player(index).CharNum).LegsSlot

End Function

Function GetPlayerLevel(ByVal index As Long) As Long

    GetPlayerLevel = Player(index).Char(Player(index).CharNum).Level

End Function

Function getplayerlocked(ByVal index As Long)

    If Player(index).locked Then
        getplayerlocked = 1
     Else
        getplayerlocked = 0
    End If

End Function


' :::::::::::::::::::::::::::::::::::::::::::
' :: Functions/Subs Available To The Owner ::
' :::::::::::::::::::::::::::::::::::::::::::
Function GetPlayerLogin(ByVal index As Long) As String

    GetPlayerLogin = Trim(Player(index).Login)

End Function

Function GetPlayerMAGI(ByVal index As Long) As Long

    GetPlayerMAGI = Player(index).Char(Player(index).CharNum).Magi

End Function

Function GetPlayerMap(ByVal index As Long) As Long

    GetPlayerMap = Player(index).Char(Player(index).CharNum).Map

End Function

Function GetPlayerMaxHP(ByVal index As Long) As Long

  Dim CharNum As Long
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

Function getplayermenuclickmsg(ByVal player_index As Long)

  Dim Msg As String

    If player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Function
    End If

    Msg = Player(player_index).custom_msg

    getplayermenuclickmsg = Trim(Msg)

End Function

Function getplayermenuclicktitle(ByVal player_index As Long)

  Dim Msg As String

    If player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Function
    End If

    Msg = Player(player_index).custom_title

    getplayermenuclicktitle = Trim(Msg)

End Function

Function getplayermenutext(ByVal player_index As Long)

    getplayermenutext = Player(player_index).custom_msg

End Function

Function GetPlayerMP(ByVal index As Long) As Long

    GetPlayerMP = Player(index).Char(Player(index).CharNum).MP

End Function

Function GetPlayerName(ByVal index As Long) As String

    GetPlayerName = Trim(Player(index).Char(Player(index).CharNum).Name)

End Function

Function GetPlayerNecklaceSlot(ByVal index As Long) As Long

    GetPlayerNecklaceSlot = Player(index).Char(Player(index).CharNum).NecklaceSlot

End Function

Function GetPlayerNextLevel(ByVal index As Long) As Long

    If GetPlayerLevel(index) <= MAX_LEVEL Then
        GetPlayerNextLevel = Experience(GetPlayerLevel(index))
    End If

End Function

Function GetPlayerPaperdoll(ByVal index As Long) As Byte

    If index < MAX_PLAYERS And index > 0 Then
        If Player(index).InGame Then
            GetPlayerPaperdoll = Player(index).Char(GetPlayerCharNum(index)).Paperdoll
        End If

    End If

End Function

Function GetPlayerPet(ByVal index As Long)

    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerPet = Player(index).pet
    End If

End Function

Function GetPlayerPK(ByVal index As Long) As Long

    GetPlayerPK = Player(index).Char(Player(index).CharNum).PK

End Function

Function GetPlayerPOINTS(ByVal index As Long) As Long

    GetPlayerPOINTS = Player(index).Char(Player(index).CharNum).POINTS

End Function

Function GetPlayerRingSlot(ByVal index As Long) As Long

    GetPlayerRingSlot = Player(index).Char(Player(index).CharNum).RingSlot

End Function

Function GetPlayerShieldSlot(ByVal index As Long) As Long

    GetPlayerShieldSlot = Player(index).Char(Player(index).CharNum).ShieldSlot

End Function

Function GetPlayerSkillExp(ByVal index As Long, ByVal skill As Long)

    On Error Resume Next

    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerSkillExp = Player(index).Char(Player(index).CharNum).SkillExp(skill)
    End If

End Function

Function GetPlayerSkillLvl(ByVal index As Long, ByVal skill As Long)

    On Error Resume Next

    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerSkillLvl = Player(index).Char(Player(index).CharNum).SkillLvl(skill)
    End If

End Function

Function getPlayersOnMap(ByVal MapNum As Long)

  Dim i As Long
  Dim n As Long

    n = 0

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If

    Next i

    getPlayersOnMap = n

End Function

Function GetPlayerSP(ByVal index As Long) As Long

    GetPlayerSP = Player(index).Char(Player(index).CharNum).SP

End Function

Function GetPlayerSPEED(ByVal index As Long) As Long

    GetPlayerSPEED = Player(index).Char(Player(index).CharNum).Speed

End Function

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long) As Long

    GetPlayerSpell = Player(index).Char(Player(index).CharNum).Spell(SpellSlot)

End Function

Function GetPlayerSprite(ByVal index As Long) As Long

    GetPlayerSprite = Player(index).Char(Player(index).CharNum).Sprite

End Function

Function GetPlayerSTR(ByVal index As Long) As Long

    GetPlayerSTR = Player(index).Char(Player(index).CharNum).STR

End Function

Function GetPlayerTarget(ByVal index As Long)

    If Player(index).TargetType = TARGET_TYPE_PLAYER Then
        GetPlayerTarget = Player(index).Target
     Else
        GetPlayerTarget = -1
    End If

End Function

Function GetPlayerTargetNpc(ByVal index As Long)

    If index > 0 Then
        GetPlayerTargetNpc = Player(index).targetnpc
    End If

End Function

Function GetPlayerWeaponSlot(ByVal index As Long) As Long

    GetPlayerWeaponSlot = Player(index).Char(Player(index).CharNum).WeaponSlot

End Function

Function GetPlayerX(ByVal index As Long) As Long

    GetPlayerX = Player(index).Char(Player(index).CharNum).X

End Function

Function GetPlayerY(ByVal index As Long) As Long

    GetPlayerY = Player(index).Char(Player(index).CharNum).Y

End Function

Function getservertimehour()

    getservertimehour = Hours

End Function

Function getservertimeminute()

    getservertimeminute = Minutes

End Function

Function getservertimesecond()

    getservertimesecond = Seconds

End Function

Function GetTileData1(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)

    GetTileData1 = Map(mapper).tile(X, Y).Data1

End Function

Function GetTileData2(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)

    GetTileData2 = Map(mapper).tile(X, Y).Data2

End Function

Function GetTileData3(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)

    GetTileData3 = Map(mapper).tile(X, Y).Data3

End Function

Function GetTileSet(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)

    Select Case layer
     Case 0
        GetTileSet = Map(mapper).tile(X, Y).GroundSet

     Case 1
        GetTileSet = Map(mapper).tile(X, Y).MaskSet

     Case 2
        GetTileSet = Map(mapper).tile(X, Y).AnimSet

     Case 3
        GetTileSet = Map(mapper).tile(X, Y).Mask2Set

     Case 4
        GetTileSet = Map(mapper).tile(X, Y).M2AnimSet

     Case 5
        GetTileSet = Map(mapper).tile(X, Y).FringeSet

     Case 6
        GetTileSet = Map(mapper).tile(X, Y).FAnimSet

     Case 7
        GetTileSet = Map(mapper).tile(X, Y).Fringe2Set

     Case 8
        GetTileSet = Map(mapper).tile(X, Y).F2AnimSet
    End Select

End Function

Function GetTileString1(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)

    GetTileString1 = Map(mapper).tile(X, Y).String1

End Function

Function GetTileString2(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)

    GetTileString2 = Map(mapper).tile(X, Y).String2

End Function

Function GetTileString3(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)

    GetTileString3 = Map(mapper).tile(X, Y).String3

End Function

Function GetTileX(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)

    Select Case layer
     Case 0
        GetTileX = Map(mapper).tile(X, Y).Ground - Int(Map(mapper).tile(X, Y).Ground / 14) * 14

     Case 1
        GetTileX = Map(mapper).tile(X, Y).Mask - Int(Map(mapper).tile(X, Y).Mask / 14) * 14

     Case 2
        GetTileX = Map(mapper).tile(X, Y).Anim - Int(Map(mapper).tile(X, Y).Anim / 14) * 14

     Case 3
        GetTileX = Map(mapper).tile(X, Y).Mask2 - Int(Map(mapper).tile(X, Y).Mask2 / 14) * 14

     Case 4
        GetTileX = Map(mapper).tile(X, Y).M2Anim - Int(Map(mapper).tile(X, Y).M2Anim / 14) * 14

     Case 5
        GetTileX = Map(mapper).tile(X, Y).Fringe - Int(Map(mapper).tile(X, Y).Fringe / 14) * 14

     Case 6
        GetTileX = Map(mapper).tile(X, Y).FAnim - Int(Map(mapper).tile(X, Y).FAnim / 14) * 14

     Case 7
        GetTileX = Map(mapper).tile(X, Y).Fringe2 - Int(Map(mapper).tile(X, Y).Fringe2 / 14) * 14

     Case 8
        GetTileX = Map(mapper).tile(X, Y).F2Anim - Int(Map(mapper).tile(X, Y).F2Anim / 14) * 14
    End Select

End Function

Function GetTileY(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)

    Select Case layer
     Case 0
        GetTileY = Int(Map(mapper).tile(X, Y).Ground / 14)

     Case 1
        GetTileY = Int(Map(mapper).tile(X, Y).Mask / 14)

     Case 2
        GetTileY = Int(Map(mapper).tile(X, Y).Anim / 14)

     Case 3
        GetTileY = Int(Map(mapper).tile(X, Y).Mask2 / 14)

     Case 4
        GetTileY = Int(Map(mapper).tile(X, Y).M2Anim / 14)

     Case 5
        GetTileY = Int(Map(mapper).tile(X, Y).Fringe / 14)

     Case 6
        GetTileY = Int(Map(mapper).tile(X, Y).FAnim / 14)

     Case 7
        GetTileY = Int(Map(mapper).tile(X, Y).Fringe2 / 14)

     Case 8
        GetTileY = Int(Map(mapper).tile(X, Y).F2Anim / 14)
    End Select

End Function

Function GetTimer(ByVal Name As String) As Long

    GetTimer = GetTimeLeft(Name)

End Function

Public Function GetVar(file As String, Header As String, Var As String) As String

  Dim sSpaces As String   ' Max string length
  Dim szReturn As String  ' Return default value if not found

    szReturn = ""

    sSpaces = Space(5000)

    file = App.Path & "\" & file

    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), file)

    GetVar = RTrim(sSpaces)
    GetVar = left(GetVar, Len(GetVar) - 1)

End Function

Function Getweather(ByVal MapNum As Long)

    Getweather = Map(MapNum).Weather

End Function

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

Sub giveplayeritemtoslot(index, F_slot, F_num, F_Val, F_dur)

  Dim packet

    If F_num < 0 Or F_num > MAX_ITEMS Or F_Val < 0 Or F_slot < 1 Or F_slot > 24 Then
        Exit Sub
    End If

    Player(index).Char(Player(index).CharNum).Inv(F_slot).num = F_num
    Player(index).Char(Player(index).CharNum).Inv(F_slot).Value = F_Val
    Player(index).Char(Player(index).CharNum).Inv(F_slot).Dur = F_dur

    packet = PacketID.PlayerInvUpdate & SEP_CHAR & F_slot & SEP_CHAR & index & SEP_CHAR & GetPlayerInvItemNum(index, F_slot) & SEP_CHAR & GetPlayerInvItemValue(index, F_slot) & SEP_CHAR & GetPlayerInvItemDur(index, F_slot) & SEP_CHAR & index & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)

End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)

  Dim packet As String

    packet = PacketID.GlobalMsg & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR

    Call SendDataToAll(packet)

End Sub

Sub GoShopping(ByVal index As Long, ByVal ShopNum As Long)

    'Go shoppin'
    Call SendTrade(index, ShopNum)

End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)

    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & ".")
    End If

End Sub

Sub HidePlayerPaperdoll(ByVal index As Long)

    If index < MAX_PLAYERS And index > 0 Then
        If Player(index).InGame Then
            Player(index).Char(GetPlayerCharNum(index)).Paperdoll = 0
        End If

    End If

End Sub

Sub Image(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal tile As Long, ByVal top As Long, ByVal height As Long, ByVal left As Long, ByVal width As Long)

  Dim packet As String

    packet = PacketID.Fog & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & tile & SEP_CHAR & top & SEP_CHAR & height & SEP_CHAR & left & SEP_CHAR & width & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

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

Function IsScrolling()

    If IS_SCROLLING = 1 Then
        IsScrolling = 1
     Else
        IsScrolling = 0
    End If

End Function

Sub Lights(ByVal Map As Long, ByVal light As Long)

  Dim packet As String

    packet = PacketID.Lights & SEP_CHAR & Map & SEP_CHAR & light & SEP_CHAR & END_CHAR
    Call SendDataToMap(Map, packet)

End Sub

Sub Loadinternet(ByVal index As Long, ByVal address As String)

  Dim packet As String

    If index <= 0 Or index > MAX_PLAYERS Then
        Exit Sub
    End If

    packet = PacketID.LoadInternetWindow & SEP_CHAR & address & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub lockattack(ByVal index As Long, ByVal locked As Long)

    If locked > 1 Or locked < 0 Then
        Exit Sub
    End If

    If locked = 0 Then
        Player(index).lockedattack = False
        Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
    End If

    If locked = 1 Then
        Player(index).lockedattack = True
        Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
    End If

End Sub

Sub lockitems(ByVal index As Long, ByVal locked As Long)

    If locked > 1 Or locked < 0 Then
        Exit Sub
    End If

    If locked = 0 Then
        Player(index).lockeditems = False
        Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
    End If

    If locked = 1 Then
        Player(index).lockeditems = True
        Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
    End If

End Sub

Sub lockplayer(ByVal index As Long, ByVal locked As Long)

    If locked > 1 Or locked < 0 Then
        Exit Sub
    End If

    If locked = 0 Then
        Player(index).locked = False
        Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
    End If

    If locked = 1 Then
        Player(index).locked = True
        Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
    End If

End Sub

Sub LockSpells(ByVal index As Long, ByVal locked As Long)

    If locked > 1 Or locked < 0 Then
        Exit Sub
    End If

    If locked = 0 Then
        Player(index).lockedspells = False
        Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
    End If

    If locked = 1 Then
        Player(index).lockedspells = True
        Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)
    End If

End Sub

Sub MakeDay()

    GameTime = TIME_DAY

End Sub

Sub MakeNight()

    GameTime = TIME_NIGHT

End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)

  Dim packet As String

    packet = PacketID.MapMsg & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR

    Call SendDataToMap(MapNum, packet)

End Sub

Sub maptextbubble(bubble_index, F_text, F_map, F_x, F_y, F_colour)

  Dim packet
  Dim i

    If bubble_index < 1 Or bubble_index > 20 Or F_map < 1 Or F_map > MAX_MAPS Or F_colour < 0 Or F_colour > 15 Then
        Exit Sub
    End If

    packet = PacketID.ScriptBubble & SEP_CHAR & bubble_index & SEP_CHAR & F_text & SEP_CHAR & F_map & SEP_CHAR & F_x & SEP_CHAR & F_y & SEP_CHAR & F_colour & SEP_CHAR & END_CHAR

    For i = 1 To MAX_PLAYERS

        If GetPlayerMap(i) = F_map Then
            Call SendDataTo(i, packet)
        End If

    Next i

End Sub

Sub MovePlayer(ByVal index As Long, ByVal direction As Long, ByVal Movement As Long)

    If index <> 0 And index < MAX_PLAYERS Then Call PlayerMove(index, direction, Movement)

End Sub

Sub MutePlayer(ByVal index As Integer)

    ' Check for subscript out of range

    If index <> 0 And index < MAX_PLAYERS Then
        Player(index).Mute = True
    End If

End Sub

Sub npcmoving(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal direction As Long, ByVal Speed As Long)

    If CanNpcMove(MapNum, MapNpcNum, direction) Then
        Call NPCMove(MapNum, MapNpcNum, direction, Speed)
    End If

End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)

  Dim packet As String

    packet = PacketID.PlayerMsg & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR

    Call SendDataTo(index, packet)

End Sub

Sub PlayerQueryBox(ByVal index As Long, ByVal Message As String, ByVal Script As Long)

    Call SendDataTo(index, PacketID.QueryBox & SEP_CHAR & Message & SEP_CHAR & Script & SEP_CHAR & END_CHAR)

End Sub

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)

  Dim packet As String
  Dim OldMap As Long
  Dim npcn As Long

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
    Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "warp" & SEP_CHAR & END_CHAR)
    Call SendDataTo(index, PacketID.CheckForMap & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & END_CHAR)

    Call SendInventory(index)
    'Call SendWornEquipment(index)
    Call SendIndexWornEquipmentFromMap(index)

    If Scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnMapLoad " & index
    End If

    packet = PacketID.ForceHouseClose & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

    If Player(index).pet <> 0 Then

        Do While npcn < MAX_MAP_NPCS

            If 0 + Map(MapNum).Npc(npcn + 1) = 0 Then
                Call ScriptSpawnNpc(npcn + 1, MapNum, GetPlayerX(index), GetPlayerY(index), Player(index).pet)
                MapNpc(MapNum, npcn + 1).owner = index
                npcn = MAX_MAP_NPCS
             Else
                npcn = npcn + 1
            End If

        Loop
    End If

End Sub

Sub PlaySound(ByVal index As Long, ByVal Sound As String)

    Call SendDataTo(index, PacketID.Sound & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & SEP_CHAR & END_CHAR)

End Sub

Sub PlaySoundToMap(ByVal index As Long, ByVal Sound As String)

    Call SendDataToMap(GetPlayerMap(index), PacketID.Sound & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & SEP_CHAR & END_CHAR)

End Sub

Sub Prompt(ByVal index As Long, ByVal question As String, ByVal Value As Long)

    Call SendDataTo(index, PacketID.Prompt & SEP_CHAR & question & SEP_CHAR & Value & SEP_CHAR & END_CHAR)

End Sub

Sub PutVar(file As String, Header As String, Var As String, Value As String)

    file = App.Path & "\" & file
    Call WritePrivateProfileString(Header, Var, Value, file)

End Sub

Public Function Rand(ByVal Low As Long, _
                     ByVal High As Long) As Long

    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low

End Function

Sub reloadshop(ByVal ShopNum As Long)

  Dim Max
  Dim c
  Dim i

    Max = 2
    For c = 1 To 7

        For i = 1 To Max

            If ReadINI("SHOP" & ShopNum, "GetItem" & i, App.Path & "\Shop.ini") <> vbNullString Then
                'Shop(ShopNum).TradeItem(c).Value(I).GetItem = 0 + ReadINI("SHOP" & ShopNum, "GetItem" & I, App.Path & "\Shop.ini")
                'Shop(ShopNum).TradeItem(c).Value(I).GetValue = 0 + ReadINI("SHOP" & ShopNum, "GetValue" & I, App.Path & "\Shop.ini")
                'Shop(ShopNum).TradeItem(c).Value(I).GiveItem = 0 + ReadINI("SHOP" & ShopNum, "GiveItem" & I, App.Path & "\Shop.ini")
                'Shop(ShopNum).TradeItem(c).Value(I).GiveValue = 0 + ReadINI("SHOP" & ShopNum, "GiveValue" & I, App.Path & "\Shop.ini")
                Max = Max + 1
             Else
                i = Max
            End If

        Next i

    Next c

End Sub

Sub RemoveTimer(ByVal Name As String)

    Call GetRidOfTimer(Name)

End Sub

Sub SavePlayerProgress(ByVal index As Long)

    If index <> 0 And index < MAX_PLAYERS Then Call SavePlayer(index)

End Sub

Sub Script(ByVal file As Long, ByVal suby As Long, Optional code1 As Long, Optional code2 As Long, Optional code3 As Long, Optional code4 As Long, Optional code5 As Long)

    On Error GoTo Err
    MyScript.ReadInCode App.Path & "\Scripts\" & file, "Scripts\" & file, MyScript.SControl, False

    If code1 <> "" Then
        If code2 <> "" Then
            If code3 <> "" Then
                If code4 <> "" Then
                    If 0 + code5 <> "" Then
                        MyScript.ExecuteStatement "Scripts\" & file, suby & " " & code1 & "," & code2 & "," & code3 & "," & code4 & "," & code5
                     Else
                        MyScript.ExecuteStatement "Scripts\" & file, suby & " " & code1 & "," & code2 & "," & code3 & "," & code4
                    End If

                 Else
                    MyScript.ExecuteStatement "Scripts\" & file, suby & " " & code1 & "," & code2 & "," & code3
                End If

             Else
                MyScript.ExecuteStatement "Scripts\" & file, suby & " " & code1 & "," & code2
            End If

         Else
            MyScript.ExecuteStatement "Scripts\" & file, suby & " " & code1
        End If

     Else
        MyScript.ExecuteStatement "Scripts\" & file, suby & " "
    End If

    Exit Sub
Err:
    Call AddText(frmServer.txtText(0), "Error executing Script statement at file " & file & ", sub " & suby, 15)

End Sub

Sub ScriptSpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long, ByVal spawn_x As Long, ByVal spawn_y As Long, ByVal NpcNum As Long)

    '                         NPC_index               map_number          X spawn          y spawn            NPC_number
  Dim packet As String
  Dim i As Long

    ' Check for subscript out of range

    If MapNpcNum < 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    If NpcNum = 0 Then
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        MapNpc(MapNum, MapNpcNum).num = 0
        Map(MapNum).Npc(MapNpcNum) = 0
        MapNpc(MapNum, MapNpcNum).Target = 0
        MapNpc(MapNum, MapNpcNum).HP = 0
        MapNpc(MapNum, MapNpcNum).MP = 0
        MapNpc(MapNum, MapNpcNum).SP = 0
        MapNpc(MapNum, MapNpcNum).Dir = 0
        MapNpc(MapNum, MapNpcNum).X = 0
        MapNpc(MapNum, MapNpcNum).Y = 0

        'Packet = PacketID.SpawnNPC & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(mapnum, MapNpcNum).num & SEP_CHAR & MapNpc(mapnum, MapNpcNum).x & SEP_CHAR & MapNpc(mapnum, MapNpcNum).y & SEP_CHAR & MapNpc(mapnum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(mapnum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
        'Call SendDataToMap(mapnum, Packet)
        Call SaveMap(MapNum)
    End If

    'MapNpc(mapnum, MapNpcNum).num = 0
    'MapNpc(mapnum, MapNpcNum).SpawnWait = GetTickCount
    'MapNpc(mapnum, MapNpcNum).HP = 0
    'Call SendDataToMap(mapnum, PacketID.NPCDead & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)

    Map(MapNum).Revision = Map(MapNum).Revision + 1

    MapNpc(MapNum, MapNpcNum).num = NpcNum
    Map(MapNum).Npc(MapNpcNum) = NpcNum

    MapNpc(MapNum, MapNpcNum).Target = 0

    MapNpc(MapNum, MapNpcNum).HP = GetNpcMaxhp(NpcNum)
    MapNpc(MapNum, MapNpcNum).MP = GetNpcMaxMP(NpcNum)
    MapNpc(MapNum, MapNpcNum).SP = GetNpcMaxSP(NpcNum)

    MapNpc(MapNum, MapNpcNum).Dir = Int(Rnd * 4)

    MapNpc(MapNum, MapNpcNum).X = spawn_x
    MapNpc(MapNum, MapNpcNum).Y = spawn_y

    packet = PacketID.SpawnNPC & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, packet)

    Call SaveMap(MapNum)

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call SendDataTo(i, PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & SEP_CHAR & END_CHAR)
        End If

    Next i

End Sub

Sub SendDataTo(ByVal index As Long, ByVal Data As String)

    If IsConnected(index) Then
        DebugTrackPacket Data
        frmServer.Socket(index).SendData Data
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

Sub SendDataToAllBut(ByVal index As Long, ByVal Data As String)

  Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> index Then
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

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByVal Data As String)

  Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> index Then
                Call SendDataTo(i, Data)
            End If

        End If
    Next i

End Sub

Sub SendHP(ByVal index As Long)

  Dim packet As String

    packet = PacketID.PlayerHP & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

    packet = PacketID.PlayerPoints & SEP_CHAR & GetPlayerPOINTS(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

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

    If GetPlayerArmorSlot(index) > 0 Then Armor = GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))
    If GetPlayerHelmetSlot(index) > 0 Then Helmet = GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))
    If GetPlayerShieldSlot(index) > 0 Then Shield = GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))
    If GetPlayerWeaponSlot(index) > 0 Then Weapon = GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))
    If GetPlayerLegsSlot(index) > 0 Then Legs = GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))
    If GetPlayerRingSlot(index) > 0 Then Ring = GetPlayerInvItemNum(index, GetPlayerRingSlot(index))
    If GetPlayerNecklaceSlot(index) > 0 Then Necklace = GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))

    packet = PacketID.ItemWorn & SEP_CHAR & index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)

End Sub

Sub SendIndexWornEquipmentFromMap(ByVal index As Long)

  Dim packet As String
  Dim i As Long
  Dim Armor As Long
  Dim Helmet As Long
  Dim Shield As Long
  Dim Weapon As Long
  Dim Legs As Long
  Dim Ring As Long
  Dim Necklace As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) = True Then
            If GetPlayerMap(index) = GetPlayerMap(i) Then
                Armor = 0
                Helmet = 0
                Shield = 0
                Weapon = 0
                Legs = 0
                Ring = 0
                Necklace = 0

                If GetPlayerArmorSlot(i) > 0 Then Armor = GetPlayerInvItemNum(i, GetPlayerArmorSlot(i))
                If GetPlayerHelmetSlot(i) > 0 Then Helmet = GetPlayerInvItemNum(i, GetPlayerHelmetSlot(i))
                If GetPlayerShieldSlot(i) > 0 Then Shield = GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))
                If GetPlayerWeaponSlot(i) > 0 Then Weapon = GetPlayerInvItemNum(i, GetPlayerWeaponSlot(i))
                If GetPlayerLegsSlot(i) > 0 Then Legs = GetPlayerInvItemNum(i, GetPlayerLegsSlot(i))
                If GetPlayerRingSlot(i) > 0 Then Ring = GetPlayerInvItemNum(i, GetPlayerRingSlot(i))
                If GetPlayerNecklaceSlot(i) > 0 Then Necklace = GetPlayerInvItemNum(i, GetPlayerNecklaceSlot(i))

                packet = PacketID.ItemWorn & SEP_CHAR & index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & SEP_CHAR & END_CHAR
                Call SendDataTo(index, packet)
            End If

        End If
    Next i

End Sub

Sub SendInventory(ByVal index As Long)

  Dim packet As String
  Dim i As Long

    packet = PacketID.PlayerInv & SEP_CHAR & index & SEP_CHAR

    For i = 1 To MAX_INV
        packet = packet & GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR
    Next i

    packet = packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(index), packet)

End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)

  Dim packet As String

    packet = PacketID.PlayerInvUpdate & SEP_CHAR & InvSlot & SEP_CHAR & index & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & index & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)

End Sub

Sub SendMap(ByVal mapper As Long)

    Call SendDataToMap(mapper, PacketID.CheckForMap & SEP_CHAR & mapper & SEP_CHAR & (Map(mapper).Revision + 1) & SEP_CHAR & END_CHAR)

End Sub

Sub SendMP(ByVal index As Long)

  Dim packet As String

    packet = PacketID.PlayerMP & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub sendNPC(ByVal MapNum As Long, ByVal MapNpcNum As Long)

  Dim packet

    packet = PacketID.SpawnNPC & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).num & SEP_CHAR & MapNpc(MapNum, MapNpcNum).X & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Y & SEP_CHAR & MapNpc(MapNum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(MapNum, MapNpcNum).num).Big & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, packet)

End Sub

Sub SendPlayerData(ByVal index As Long)

  Dim packet As String
  Dim j As Long

    ' Send index's player data to everyone including himself on th emap
    packet = PacketID.PlayerData & SEP_CHAR
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

    For j = 1 To MAX_SKILLS
        packet = packet & GetPlayerSkillLvl(j, index) & SEP_CHAR
        packet = packet & GetPlayerSkillExp(j, index) & SEP_CHAR
    Next j

    packet = packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), packet)

End Sub

Sub SendSP(ByVal index As Long)

  Dim packet As String

    packet = PacketID.PlayerSP & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendStats(ByVal index As Long)

  Dim packet As String

    packet = PacketID.PlayerStatsPacket & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & GetPlayerNextLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)

End Sub

Sub SendWhosOnline(ByVal index As Integer)

  Dim s As String
  Dim n As Long
  Dim i As Long

    s = ""
    n = 0

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And i <> index Then
            s = s & GetPlayerName(i) & ", "
            n = n + 1
        End If

    Next i

    If n = 0 Then
        s = "There are no other players online."
     Else
        s = Mid(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(index, s, WhoColor)

End Sub

Sub SendWornEquipment(ByVal index As Long)

  Dim packet As String

    If IsPlaying(index) Then
        packet = PacketID.PlayerWornEQ & SEP_CHAR & index & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & GetPlayerLegsSlot(index) & SEP_CHAR & GetPlayerRingSlot(index) & SEP_CHAR & GetPlayerNecklaceSlot(index) & SEP_CHAR & END_CHAR
        Call SendDataToMap(GetPlayerMap(index), packet)
    End If

End Sub

Sub SetAttribute(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)

    Call ScriptSetAttribute(mapper, X, Y, Attrib, Data1, Data2, Data3, String1, String2, String3)

End Sub

Sub SetMapNpcdir(ByVal MapNum As Long, ByVal index As Long, ByVal direction As Long)

    MapNpc(MapNum, index).Dir = direction

End Sub

Sub SetMapNpcHP(ByVal MapNum As Long, ByVal index As Long, ByVal HitPoints As Long)

    MapNpc(MapNum, index).HP = HitPoints

End Sub

Sub SetMapNpcNumber(ByVal MapNum As Long, ByVal index As Long, ByVal number As Long)

    If number >= 1 Then
        MapNpc(MapNum, index).num = number
    End If

End Sub

Sub SetMapNpcTARGET(ByVal MapNum As Long, ByVal index As Long, ByVal Target As Long)

    MapNpc(MapNum, index).num = Target

End Sub

Sub SetMapNpcX(ByVal MapNum As Long, ByVal index As Long, ByVal NPC_X As Long)

    If NPC_X <= MAX_MAPX And NPC_X >= 0 Then
        MapNpc(MapNum, index).X = NPC_X
    End If

End Sub

Sub SetMapNpcY(ByVal MapNum As Long, ByVal index As Long, ByVal NPC_Y As Long)

    If NPC_Y <= MAX_MAPY And NPC_Y >= 0 Then
        MapNpc(MapNum, index).Y = NPC_Y
    End If

End Sub

Sub SetPlayerAccess(ByVal index As Long, ByVal access As Long)

    Player(index).Char(Player(index).CharNum).access = access

End Sub

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)

    Player(index).Char(Player(index).CharNum).ArmorSlot = InvNum

End Sub

Sub SetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)

    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Dur = ItemDur

End Sub

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)

    Player(index).Char(Player(index).CharNum).Bank(BankSlot).num = ItemNum
    Call SendBankUpdate(index, BankSlot)

End Sub

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)

    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Value = ItemValue
    Call SendBankUpdate(index, BankSlot)

End Sub

Sub SetPlayerBody(ByVal index As Long, ByVal body As Long)

    If index > 0 And index < MAX_PLAYERS Then
        Player(index).Char(Player(index).CharNum).body = body
    End If

End Sub

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)

    Player(index).Char(Player(index).CharNum).Class = ClassNum

End Sub

Sub SetPlayerDEF(ByVal index As Long, ByVal DEF As Long)

    Player(index).Char(Player(index).CharNum).DEF = DEF

End Sub

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)

    Player(index).Char(Player(index).CharNum).Dir = Dir

End Sub

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)

    Player(index).Char(Player(index).CharNum).Exp = Exp

End Sub

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal Guildaccess As Long)

    Player(index).Char(Player(index).CharNum).Guildaccess = Guildaccess

End Sub

Sub setplayerguildname(ByVal index As Long, ByVal guildname As String)

    Player(index).Char(Player(index).CharNum).Guild = Trim(guildname)

End Sub

Sub SetPlayerHead(ByVal index As Long, ByVal head As Long)

    If index > 0 And index < MAX_PLAYERS Then
        Player(index).Char(Player(index).CharNum).head = head
    End If

End Sub

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)

    Player(index).Char(Player(index).CharNum).HelmetSlot = InvNum

End Sub

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

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)

    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur = ItemDur

End Sub

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)

    Player(index).Char(Player(index).CharNum).Inv(InvSlot).num = ItemNum

End Sub

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)

    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value = ItemValue

End Sub

Sub SetPlayerLeg(ByVal index As Long, ByVal leg As Long)

    If index > 0 And index < MAX_PLAYERS Then
        Player(index).Char(Player(index).CharNum).leg = leg
    End If

End Sub

Sub SetPlayerLegsSlot(ByVal index As Long, InvNum As Long)

    Player(index).Char(Player(index).CharNum).LegsSlot = InvNum

End Sub

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)

    Player(index).Char(Player(index).CharNum).Level = Level

End Sub

Sub SetPlayerMAGI(ByVal index As Long, ByVal Magi As Long)

    Player(index).Char(Player(index).CharNum).Magi = Magi

End Sub

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(index).Char(Player(index).CharNum).Map = MapNum
    End If

End Sub

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)

    Player(index).Char(Player(index).CharNum).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        Player(index).Char(Player(index).CharNum).MP = GetPlayerMaxMP(index)
    End If

    If GetPlayerMP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).MP = 0
    End If

End Sub

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)

    Player(index).Char(Player(index).CharNum).Name = Name

End Sub

Sub SetPlayerNecklaceSlot(ByVal index As Long, InvNum As Long)

    Player(index).Char(Player(index).CharNum).NecklaceSlot = InvNum

End Sub

Sub SetPlayerPet(ByVal index As Long, ByVal NpcNum As Long)

  Dim npcn As Long

    npcn = 1
    Do While npcn <= MAX_MAP_NPCS

        If 0 + MapNpc(GetPlayerMap(index), npcn).num = 0 Then
            Call ScriptSpawnNpc(npcn, GetPlayerMap(index), GetPlayerX(index) - 1, GetPlayerY(index), NpcNum)
            MapNpc(GetPlayerMap(index), NpcNum).owner = index
            Player(index).pet = NpcNum
            Exit Sub
         Else
            npcn = npcn + 1
        End If

    Loop

End Sub

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)

    Player(index).Char(Player(index).CharNum).PK = PK

End Sub

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)

    Player(index).Char(Player(index).CharNum).POINTS = POINTS

End Sub

Sub SetPlayerRingSlot(ByVal index As Long, InvNum As Long)

    Player(index).Char(Player(index).CharNum).RingSlot = InvNum

End Sub

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)

    Player(index).Char(Player(index).CharNum).ShieldSlot = InvNum

End Sub

Sub SetPlayerSkillExp(ByVal index As Long, ByVal skill As Long, ByVal lvl As Long)

    If index > 0 And index < MAX_PLAYERS Then
        Player(index).Char(Player(index).CharNum).SkillExp(skill) = lvl
    End If

End Sub

Sub SetPlayerSkillLvl(ByVal index As Long, ByVal skill As Long, ByVal lvl As Long)

    If index > 0 And index < MAX_PLAYERS Then
        Player(index).Char(Player(index).CharNum).SkillLvl(skill) = lvl
    End If

End Sub

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)

    Player(index).Char(Player(index).CharNum).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        Player(index).Char(Player(index).CharNum).SP = GetPlayerMaxSP(index)
    End If

    If GetPlayerSP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).SP = 0
    End If

End Sub

Sub SetPlayerSPEED(ByVal index As Long, ByVal Speed As Long)

    Player(index).Char(Player(index).CharNum).Speed = Speed

End Sub

Sub SetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)

    Player(index).Char(Player(index).CharNum).Spell(SpellSlot) = SpellNum

End Sub

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)

    If index > 0 And index <= MAX_PLAYERS Then
        Player(index).Char(Player(index).CharNum).Sprite = Sprite
    End If

End Sub

Sub SetPlayerSTR(ByVal index As Long, ByVal STR As Long)

    Player(index).Char(Player(index).CharNum).STR = STR

End Sub

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)

    Player(index).Char(Player(index).CharNum).WeaponSlot = InvNum

End Sub

Sub SetPlayerX(ByVal index As Long, ByVal X As Long)

    Player(index).Char(Player(index).CharNum).X = X

End Sub

Sub SetPlayerY(ByVal index As Long, ByVal Y As Long)

    Player(index).Char(Player(index).CharNum).Y = Y

End Sub

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

Sub SetSpeed(ByVal index As Long, ByVal Movement As String, ByVal Speed As Long)

  Dim packet As String
  Dim div As Integer 'Divisor for power of 2 loop

    'Prevent ridiculous inputs that would crash

    If Speed < 1 Then
        Call PlayerMsg(index, "Automated message: script error, value " & Speed & " at SetSpeed is too low - must be above 1.", Red)
     ElseIf Speed > 32 Then
        Call PlayerMsg(index, "Automated message: script error, value " & Speed & " at SetSpeed is to high - must be below 32.", Red)
     Else
        'Make sure it's a power of 2 (we have to prevent errors, especially this one)
        div = 2

        Do While div <= 32

            If Not (Speed / div) = 1 Then
                If (Speed / div) Mod 2 = 0 Then
                    div = div * 2
                 Else
                    Exit Sub 'NOT a power of 2!
                End If

             Else
                Exit Do 'WAS a power of 2!
            End If

        Loop
        'Okey dokey sends the packet plz :D
        packet = PacketID.SetSpeed & SEP_CHAR & Movement & SEP_CHAR & Speed & SEP_CHAR & END_CHAR
        Call SendDataTo(index, packet)
    End If

End Sub

Sub SetTile(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal xset As Long, ByVal yset As Long, ByVal tileset As Long, ByVal layer As Long)

    Call ScriptSetTile(mapper, X, Y, xset, yset, tileset, layer)

End Sub

Sub SetTimer(ByVal Name As String, ByVal Interval As Long)

    Call AddNewTimer(Name, Interval)

End Sub

Sub setweather(ByVal MapNum As Long, ByVal Weather As Long, ByVal Interval As Long)

  Dim packet As String

    packet = PacketID.MapWeather & SEP_CHAR & MapNum & SEP_CHAR & Weather & SEP_CHAR & Interval & SEP_CHAR & END_CHAR
    Call SendDataToMap(MapNum, packet)

End Sub

Sub ShowPlayerPaperdoll(ByVal index As Long)

    If index < MAX_PLAYERS And index > 0 Then
        If Player(index).InGame Then
            Player(index).Char(GetPlayerCharNum(index)).Paperdoll = 1
        End If

    End If

End Sub

Sub SpawnItemOnMap(ByVal Map As Long, ByVal X As Long, ByVal Y As Long, ByVal ItemNum As Long, ByVal Amount As Long, ByVal durability As Long)

  Dim i As Long

    If ItemNum > MAX_ITEMS Or ItemNum = 0 Then
        Exit Sub
    End If

    If Map < 0 Or Map > MAX_MAPS Then
        Exit Sub
    End If

    If Amount < 0 Then
        Exit Sub
    End If

    i = FindOpenMapItemSlot(Map)

    If i <> 0 Then

        Call SpawnItemSlot(i, ItemNum, Amount, durability, Map, X, Y)
     Else
        i = MAX_MAP_ITEMS
        Call SpawnItemSlot(i, ItemNum, Amount, durability, Map, X, Y)
    End If

End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)

  Dim packet As String
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
            If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_NECKLACE) Then
                MapItem(MapNum, i).Dur = ItemDur
             Else
                MapItem(MapNum, i).Dur = 0
            End If

         Else
            MapItem(MapNum, i).Dur = 0
        End If

        MapItem(MapNum, i).X = X
        MapItem(MapNum, i).Y = Y

        packet = PacketID.SpawnItem & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, packet)
    End If

End Sub

Sub SpellAnim(ByVal SpellNum As Long, ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)

    Call SendDataToMap(mapper, PacketID.ScriptSpellAnim & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & END_CHAR)

End Sub

Sub textbubble(player_index, bubble_index, F_text, F_map, F_x, F_y, F_colour)

  Dim packet

    If player_index < 1 Or player_index > MAX_PLAYERS Or bubble_index < 1 Or bubble_index > 20 Or F_map < 1 Or F_map > MAX_MAPS Or F_colour < 0 Or F_colour > 15 Then
        Exit Sub
    End If

    packet = PacketID.ScriptBubble & SEP_CHAR & bubble_index & SEP_CHAR & F_text & SEP_CHAR & F_map & SEP_CHAR & F_x & SEP_CHAR & F_y & SEP_CHAR & F_colour & SEP_CHAR & END_CHAR
    Call SendDataTo(player_index, packet)

End Sub

Sub UnMutePlayer(ByVal index As Integer)

    ' Check for subscript out of range

    If index <> 0 And index < MAX_PLAYERS Then
        Player(index).Mute = False
    End If

End Sub

Sub UpdatePaperDoll(ByVal index As Long)

    Call SendDataToMap(GetPlayerMap(index), PacketID.CheckForMap & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & SEP_CHAR & END_CHAR)

End Sub

Sub updateplayermenutext(ByVal player_index As Long, ByVal box_index As Long)

  Dim packet As String

    packet = PacketID.ReturnCustomBoxMsg & SEP_CHAR & box_index & SEP_CHAR & END_CHAR
    Call SendDataTo(player_index, packet)

End Sub

Sub UpdateSprite(ByVal index As Long)

    Call SendDataToMap(GetPlayerMap(index), PacketID.UpdateSprite & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index))

End Sub

Public Function VarExists(file As String, Header As String, Var As String) As Boolean

  Dim tmp

    tmp = GetVar(file, Header, Var)
    tmp = tmp & ""

    If tmp = "" Then
        VarExists = False
     Else
        VarExists = True
    End If

End Function

