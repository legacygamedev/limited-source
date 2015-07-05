Attribute VB_Name = "modDatabase"
Option Explicit

Sub SaveLocalMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Map(MapNum)
    Close #f
End Sub

Sub LoadMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"

    If FileExists("maps\map" & MapNum & ".dat") = False Then
        Exit Sub
    End If
    f = FreeFile
    Open FileName For Binary As #f
    Get #f, , Map(MapNum)
    Close #f
End Sub

Function GetMapRevision(ByVal MapNum As Long) As Long
    GetMapRevision = Map(MapNum).Revision
End Function

Sub ClearTempTile()
    Dim X As Long, y As Long

    For y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            TempTile(X, y).DoorOpen = NO
        Next X
    Next y
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long
    Dim n As Long

    Player(Index).Name = vbNullString
    Player(Index).Guild = vbNullString
    Player(Index).Guildaccess = 0
    Player(Index).Class = 0
    Player(Index).Level = 0
    Player(Index).Sprite = 0
    Player(Index).Exp = 0
    Player(Index).Access = 0
    Player(Index).PK = NO

    Player(Index).HP = 0
    Player(Index).MP = 0
    Player(Index).SP = 0

    Player(Index).STR = 0
    Player(Index).DEF = 0
    Player(Index).speed = 0
    Player(Index).MAGI = 0

    For n = 1 To MAX_INV
        Player(Index).Inv(n).num = 0
        Player(Index).Inv(n).Value = 0
        Player(Index).Inv(n).Dur = 0
    Next n

    For n = 1 To MAX_BANK
        Player(Index).Bank(n).num = 0
        Player(Index).Bank(n).Value = 0
        Player(Index).Bank(n).Dur = 0
    Next n

    Player(Index).ArmorSlot = 0
    Player(Index).WeaponSlot = 0
    Player(Index).HelmetSlot = 0
    Player(Index).ShieldSlot = 0
    Player(Index).LegsSlot = 0
    Player(Index).RingSlot = 0
    Player(Index).ArmorSlot = 0

    Player(Index).Map = 0
    Player(Index).X = 0
    Player(Index).y = 0
    Player(Index).Dir = 0

    ' Client use only
    Player(Index).MaxHp = 0
    Player(Index).MaxMP = 0
    Player(Index).MaxSP = 0
    Player(Index).xOffset = 0
    Player(Index).yOffset = 0
    Player(Index).MovingH = 0
    Player(Index).MovingV = 0
    Player(Index).Moving = 0
    Player(Index).Attacking = 0
    Player(Index).AttackTimer = 0
    Player(Index).MapGetTimer = 0
    Player(Index).CastedSpell = NO
    Player(Index).EmoticonNum = -1
    Player(Index).EmoticonTime = 0
    Player(Index).EmoticonVar = 0

    For i = 1 To MAX_SPELL_ANIM
        Player(Index).SpellAnim(i).CastedSpell = NO
        Player(Index).SpellAnim(i).SpellTime = 0
        Player(Index).SpellAnim(i).SpellVar = 0
        Player(Index).SpellAnim(i).SpellDone = 0

        Player(Index).SpellAnim(i).Target = 0
        Player(Index).SpellAnim(i).TargetType = 0
    Next i

    Player(Index).SpellNum = 0

    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i

    Inventory = 1
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index).Name = vbNullString
    Item(Index).desc = vbNullString

    Item(Index).Type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
    Item(Index).StrReq = 0
    Item(Index).DefReq = 0
    Item(Index).SpeedReq = 0
    Item(Index).MagicReq = 0
    Item(Index).ClassReq = -1
    Item(Index).AccessReq = 0

    Item(Index).AddHP = 0
    Item(Index).AddMP = 0
    Item(Index).AddSP = 0
    Item(Index).AddSTR = 0
    Item(Index).AddDEF = 0
    Item(Index).AddMAGI = 0
    Item(Index).AddSpeed = 0
    Item(Index).AddEXP = 0
    Item(Index).AttackSpeed = 1000
    Item(Index).Stackable = 0
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearMapItem(ByVal Index As Long)
    MapItem(Index).num = 0
    MapItem(Index).Value = 0
    MapItem(Index).Dur = 0
    MapItem(Index).X = 0
    MapItem(Index).y = 0
End Sub

Sub ClearMap()
    Dim i As Long
    Dim X As Long
    Dim y As Long

    For i = 1 To MAX_MAPS
        Map(i).Name = vbNullString
        Map(i).Revision = 0
        Map(i).Moral = 0
        Map(i).Up = 0
        Map(i).Down = 0
        Map(i).Left = 0
        Map(i).Right = 0
        Map(i).Indoors = 0

        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(i).Tile(X, y).Ground = 0
                Map(i).Tile(X, y).Mask = 0
                Map(i).Tile(X, y).Anim = 0
                Map(i).Tile(X, y).Mask2 = 0
                Map(i).Tile(X, y).M2Anim = 0
                Map(i).Tile(X, y).Fringe = 0
                Map(i).Tile(X, y).FAnim = 0
                Map(i).Tile(X, y).Fringe2 = 0
                Map(i).Tile(X, y).F2Anim = 0
                Map(i).Tile(X, y).Type = 0
                Map(i).Tile(X, y).Data1 = 0
                Map(i).Tile(X, y).Data2 = 0
                Map(i).Tile(X, y).Data3 = 0
                Map(i).Tile(X, y).String1 = vbNullString
                Map(i).Tile(X, y).String2 = vbNullString
                Map(i).Tile(X, y).String3 = vbNullString
                Map(i).Tile(X, y).light = 0
                Map(i).Tile(X, y).GroundSet = 0
                Map(i).Tile(X, y).MaskSet = 0
                Map(i).Tile(X, y).AnimSet = 0
                Map(i).Tile(X, y).Mask2Set = 0
                Map(i).Tile(X, y).M2AnimSet = 0
                Map(i).Tile(X, y).FringeSet = 0
                Map(i).Tile(X, y).FAnimSet = 0
                Map(i).Tile(X, y).Fringe2Set = 0
                Map(i).Tile(X, y).F2AnimSet = 0
            Next X
        Next y
    Next i
End Sub

Sub ClearMapItems()
    Dim X As Long

    For X = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(X)
    Next X
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    MapNpc(Index).num = 0
    MapNpc(Index).Target = 0
    MapNpc(Index).HP = 0
    MapNpc(Index).MP = 0
    MapNpc(Index).SP = 0
    MapNpc(Index).Map = 0
    MapNpc(Index).X = 0
    MapNpc(Index).y = 0
    MapNpc(Index).Dir = 0

    ' Client use only
    MapNpc(Index).xOffset = 0
    MapNpc(Index).yOffset = 0
    MapNpc(Index).Moving = 0
    MapNpc(Index).Attacking = 0
    MapNpc(Index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next i
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    If Index < 1 Or Index > MAX_PLAYERS Then
        Exit Function
    End If
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Guildaccess As Long)
    Player(Index).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).HP = HP

    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).HP = GetPlayerMaxHP(Index)
    End If
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).MP = GetPlayerMaxMP(Index)
    End If
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).SP = GetPlayerMaxSP(Index)
    End If
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    GetPlayerMaxHP = Player(Index).MaxHp
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    GetPlayerMaxMP = Player(Index).MaxMP
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    GetPlayerMaxSP = Player(Index).MaxSP
End Function

Function GetPlayerSTR(ByVal Index As Long) As Long
    GetPlayerSTR = Player(Index).STR
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).DEF
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal speed As Long)
    Player(Index).speed = speed
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).MAGI
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal MAGI As Long)
    Player(Index).MAGI = MAGI
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    If Index <= 0 Then
        Exit Function
    End If
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    If X >= 0 And X <= MAX_MAPX Then
        Player(Index).X = X
    End If
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).y = y
End Sub
Sub SetPlayerLoc(ByVal Index As Long, ByVal X As Long, ByVal y As Long)
    Player(Index).X = X
    Player(Index).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).ShieldSlot = InvNum
End Sub
Function GetPlayerLegsSlot(ByVal Index As Long) As Long
    GetPlayerLegsSlot = Player(Index).LegsSlot
End Function

Sub SetPlayerLegsSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).LegsSlot = InvNum
End Sub
Function GetPlayerRingSlot(ByVal Index As Long) As Long
    GetPlayerRingSlot = Player(Index).RingSlot
End Function

Sub SetPlayerRingSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).RingSlot = InvNum
End Sub
Function GetPlayerNecklaceSlot(ByVal Index As Long) As Long
    GetPlayerNecklaceSlot = Player(Index).NecklaceSlot
End Function

Sub SetPlayerNecklaceSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).NecklaceSlot = InvNum
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    If BankSlot > MAX_BANK Then
        Exit Function
    End If
    GetPlayerBankItemNum = Player(Index).Bank(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Player(Index).Bank(BankSlot).num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Player(Index).Bank(BankSlot).Value = ItemValue
End Sub

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemDur = Player(Index).Bank(BankSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
    Player(Index).Bank(BankSlot).Dur = ItemDur
End Sub

Function GetPlayerHead(ByVal Index As Long) As Long
    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerHead = Player(Index).head
    End If
End Function

Sub SetPlayerHead(ByVal Index As Long, ByVal head As Long)
    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).head = head
    End If
End Sub

Function GetPlayerBody(ByVal Index As Long) As Long
    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerBody = Player(Index).body
    End If
End Function

Sub SetPlayerBody(ByVal Index As Long, ByVal body As Long)
    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).body = body
    End If
End Sub

Function GetPlayerLeg(ByVal Index As Long) As Long
    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerLeg = Player(Index).leg
    End If
End Function

Sub SetPlayerLeg(ByVal Index As Long, ByVal leg As Long)
    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).leg = leg
    End If
End Sub

Function GetPlayerSkillLvl(ByVal Index As Long, ByVal skill As Long) As Long
    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerSkillLvl = Player(Index).SkilLvl(skill)
    End If
End Function

Sub SetPlayerSkillLvl(ByVal Index As Long, ByVal skill As Long, ByVal lvl As Long)
    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).SkilLvl(skill) = lvl
    End If
End Sub

Function GetPlayerSkillExp(ByVal Index As Long, ByVal skill As Long) As Long
    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerSkillExp = Player(Index).SkilExp(skill)
    End If
End Function

Sub SetPlayerSkillExp(ByVal Index As Long, ByVal skill As Long, ByVal lvl As Long)
    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).SkilExp(skill) = lvl
    End If
End Sub

Function GetPlayerPaperdoll(ByVal Index As Long) As Byte
    If Index < MAX_PLAYERS And Index > 0 Then
        If IsPlaying(Index) Then
            GetPlayerPaperdoll = Player(Index).paperdoll
        End If
    End If
End Function

Sub SetPlayerPaperdoll(ByVal Index As Long, ByVal paperdoll As Byte)
    If Index < MAX_PLAYERS And Index > 0 Then
        If IsPlaying(Index) Then
            Player(Index).paperdoll = paperdoll
        End If
    End If
End Sub
