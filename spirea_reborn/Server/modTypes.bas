Attribute VB_Name = "modTypes"
Option Explicit


Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Sex As Byte
    Class As Byte
    Sprite As Integer
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    Guild As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte
    POINTS As Byte
    
    ' Worn equipment
    ArmorSlot As Byte
    WeaponSlot As Byte
    HelmetSlot As Byte
    ShieldSlot As Byte
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
    
    ' Position
    Map As Integer
    x As Byte
    y As Byte
    Dir As Byte
End Type
    
Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
       
    ' Characters (we use 0 to prevent a crash that still needs to be figured out)
    Char(0 To MAX_CHARS) As PlayerRec
    
    ' None saved local vars
    Buffer As String
    IncBuffer As String
    CharNum As Byte
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    PartyPlayer As Long
    InParty As Byte
    TargetType As Byte
    Target As Byte
    CastedSpell As Byte
    PartyStarter As Byte
    GettingMap As Byte
End Type

Type TileRec
Ground As Integer
Mask As Integer
Anim As Integer
Mask2 As Integer
M2Anim As Integer
Fringe As Integer
FAnim As Integer
Fringe2 As Integer
F2Anim As Integer
Type As Byte
Data1 As Integer
Data2 As Integer
Data3 As Integer
End Type

Type OldMapRec
    Name As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    Music As Byte
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Shop As Byte
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
End Type

Type MapRec
    Name As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    Music As Byte
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Shop As Byte
    Indoors As Byte
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    
    Sprite As Integer
    
    STR As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    
    Pic As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Type MapItemRec
    Num As Byte
    Value As Long
    Dur As Integer
    
    x As Byte
    y As Byte
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 255
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    
    STR  As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte

End Type

Type MapNpcRec
    Num As Integer
    
    Target As Integer
    
    HP As Long
    MP As Long
    SP As Long
        
    x As Byte
    y As Byte
    Dir As Integer
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    

End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 255
    LeaveSay As String * 255
    FixesItems As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type
    
Type SpellRec
    Name As String * NAME_LENGTH
    ClassReq As Byte
    LevelReq As Byte
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Type TempTileRec
    DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY)  As Byte
    DoorTimer As Long
End Type

Type GuildRec
    Name As String * NAME_LENGTH
    Founder As String * NAME_LENGTH
    Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
End Type



Sub ClearTempTile()
Dim i As Long, y As Long, x As Long

    For i = 1 To MAX_MAPS
        TempTile(i).DoorTimer = 0
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                TempTile(i).DoorOpen(x, y) = NO
            Next x
        Next y
    Next i
End Sub

Sub ClearClasses()
Dim i As Long

    For i = 0 To Max_Classes
        Class(i).Name = ""
        Class(i).STR = 0
        Class(i).DEF = 0
        Class(i).SPEED = 0
        Class(i).MAGI = 0
    Next i
End Sub

Sub ClearPlayer(ByVal index As Long)
Dim i As Long
Dim n As Long

    Player(index).Login = ""
    Player(index).Password = ""
    
    For i = 1 To MAX_CHARS
        Player(index).Char(i).Name = ""
        Player(index).Char(i).Class = 0
        Player(index).Char(i).Level = 0
        Player(index).Char(i).Sprite = 0
        Player(index).Char(i).Exp = 0
        Player(index).Char(i).Access = 0
        Player(index).Char(i).PK = NO
        Player(index).Char(i).POINTS = 0
        Player(index).Char(i).Guild = 0
        
        Player(index).Char(i).HP = 0
        Player(index).Char(i).MP = 0
        Player(index).Char(i).SP = 0
        
        Player(index).Char(i).STR = 0
        Player(index).Char(i).DEF = 0
        Player(index).Char(i).SPEED = 0
        Player(index).Char(i).MAGI = 0
        
        For n = 1 To MAX_INV
            Player(index).Char(i).Inv(n).Num = 0
            Player(index).Char(i).Inv(n).Value = 0
            Player(index).Char(i).Inv(n).Dur = 0
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            Player(index).Char(i).Spell(n) = 0
        Next n
        
        Player(index).Char(i).ArmorSlot = 0
        Player(index).Char(i).WeaponSlot = 0
        Player(index).Char(i).HelmetSlot = 0
        Player(index).Char(i).ShieldSlot = 0
        
        Player(index).Char(i).Map = 0
        Player(index).Char(i).x = 0
        Player(index).Char(i).y = 0
        Player(index).Char(i).Dir = 0
        
        ' Temporary vars
        Player(index).Buffer = ""
        Player(index).IncBuffer = ""
        Player(index).CharNum = 0
        Player(index).InGame = False
        Player(index).AttackTimer = 0
        Player(index).DataTimer = 0
        Player(index).DataBytes = 0
        Player(index).DataPackets = 0
        Player(index).PartyPlayer = 0
        Player(index).InParty = 0
        Player(index).Target = 0
        Player(index).TargetType = 0
        Player(index).CastedSpell = NO
        Player(index).PartyStarter = NO
        Player(index).GettingMap = NO
    Next i
End Sub

Sub ClearChar(ByVal index As Long, ByVal CharNum As Long)
Dim n As Long
    
    Player(index).Char(CharNum).Name = ""
    Player(index).Char(CharNum).Class = 0
    Player(index).Char(CharNum).Sprite = 0
    Player(index).Char(CharNum).Level = 0
    Player(index).Char(CharNum).Exp = 0
    Player(index).Char(CharNum).Access = 0
    Player(index).Char(CharNum).PK = NO
    Player(index).Char(CharNum).POINTS = 0
    Player(index).Char(CharNum).Guild = 0
    
    Player(index).Char(CharNum).HP = 0
    Player(index).Char(CharNum).MP = 0
    Player(index).Char(CharNum).SP = 0
    
    Player(index).Char(CharNum).STR = 0
    Player(index).Char(CharNum).DEF = 0
    Player(index).Char(CharNum).SPEED = 0
    Player(index).Char(CharNum).MAGI = 0
    
    For n = 1 To MAX_INV
        Player(index).Char(CharNum).Inv(n).Num = 0
        Player(index).Char(CharNum).Inv(n).Value = 0
        Player(index).Char(CharNum).Inv(n).Dur = 0
    Next n
    
    For n = 1 To MAX_PLAYER_SPELLS
        Player(index).Char(CharNum).Spell(n) = 0
    Next n
    
    Player(index).Char(CharNum).ArmorSlot = 0
    Player(index).Char(CharNum).WeaponSlot = 0
    Player(index).Char(CharNum).HelmetSlot = 0
    Player(index).Char(CharNum).ShieldSlot = 0
    
    Player(index).Char(CharNum).Map = 0
    Player(index).Char(CharNum).x = 0
    Player(index).Char(CharNum).y = 0
    Player(index).Char(CharNum).Dir = 0
End Sub
    
Sub ClearItem(ByVal index As Long)
    Item(index).Name = ""
    
    Item(index).Type = 0
    Item(index).Data1 = 0
    Item(index).Data2 = 0
    Item(index).Data3 = 0
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearNpc(ByVal index As Long)
    Npc(index).Name = ""
    Npc(index).AttackSay = ""
    Npc(index).Sprite = 0
    Npc(index).SpawnSecs = 0
    Npc(index).Behavior = 0
    Npc(index).Range = 0
    Npc(index).DropChance = 0
    Npc(index).DropItem = 0
    Npc(index).DropItemValue = 0
    Npc(index).STR = 0
    Npc(index).DEF = 0
    Npc(index).SPEED = 0
    Npc(index).MAGI = 0

End Sub

Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next i
End Sub

Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Long)
    MapItem(MapNum, index).Num = 0
    MapItem(MapNum, index).Value = 0
    MapItem(MapNum, index).Dur = 0
    MapItem(MapNum, index).x = 0
    MapItem(MapNum, index).y = 0
End Sub

Sub ClearMapItems()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next x
    Next y
End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal MapNum As Long)
    MapNpc(MapNum, index).Num = 0
    MapNpc(MapNum, index).Target = 0
    MapNpc(MapNum, index).HP = 0
    MapNpc(MapNum, index).MP = 0
    MapNpc(MapNum, index).SP = 0
    MapNpc(MapNum, index).x = 0
    MapNpc(MapNum, index).y = 0
    MapNpc(MapNum, index).Dir = 0
    
    ' Server use only
    MapNpc(MapNum, index).SpawnWait = 0
    MapNpc(MapNum, index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next x
    Next y
End Sub

Sub ClearMap(ByVal MapNum As Long)
Dim i As Long
Dim x As Long
Dim y As Long

    Map(MapNum).Name = ""
    Map(MapNum).Revision = 0
    Map(MapNum).Moral = 0
    Map(MapNum).Up = 0
    Map(MapNum).Down = 0
    Map(MapNum).Left = 0
    Map(MapNum).Right = 0
        
   For y = 0 To MAX_MAPY
For x = 0 To MAX_MAPX
Map(MapNum).Tile(x, y).Ground = 0
Map(MapNum).Tile(x, y).Mask = 0
Map(MapNum).Tile(x, y).Anim = 0
Map(MapNum).Tile(x, y).Mask2 = 0
Map(MapNum).Tile(x, y).M2Anim = 0
Map(MapNum).Tile(x, y).Fringe = 0
Map(MapNum).Tile(x, y).FAnim = 0
Map(MapNum).Tile(x, y).Fringe2 = 0
Map(MapNum).Tile(x, y).F2Anim = 0
Map(MapNum).Tile(x, y).Type = 0
Map(MapNum).Tile(x, y).Data1 = 0
Map(MapNum).Tile(x, y).Data2 = 0
Map(MapNum).Tile(x, y).Data3 = 0
Next x
Next y
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
End Sub

Sub ClearMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next i
End Sub

Sub ClearShop(ByVal index As Long)
Dim i As Long

    Shop(index).Name = ""
    Shop(index).JoinSay = ""
    Shop(index).LeaveSay = ""
    
    For i = 1 To MAX_TRADES
        Shop(index).TradeItem(i).GiveItem = 0
        Shop(index).TradeItem(i).GiveValue = 0
        Shop(index).TradeItem(i).GetItem = 0
        Shop(index).TradeItem(i).GetValue = 0
    Next i
End Sub

Sub ClearShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next i
End Sub

Sub ClearSpell(ByVal index As Long)
    Spell(index).Name = ""
    Spell(index).ClassReq = 0
    Spell(index).LevelReq = 0
    Spell(index).Type = 0
    Spell(index).Data1 = 0
    Spell(index).Data2 = 0
    Spell(index).Data3 = 0
End Sub

Sub ClearSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next i
End Sub




' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    GetPlayerPassword = Trim(Player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    Player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String
    GetPlayerName = Trim(Player(index).Char(Player(index).CharNum).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Char(Player(index).CharNum).Name = Name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Char(Player(index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Char(Player(index).CharNum).Class = ClassNum
End Sub

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
    GetPlayerNextLevel = (GetPlayerLevel(index) + 1) * (GetPlayerSTR(index) + GetPlayerDEF(index) + GetPlayerMAGI(index) + GetPlayerSPEED(index) + GetPlayerPOINTS(index)) * 25
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
Dim i As Long

    CharNum = Player(index).CharNum
    GetPlayerMaxHP = (Player(index).Char(CharNum).Level + Int(GetPlayerSTR(index) / 2) + Class(Player(index).Char(CharNum).Class).STR) * 2
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
Dim CharNum As Long

    CharNum = Player(index).CharNum
    GetPlayerMaxMP = (Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).MAGI) * 2
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long
Dim CharNum As Long

    CharNum = Player(index).CharNum
    GetPlayerMaxSP = (Player(index).Char(CharNum).Level + Int(GetPlayerSPEED(index) / 2) + Class(Player(index).Char(CharNum).Class).SPEED) * 2
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim(Class(ClassNum).Name)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(Class(ClassNum).STR / 2) + Class(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(Class(ClassNum).MAGI / 2) + Class(ClassNum).MAGI) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(Class(ClassNum).SPEED / 2) + Class(ClassNum).SPEED) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = Class(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = Class(ClassNum).DEF
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = Class(ClassNum).SPEED
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = Class(ClassNum).MAGI
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
    GetPlayerSPEED = Player(index).Char(Player(index).CharNum).SPEED
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal SPEED As Long)
    Player(index).Char(Player(index).CharNum).SPEED = SPEED
End Sub

Function GetPlayerMAGI(ByVal index As Long) As Long
    GetPlayerMAGI = Player(index).Char(Player(index).CharNum).MAGI
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal MAGI As Long)
    Player(index).Char(Player(index).CharNum).MAGI = MAGI
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
    GetPlayerInvItemNum = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal itemvalue As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value = itemvalue
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

