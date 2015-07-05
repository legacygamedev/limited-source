Attribute VB_Name = "modTypes"
Option Explicit

Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Type PlayerRec
    ' General
    FKey As Long
    Account As Long
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
    str As Byte
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
    X As Byte
    Y As Byte
    Dir As Byte
End Type
    
Type AccountRec
    ' Account
    FKey As Long
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    EKey As String
    
    HDModel As String
    HDSerial As String
       
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
    Fringe As Integer
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
    FKey As Long
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
    FKey As Long
    Name As String * NAME_LENGTH
    Sprite As Integer
    str As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte
End Type

Type ItemRec
    FKey As Long
    Name As String * NAME_LENGTH
    Pic As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    Unbreakable As Long
    Locked As Long
    Disabled As Long
    Assigned As Long
End Type

Type MapItemRec
    Num As Byte
    Value As Long
    Dur As Integer
    
    X As Byte
    Y As Byte
End Type

Type NpcRec
    FKey As Long
    Name As String * NAME_LENGTH
    AttackSay As String * 255
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    
    str  As Byte
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
        
    X As Byte
    Y As Byte
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
    FKey As Long
    Name As String * NAME_LENGTH
    JoinSay As String * 255
    LeaveSay As String * 255
    FixesItems As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type
    
Type SpellRec
    FKey As Long
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

Type xMapRec
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As Boolean 'Check to see if an npc or player is on the map.
End Type

Type uCache
    nDate As String
    nTime As String
    nCache As String
End Type
Public Cache(1 To MAX_CACHE) As uCache

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

Public Map(1 To MAX_MAPS) As MapRec
Public xMap(1 To MAX_MAPS) As xMapRec
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public Player(1 To MAX_PLAYERS) As AccountRec
Public Class() As ClassRec
Public Item(0 To MAX_ITEMS) As ItemRec
Public Npc(0 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Guild(1 To MAX_GUILDS) As GuildRec

Public pCount As Long
Public MOTD As String

Sub ClearTempTile()
Dim I As Long, Y As Long, X As Long

    For I = 1 To MAX_MAPS
        TempTile(I).DoorTimer = 0
        
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                TempTile(I).DoorOpen(X, Y) = NO
            Next X
        Next Y
    Next I
End Sub

Sub ClearClasses()
Dim I As Long

    For I = 0 To Max_Classes
        With Class(I)
            .FKey = 0
            .Name = ""
            .str = 0
            .DEF = 0
            .SPEED = 0
            .MAGI = 0
        End With
    Next I
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim I As Long
Dim N As Long

With Player(Index)
    .Login = ""
    .Password = ""
    .FKey = 0
    .EKey = ""
    .HDModel = ""
    .HDSerial = ""
End With

For I = 1 To MAX_CHARS
    With Player(Index).Char(I)
        .FKey = 0
        .Name = ""
        .Class = 0
        .Level = 0
        .Sprite = 0
        .Exp = 0
        .Access = 0
        .PK = NO
        .POINTS = 0
        .Guild = 0
        
        .HP = 0
        .MP = 0
        .SP = 0
        
        .str = 0
        .DEF = 0
        .SPEED = 0
        .MAGI = 0
        
        For N = 1 To MAX_INV
            .Inv(N).Num = 0
            .Inv(N).Value = 0
            .Inv(N).Dur = 0
        Next N
        
        For N = 1 To MAX_PLAYER_SPELLS
            .Spell(N) = 0
        Next N
        
        .ArmorSlot = 0
        .WeaponSlot = 0
        .HelmetSlot = 0
        .ShieldSlot = 0
        
        .Map = 0
        .X = 0
        .Y = 0
        .Dir = 0
    End With
Next I

' Temporary vars
With Player(Index)
    .Buffer = ""
    .IncBuffer = ""
    .CharNum = 0
    .InGame = False
    .AttackTimer = 0
    .DataTimer = 0
    .DataBytes = 0
    .DataPackets = 0
    .PartyPlayer = 0
    .InParty = 0
    .Target = 0
    .TargetType = 0
    .CastedSpell = NO
    .PartyStarter = NO
    .GettingMap = NO
End With
End Sub

Sub ClearChar(ByVal Index As Long, ByVal CharNum As Long)
Dim N As Long
    
    Player(Index).Char(CharNum).Name = ""
    Player(Index).Char(CharNum).Class = 0
    Player(Index).Char(CharNum).Sprite = 0
    Player(Index).Char(CharNum).Level = 0
    Player(Index).Char(CharNum).Exp = 0
    Player(Index).Char(CharNum).Access = 0
    Player(Index).Char(CharNum).PK = NO
    Player(Index).Char(CharNum).POINTS = 0
    Player(Index).Char(CharNum).Guild = 0
    
    Player(Index).Char(CharNum).HP = 0
    Player(Index).Char(CharNum).MP = 0
    Player(Index).Char(CharNum).SP = 0
    
    Player(Index).Char(CharNum).str = 0
    Player(Index).Char(CharNum).DEF = 0
    Player(Index).Char(CharNum).SPEED = 0
    Player(Index).Char(CharNum).MAGI = 0
    
    For N = 1 To MAX_INV
        Player(Index).Char(CharNum).Inv(N).Num = 0
        Player(Index).Char(CharNum).Inv(N).Value = 0
        Player(Index).Char(CharNum).Inv(N).Dur = 0
    Next N
    
    For N = 1 To MAX_PLAYER_SPELLS
        Player(Index).Char(CharNum).Spell(N) = 0
    Next N
    
    Player(Index).Char(CharNum).ArmorSlot = 0
    Player(Index).Char(CharNum).WeaponSlot = 0
    Player(Index).Char(CharNum).HelmetSlot = 0
    Player(Index).Char(CharNum).ShieldSlot = 0
    
    Player(Index).Char(CharNum).Map = 0
    Player(Index).Char(CharNum).X = 0
    Player(Index).Char(CharNum).Y = 0
    Player(Index).Char(CharNum).Dir = 0
End Sub
    
Sub ClearItem(ByVal Index As Long)
With Item(Index)
    .FKey = 0
    .Name = ""
    .Type = 0
    .Data1 = 0
    .Data2 = 0
    .Data3 = 0
    .Unbreakable = 0
    .Locked = 0
    .Disabled = 0
    .Assigned = 0
End With
End Sub

Sub ClearItems()
Dim I As Long

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next I
End Sub

Sub ClearNpc(ByVal Index As Long)
With Npc(Index)
    .FKey = 0
    .Name = ""
    .AttackSay = ""
    .Sprite = 0
    .SpawnSecs = 0
    .Behavior = 0
    .Range = 0
    .DropChance = 0
    .DropItem = 0
    .DropItemValue = 0
    .str = 0
    .DEF = 0
    .SPEED = 0
    .MAGI = 0
End With
End Sub

Sub ClearNpcs()
Dim I As Long

    For I = 1 To MAX_NPCS
        Call ClearNpc(I)
    Next I
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    MapItem(MapNum, Index).Num = 0
    MapItem(MapNum, Index).Value = 0
    MapItem(MapNum, Index).Dur = 0
    MapItem(MapNum, Index).X = 0
    MapItem(MapNum, Index).Y = 0
End Sub

Sub ClearMapItems()
Dim X As Long
Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, Y)
        Next X
    Next Y
End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    MapNpc(MapNum, Index).Num = 0
    MapNpc(MapNum, Index).Target = 0
    MapNpc(MapNum, Index).HP = 0
    MapNpc(MapNum, Index).MP = 0
    MapNpc(MapNum, Index).SP = 0
    MapNpc(MapNum, Index).X = 0
    MapNpc(MapNum, Index).Y = 0
    MapNpc(MapNum, Index).Dir = 0
    
    ' Server use only
    MapNpc(MapNum, Index).SpawnWait = 0
    MapNpc(MapNum, Index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
Dim X As Long
Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, Y)
        Next X
    Next Y
End Sub

Sub ClearMap(ByVal MapNum As Long)
Dim I As Long
Dim X As Long
Dim Y As Long

    Map(MapNum).FKey = 0
    Map(MapNum).Name = ""
    Map(MapNum).Revision = 0
    Map(MapNum).Moral = 0
    Map(MapNum).Up = 0
    Map(MapNum).Down = 0
    Map(MapNum).Left = 0
    Map(MapNum).Right = 0
        
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map(MapNum).Tile(X, Y).Ground = 0
            Map(MapNum).Tile(X, Y).Mask = 0
            Map(MapNum).Tile(X, Y).Anim = 0
            Map(MapNum).Tile(X, Y).Fringe = 0
            Map(MapNum).Tile(X, Y).Type = 0
            Map(MapNum).Tile(X, Y).Data1 = 0
            Map(MapNum).Tile(X, Y).Data2 = 0
            Map(MapNum).Tile(X, Y).Data3 = 0
        
            'Do xMaps
            'xMap(MapNum).Tile(X, Y) = False
        Next X
    Next Y
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
End Sub

Sub ClearMaps()
Dim I As Long

    For I = 1 To MAX_MAPS
        Call ClearMap(I)
    Next I
End Sub

Sub ClearShop(ByVal Index As Long)
Dim I As Long

    Shop(Index).Name = ""
    Shop(Index).JoinSay = ""
    Shop(Index).LeaveSay = ""
    
    For I = 1 To MAX_TRADES
        Shop(Index).TradeItem(I).GiveItem = 0
        Shop(Index).TradeItem(I).GiveValue = 0
        Shop(Index).TradeItem(I).GetItem = 0
        Shop(Index).TradeItem(I).GetValue = 0
    Next I
End Sub

Sub ClearShops()
Dim I As Long

    For I = 1 To MAX_SHOPS
        Call ClearShop(I)
    Next I
End Sub

Sub ClearSpell(ByVal Index As Long)
    Spell(Index).Name = ""
    Spell(Index).ClassReq = 0
    Spell(Index).LevelReq = 0
    Spell(Index).Type = 0
    Spell(Index).Data1 = 0
    Spell(Index).Data2 = 0
    Spell(Index).Data3 = 0
End Sub

Sub ClearSpells()
Dim I As Long

    For I = 1 To MAX_SPELLS
        Call ClearSpell(I)
    Next I
End Sub




' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim(Player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

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
    GetPlayerNextLevel = (GetPlayerLevel(Index) + 1) * (GetPlayerSTR(Index) + GetPlayerDEF(Index) + GetPlayerMAGI(Index) + GetPlayerSPEED(Index) + GetPlayerPOINTS(Index)) * 25
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Char(Player(Index).CharNum).Exp = Exp
End Sub

Public Function GetPlayerFKey(ByVal Index As Long) As Long
GetPlayerFKey = Player(Index).Char(Player(Index).CharNum).FKey
End Function

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
Dim I As Long

    CharNum = Player(Index).CharNum
    GetPlayerMaxHP = (Player(Index).Char(CharNum).Level + Int(GetPlayerSTR(Index) / 2) + Class(Player(Index).Char(CharNum).Class).str) * 2
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
Dim CharNum As Long

    CharNum = Player(Index).CharNum
    GetPlayerMaxMP = (Player(Index).Char(CharNum).Level + Int(GetPlayerMAGI(Index) / 2) + Class(Player(Index).Char(CharNum).Class).MAGI) * 2
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
Dim CharNum As Long

    CharNum = Player(Index).CharNum
    GetPlayerMaxSP = (Player(Index).Char(CharNum).Level + Int(GetPlayerSPEED(Index) / 2) + Class(Player(Index).Char(CharNum).Class).SPEED) * 2
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim(Class(ClassNum).Name)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(Class(ClassNum).str / 2) + Class(ClassNum).str) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(Class(ClassNum).MAGI / 2) + Class(ClassNum).MAGI) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(Class(ClassNum).SPEED / 2) + Class(ClassNum).SPEED) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = Class(ClassNum).str
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

Function GetPlayerSTR(ByVal Index As Long) As Long
    GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).str
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal str As Long)
    Player(Index).Char(Player(Index).CharNum).str = str
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).SPEED
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal SPEED As Long)
    Player(Index).Char(Player(Index).CharNum).SPEED = SPEED
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).MAGI
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal MAGI As Long)
    Player(Index).Char(Player(Index).CharNum).MAGI = MAGI
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
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Num = ItemNum
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

