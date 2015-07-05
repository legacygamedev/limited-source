Attribute VB_Name = "modTypes"
Option Explicit

' General constants
Public GAME_NAME As String
Public WEBSITE As String
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_EMOTICONS As Long
Public MAX_SPELL_ANIM As Long
Public MAX_BLT_LINE As Long

Public Const MAX_ARROWS = 100
Public Const MAX_PLAYER_ARROWS = 100

Public Const MAX_INV = 24
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 66
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 = "nubraeigjrzmcxvaqwojfdsnvafnpingusagaourghaonvifaubhvkjcvnbaiusdgpinguusgpingusadfgyefigoadsfhvbasdufgeriygfuagdfhbdvhasgcuygefyg"
Public Const SEC_CODE2 = "aetubhnaurghaowriqpingureoldfahvnaidfbvauansdmxjcnvajdfbgiudafhgvjncjcxnvbaidhfbgisahdoadsujfouijaosdijfaocnonbaojfdoirjfoijaefnh"
Public Const SEC_CODE3 = "ajdbioujhandfvjnjwdfnxmxzjisoiadjhaznoigjaskglnmiodgahjaudfngjacvnacxnvmbpingufdogajfdgiohaxcnjvnacibgadifbjadjbnxcnvboijdsfgsdfh"
Public Const SEC_CODE4 = "198723460982745601473827562347594234623467295768794356217645643927652936759672496247629058762398456673076258720937183945943615796"

' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map constants
'Public Const MAX_MAPX = 30
'Public Const MAX_MAPY = 30
Public MAX_MAPX As Variant
Public MAX_MAPY As Variant
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_NO_PENALTY = 2

' Image constants
Public Const PIC_X = 32
Public Const PIC_Y = 32

' Tile consants
Public Const TILE_TYPE_WALKABLE = 0
Public Const TILE_TYPE_BLOCKED = 1
Public Const TILE_TYPE_WARP = 2
Public Const TILE_TYPE_ITEM = 3
Public Const TILE_TYPE_NPCAVOID = 4
Public Const TILE_TYPE_KEY = 5
Public Const TILE_TYPE_KEYOPEN = 6
Public Const TILE_TYPE_HEAL = 7
Public Const TILE_TYPE_KILL = 8
Public Const TILE_TYPE_SHOP = 9
Public Const TILE_TYPE_CBLOCK = 10
Public Const TILE_TYPE_ARENA = 11
Public Const TILE_TYPE_SOUND = 12
Public Const TILE_TYPE_SPRITE_CHANGE = 13
Public Const TILE_TYPE_SIGN = 14
Public Const TILE_TYPE_DOOR = 15
Public Const TILE_TYPE_NOTICE = 16
Public Const TILE_TYPE_CHEST = 17
Public Const TILE_TYPE_CLASS_CHANGE = 18
Public Const TILE_TYPE_SCRIPTED = 19
Public Const TILE_TYPE_NPC_SPAWN = 20

' Item constants
Public Const ITEM_TYPE_NONE = 0
Public Const ITEM_TYPE_WEAPON = 1
Public Const ITEM_TYPE_ARMOR = 2
Public Const ITEM_TYPE_HELMET = 3
Public Const ITEM_TYPE_SHIELD = 4
Public Const ITEM_TYPE_POTIONADDHP = 5
Public Const ITEM_TYPE_POTIONADDMP = 6
Public Const ITEM_TYPE_POTIONADDSP = 7
Public Const ITEM_TYPE_POTIONSUBHP = 8
Public Const ITEM_TYPE_POTIONSUBMP = 9
Public Const ITEM_TYPE_POTIONSUBSP = 10
Public Const ITEM_TYPE_KEY = 11
Public Const ITEM_TYPE_CURRENCY = 12
Public Const ITEM_TYPE_SPELL = 13

' Direction constants
Public Const DIR_UP = 0
Public Const DIR_DOWN = 1
Public Const DIR_LEFT = 2
Public Const DIR_RIGHT = 3

' Constants for player movement
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2

' Weather constants
Public Const WEATHER_NONE = 0
Public Const WEATHER_RAINING = 1
Public Const WEATHER_SNOWING = 2
Public Const WEATHER_THUNDER = 3

' Time constants
Public Const TIME_DAY = 0
Public Const TIME_NIGHT = 1

' Admin constants
Public Const ADMIN_MONITER = 1
Public Const ADMIN_MAPPER = 2
Public Const ADMIN_DEVELOPER = 3
Public Const ADMIN_CREATOR = 4

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED = 1
Public Const NPC_BEHAVIOR_FRIENDLY = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER = 3
Public Const NPC_BEHAVIOR_GUARD = 4

' Speach bubble constants
Public Const DISPLAY_BUBBLE_TIME = 2000 ' In milliseconds.
Public DISPLAY_BUBBLE_WIDTH As Byte
Public Const MAX_BUBBLE_WIDTH = 6 ' In tiles. Includes corners.
Public Const MAX_LINE_LENGTH = 23 ' In characters.
Public Const MAX_LINES = 3

' Spell constants
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5

Type ChatBubble
    Text As String
    Created As Long
End Type

Type PlayerInvRec
    Num As Long
    Value As Long
    Dur As Long
End Type

Type SpellAnimRec
    CastedSpell As Byte
    
    SpellTime As Long
    SpellVar As Long
    SpellDone As Long
    
    Target As Long
    TargetType As Long
End Type

Type PlayerArrowRec
    Arrow As Byte
    ArrowNum As Long
    ArrowAnim As Long
    ArrowTime As Long
    ArrowVarX As Long
    ArrowVarY As Long
    ArrowX As Long
    ArrowY As Long
    ArrowPosition As Byte
    ArrowAmount As Long
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Class As Long
    Sprite As Long
    Level As Long
    EXP As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Long
    DEF As Long
    Speed As Long
    MAGI As Long
    POINTS As Long
    
    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
       
    ' Position
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    
    ' Client use only
    MaxHp As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
    
    SpellNum As Long
    SpellAnim() As SpellAnimRec

    EmoticonNum As Long
    EmoticonTime As Long
    EmoticonVar As Long
    
    LevelUp As Long
    LevelUpT As Long

    Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec
End Type
    
Type TileRec
    Ground As Long
    Mask As Long
    Anim As Long
    Mask2 As Long
    M2Anim As Long
    Fringe As Long
    FAnim As Long
    Fringe2 As Long
    F2Anim As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    String2 As String
    String3 As String
    Light As Long
    GroundSet As Byte
    MaskSet As Byte
    AnimSet As Byte
    Mask2Set As Byte
    M2AnimSet As Byte
    FringeSet As Byte
    FAnimSet As Byte
    Fringe2Set As Byte
    F2AnimSet As Byte
End Type

Type MapRec
    Name As String * 40
    Revision As Long
    Moral As Byte
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    Music As String
    BootMap As Long
    BootX As Byte
    BootY As Byte
    Shop As Long
    Indoors As Byte
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    MaleSprite As Long
    FemaleSprite As Long
    
    Locked As Long
    
    STR As Long
    DEF As Long
    Speed As Long
    MAGI As Long
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    desc As String * 150
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    StrReq As Long
    DefReq As Long
    SpeedReq As Long
    ClassReq As Long
    AccessReq As Byte
    
    AddHP As Long
    AddMP As Long
    AddSP As Long
    AddStr As Long
    AddDef As Long
    AddMagi As Long
    AddSpeed As Long
    AddEXP As Long
    AttackSpeed As Long
End Type

Type MapItemRec
    Num As Long
    Value As Long
    Dur As Long
    
    X As Byte
    Y As Byte
End Type

Type NPCEditorRec
    ItemNum As Long
    ItemValue As Long
    Chance As Long
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    
    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
        
    STR  As Long
    DEF As Long
    Speed As Long
    MAGI As Long
    Big As Long
    MaxHp As Long
    EXP As Long
    SpawnTime As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
End Type

Type MapNpcRec
    Num As Long
    
    Target As Long
    
    HP As Long
    MaxHp As Long
    MP As Long
    SP As Long
    
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    Big As Byte

    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type TradeItemsRec
    Value(1 To MAX_TRADES) As TradeItemRec
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
    FixesItems As Byte
    TradeItem(1 To 6) As TradeItemsRec
End Type

Type SpellRec
    Name As String * NAME_LENGTH
    ClassReq As Long
    LevelReq As Long
    Sound As Long
    MPCost As Long
    Type As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Range As Byte
    
    SpellAnim As Long
    SpellTime As Long
    SpellDone As Long
    
    AE As Long
End Type

Type TempTileRec
    DoorOpen As Byte
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
End Type
Public Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public Trading2(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

Type EmoRec
    Pic As Long
    Command As String
End Type

Type DropRainRec
    X As Long
    Y As Long
    Randomized As Boolean
    Speed As Byte
End Type

' Bubble thing
Public Bubble() As ChatBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

Public Map() As MapRec
Public TempTile() As TempTileRec
Public Player() As PlayerRec
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Emoticons() As EmoRec
Public MapReport() As MapRec

Public MAX_RAINDROPS As Long
Public BLT_RAIN_DROPS As Long
Public DropRain() As DropRainRec

Public BLT_SNOW_DROPS As Long
Public DropSnow() As DropRainRec

Type ItemTradeRec
    ItemGetNum As Long
    ItemGiveNum As Long
    ItemGetVal As Long
    ItemGiveVal As Long
End Type
Type TradeRec
    Items(1 To MAX_TRADES) As ItemTradeRec
    Selected As Long
    SelectedItem As Long
End Type
Public Trade(1 To 6) As TradeRec

Type ArrowRec
    Name As String
    Pic As Long
    Range As Byte
    Amount As Long
End Type
Public Arrows(1 To MAX_ARROWS) As ArrowRec

Type BattleMsgRec
    Msg As String
    Index As Byte
    Color As Byte
    Time As Long
    Done As Byte
    Y As Long
End Type
Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Type ItemDurRec
    Item As Long
    Dur As Long
    Done As Byte
End Type
Public ItemDur(1 To 4) As ItemDurRec

Public Inventory As Long


Public Const MAX_ATTRIBUTE_NPCS = 25
Public MapAttributeNpc() As MapNpcRec
Public SaveMapAttributeNpc() As MapNpcRec


Sub ClearTempTile()
Dim X As Long, Y As Long

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            TempTile(X, Y).DoorOpen = NO
        Next X
    Next Y
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim i As Long
Dim n As Long

    Player(Index).Name = ""
    Player(Index).Guild = ""
    Player(Index).Guildaccess = 0
    Player(Index).Class = 0
    Player(Index).Level = 0
    Player(Index).Sprite = 0
    Player(Index).EXP = 0
    Player(Index).Access = 0
    Player(Index).PK = NO
        
    Player(Index).HP = 0
    Player(Index).MP = 0
    Player(Index).SP = 0
        
    Player(Index).STR = 0
    Player(Index).DEF = 0
    Player(Index).Speed = 0
    Player(Index).MAGI = 0
        
    For n = 1 To MAX_INV
        Player(Index).Inv(n).Num = 0
        Player(Index).Inv(n).Value = 0
        Player(Index).Inv(n).Dur = 0
    Next n
        
    Player(Index).ArmorSlot = 0
    Player(Index).WeaponSlot = 0
    Player(Index).HelmetSlot = 0
    Player(Index).ShieldSlot = 0
        
    Player(Index).Map = 0
    Player(Index).X = 0
    Player(Index).Y = 0
    Player(Index).Dir = 0
    
    ' Client use only
    Player(Index).MaxHp = 0
    Player(Index).MaxMP = 0
    Player(Index).MaxSP = 0
    Player(Index).XOffset = 0
    Player(Index).YOffset = 0
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
    Item(Index).Name = ""
    Item(Index).desc = ""
    
    Item(Index).Type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
    Item(Index).StrReq = 0
    Item(Index).DefReq = 0
    Item(Index).SpeedReq = 0
    Item(Index).ClassReq = -1
    Item(Index).AccessReq = 0
    
    Item(Index).AddHP = 0
    Item(Index).AddMP = 0
    Item(Index).AddSP = 0
    Item(Index).AddStr = 0
    Item(Index).AddDef = 0
    Item(Index).AddMagi = 0
    Item(Index).AddSpeed = 0
    Item(Index).AddEXP = 0
    Item(Index).AttackSpeed = 1000
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearMapItem(ByVal Index As Long)
    MapItem(Index).Num = 0
    MapItem(Index).Value = 0
    MapItem(Index).Dur = 0
    MapItem(Index).X = 0
    MapItem(Index).Y = 0
End Sub

Sub ClearMap()
Dim i As Long
Dim X As Long
Dim Y As Long

For i = 1 To MAX_MAPS
    Map(i).Name = ""
    Map(i).Revision = 0
    Map(i).Moral = 0
    Map(i).Up = 0
    Map(i).Down = 0
    Map(i).Left = 0
    Map(i).Right = 0
    Map(i).Indoors = 0
        
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map(i).Tile(X, Y).Ground = 0
            Map(i).Tile(X, Y).Mask = 0
            Map(i).Tile(X, Y).Anim = 0
            Map(i).Tile(X, Y).Mask2 = 0
            Map(i).Tile(X, Y).M2Anim = 0
            Map(i).Tile(X, Y).Fringe = 0
            Map(i).Tile(X, Y).FAnim = 0
            Map(i).Tile(X, Y).Fringe2 = 0
            Map(i).Tile(X, Y).F2Anim = 0
            Map(i).Tile(X, Y).Type = 0
            Map(i).Tile(X, Y).Data1 = 0
            Map(i).Tile(X, Y).Data2 = 0
            Map(i).Tile(X, Y).Data3 = 0
            Map(i).Tile(X, Y).String1 = ""
            Map(i).Tile(X, Y).String2 = ""
            Map(i).Tile(X, Y).String3 = ""
            Map(i).Tile(X, Y).Light = 0
            Map(i).Tile(X, Y).GroundSet = 0
            Map(i).Tile(X, Y).MaskSet = 0
            Map(i).Tile(X, Y).AnimSet = 0
            Map(i).Tile(X, Y).Mask2Set = 0
            Map(i).Tile(X, Y).M2AnimSet = 0
            Map(i).Tile(X, Y).FringeSet = 0
            Map(i).Tile(X, Y).FAnimSet = 0
            Map(i).Tile(X, Y).Fringe2Set = 0
            Map(i).Tile(X, Y).F2AnimSet = 0
        Next X
    Next Y
Next i
End Sub

Sub ClearMapItems()
Dim X As Long

    For X = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(X)
    Next X
End Sub

Sub ClearMapAttributeNpc(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    MapAttributeNpc(Index, X, Y).Num = 0
    MapAttributeNpc(Index, X, Y).Target = 0
    MapAttributeNpc(Index, X, Y).HP = 0
    MapAttributeNpc(Index, X, Y).MP = 0
    MapAttributeNpc(Index, X, Y).SP = 0
    MapAttributeNpc(Index, X, Y).Map = 0
    MapAttributeNpc(Index, X, Y).X = 0
    MapAttributeNpc(Index, X, Y).Y = 0
    MapAttributeNpc(Index, X, Y).Dir = 0
    
    ' Client use only
    MapAttributeNpc(Index, X, Y).XOffset = 0
    MapAttributeNpc(Index, X, Y).YOffset = 0
    MapAttributeNpc(Index, X, Y).Moving = 0
    MapAttributeNpc(Index, X, Y).Attacking = 0
    MapAttributeNpc(Index, X, Y).AttackTimer = 0
End Sub

Sub ClearMapAttributeNpcs()
Dim i As Long, X As Long, Y As Long
    
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            For i = 1 To MAX_ATTRIBUTE_NPCS
                Call ClearMapAttributeNpc(i, X, Y)
            Next i
        Next X
    Next Y
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    MapNpc(Index).Num = 0
    MapNpc(Index).Target = 0
    MapNpc(Index).HP = 0
    MapNpc(Index).MP = 0
    MapNpc(Index).SP = 0
    MapNpc(Index).Map = 0
    MapNpc(Index).X = 0
    MapNpc(Index).Y = 0
    MapNpc(Index).Dir = 0
    
    ' Client use only
    MapNpc(Index).XOffset = 0
    MapNpc(Index).YOffset = 0
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
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim(Player(Index).Guild)
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
    GetPlayerExp = Player(Index).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    Player(Index).EXP = EXP
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
    GetPlayerSPEED = Player(Index).Speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Speed = Speed
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
If Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(InvSlot).Num = ItemNum
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


