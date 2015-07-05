Attribute VB_Name = "modTypes"
Option Explicit
Global PlayerI As Byte

' Winsock globals
Public GAME_PORT As Long

' General constants
Public GAME_NAME As String
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_GUILDS As Long
Public MAX_GUILD_MEMBERS As Long
Public MAX_EMOTICONS As Long
Public MAX_LEVEL As Long
Public Scripting As Byte
Public RndWeather As Byte
Public RndWeatherTime As Long
Public MiniMap As Byte
Public MAX_PARTY_MEMBERS As Long

Public Const MAX_OBJECTS = 50
Public Const MAX_FORMS = 10
Public Const MAX_ARROWS = 100
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
Public Const SEC_CODE1 = "kivbzesiakcxnnptooiclwybaxztikpuwkskolujpapklwqslytsltjixanlqhfvzjpknoxpomhvigzqbaexlqezxhzbmkwllqzqrrafmcorkcsrqawmzkopanhuwrpss"
Public Const SEC_CODE2 = "ymivliuuvecemubfusqppqtunnrfvlwoznupllpxdjkzmxpipbpbxiqvdehuboezocksuhlncjzuarlxsrmnduzbsxtpqmviabhfazurdbgzoiivmalgxtgvmdxoezctw"
Public Const SEC_CODE3 = "kyihwzgibcurhvdlxaefscwljuuysbrocjjvhuqpvjpfhczbbudwqvpllhejmczsvfvkxeuterwemnqelhfemowkcdlznwwkglvkwjfzbpzpfxqhgczcohnzwzalguiug"
Public Const SEC_CODE4 = "668826538473833515372521847018081208676207866013682070764222154065643077153132787312722784005565144715748807558748435336860252508"

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

Public RsText As Byte

' Monster Constants
Public Const MON_X = 32
Public Const MON_Y = 32

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
Public Const TILE_TYPE_CLASS_CHANGE = 17
Public Const TILE_TYPE_SCRIPTED = 18
Public Const TILE_TYPE_MINICON = 19
Public Const TILE_TYPE_BLOCKICON = 20
Public Const TILE_TYPE_NPCSPAWN = 21

' Item constants
Public Const ITEM_TYPE_NONE = 0
Public Const ITEM_TYPE_WEAPON = 1
Public Const ITEM_TYPE_ARMOR = 2
Public Const ITEM_TYPE_HELMET = 3
Public Const ITEM_TYPE_SHIELD = 4
Public Const ITEM_TYPE_POTIONADDHP = 5
Public Const ITEM_TYPE_POTIONADDMP = 6
Public Const ITEM_TYPE_POTIONSUBHP = 7
Public Const ITEM_TYPE_POTIONSUBMP = 8
Public Const ITEM_TYPE_KEY = 9
Public Const ITEM_TYPE_CURRENCY = 10
Public Const ITEM_TYPE_SPELL = 11
Public Const ITEM_TYPE_LAMP = 12
Public Const ITEM_TYPE_POTIONADDSP = 13
Public Const ITEM_TYPE_POTIONSUBSP = 14

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

' Spell constants
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_SUBHP = 2
Public Const SPELL_TYPE_SUBMP = 3
Public Const SPELL_TYPE_SCRIPTED = 4


' Target type constants
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1

Type PlayerInvRec
    num As Long
    Value As Long
    Dur As Long
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Long
    Exp As Long
    Access As Byte
    PK As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Long
    DEF As Long
    Luck As Long
    Magi As Long
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
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
End Type

Type PartyRec
    Leader As Byte
    Member() As Byte
    ShareExp As Boolean
End Type
    
Type CustomFormRec
    Title As String
    Feild(1 To MAX_OBJECTS) As String
End Type
    
Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    Email As String
       
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
    
    SpellTime As Long
    SpellVar As Long
    SpellDone As Long
    SpellNum As Long
    
    PartyStarter As Byte
    GettingMap As Byte
    Party As PartyRec
    InvitedBy As Byte
    
    Emoticon As Long

    InTrade As Byte
    TradePlayer As Long
    TradeOk As Byte
    TradeItemMax As Byte
    TradeItemMax2 As Byte
    Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
    
    InChat As Byte
    ChatPlayer As Long
    
    Mute As Boolean
    
    Color As Byte
    CustomForm(1 To MAX_FORMS) As CustomFormRec
    Ghost As Boolean
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
    
    AdvanceFrom As Long
    LevelReq As Long
    Type As Long
    Locked As Long
    
    MaleSprite As Long
    FemaleSprite As Long
    
    STR As Long
    DEF As Long
    Luck As Long
    Magi As Long
    
    Map As Long
    X As Byte
    Y As Byte
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 150
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    StrReq As Long
    DefReq As Long
    LuckReq As Long
    ClassReq As Long
    AccessReq As Byte
    
    AddHP As Long
    AddMP As Long
    AddSP As Long
    AddStr As Long
    AddDef As Long
    AddMagi As Long
    AddLuck As Long
    AddEXP As Long
    AttackSpeed As Long
End Type

Type MapItemRec
    num As Long
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
    Luck As Long
    Magi As Long
    Big As Long
    MaxHp As Long
    Exp As Long
    SpawnTime As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
End Type

Type MapNpcRec
    num As Long
    
    Target As Long
    
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
    MPCost As Long
    Sound As Long
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
    DoorOpen()  As Byte
    DoorTimer As Long
End Type

Type GuildRec
    Name As String * NAME_LENGTH
    Founder As String * NAME_LENGTH
    Member() As String * NAME_LENGTH
End Type

Type EmoRec
    Pic As Long
    Command As String
End Type

Type CMRec
    Title As String
    Message As String
End Type

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public MAX_CLASSES As Byte

Public Map() As MapRec
Public TempTile() As TempTileRec
Public PlayersOnMap() As Long
Public Player() As AccountRec
Public Class() As ClassRec
Public Class2() As ClassRec
Public Class3() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc() As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Guild() As GuildRec
Public Emoticons() As EmoRec
Public Experience() As Long
Public CMessages(1 To 6) As CMRec

Type ArrowRec
    Name As String
    Pic As Long
    Range As Byte
    Amount As Integer
    Ammo As Long
End Type
Public Arrows(1 To MAX_ARROWS) As ArrowRec

Type StatRec
    Level As Long
    STR As Long
    DEF As Long
    Magi As Long
    Luck As Long
End Type
Public AddHP As StatRec
Public AddMP As StatRec
Public AddSP As StatRec

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

    For I = 0 To MAX_CLASSES
        Class(I).Name = ""
        Class(I).AdvanceFrom = 0
        Class(I).LevelReq = 0
        Class(I).Type = 1
        Class(I).STR = 0
        Class(I).DEF = 0
        Class(I).Luck = 0
        Class(I).Magi = 0
        Class(I).FemaleSprite = 0
        Class(I).MaleSprite = 0
        Class(I).Map = 0
        Class(I).X = 0
        Class(I).Y = 0
    Next I
End Sub

Sub ClearClasses2()
Dim I As Long

    For I = 0 To MAX_CLASSES
        Class2(I).Name = ""
        Class2(I).AdvanceFrom = 0
        Class2(I).LevelReq = 0
        Class2(I).Type = 2
        Class2(I).STR = 0
        Class2(I).DEF = 0
        Class2(I).Luck = 0
        Class2(I).Magi = 0
        Class2(I).FemaleSprite = 0
        Class2(I).MaleSprite = 0
        Class2(I).Map = 0
        Class2(I).X = 0
        Class2(I).Y = 0
    Next I
End Sub

Sub ClearClasses3()
Dim I As Long

    For I = 0 To MAX_CLASSES
        Class3(I).Name = ""
        Class3(I).AdvanceFrom = 0
        Class3(I).LevelReq = 0
        Class3(I).Type = 3
        Class3(I).STR = 0
        Class3(I).DEF = 0
        Class3(I).Luck = 0
        Class3(I).Magi = 0
        Class3(I).FemaleSprite = 0
        Class3(I).MaleSprite = 0
        Class3(I).Map = 0
        Class3(I).X = 0
        Class3(I).Y = 0
    Next I
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim I As Long
Dim n As Long

    Player(Index).Login = ""
    Player(Index).Password = ""
    
    For I = 1 To MAX_CHARS
        Player(Index).Char(I).Name = ""
        Player(Index).Char(I).Class = 0
        Player(Index).Char(I).Level = 0
        Player(Index).Char(I).Sprite = 0
        Player(Index).Char(I).Exp = 0
        Player(Index).Char(I).Access = 0
        Player(Index).Char(I).PK = NO
        Player(Index).Char(I).POINTS = 0
        Player(Index).Char(I).Guild = ""
        
        Player(Index).Char(I).HP = 0
        Player(Index).Char(I).MP = 0
        Player(Index).Char(I).SP = 0
        
        Player(Index).Char(I).STR = 0
        Player(Index).Char(I).DEF = 0
        Player(Index).Char(I).Luck = 0
        Player(Index).Char(I).Magi = 0
        
        For n = 1 To MAX_INV
            Player(Index).Char(I).Inv(n).num = 0
            Player(Index).Char(I).Inv(n).Value = 0
            Player(Index).Char(I).Inv(n).Dur = 0
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(I).Spell(n) = 0
        Next n
        
        Player(Index).Char(I).ArmorSlot = 0
        Player(Index).Char(I).WeaponSlot = 0
        Player(Index).Char(I).HelmetSlot = 0
        Player(Index).Char(I).ShieldSlot = 0
        
        Player(Index).Char(I).Map = 0
        Player(Index).Char(I).X = 0
        Player(Index).Char(I).Y = 0
        Player(Index).Char(I).Dir = 0
        
        ' Temporary vars
        Player(Index).Buffer = ""
        Player(Index).IncBuffer = ""
        Player(Index).CharNum = 0
        Player(Index).InGame = False
        Player(Index).AttackTimer = 0
        Player(Index).DataTimer = 0
        Player(Index).DataBytes = 0
        Player(Index).DataPackets = 0
        Player(Index).PartyPlayer = 0
        Player(Index).InParty = 0
        Player(Index).Target = 0
        Player(Index).TargetType = 0
        Player(Index).CastedSpell = NO
        Player(Index).PartyStarter = NO
        Player(Index).GettingMap = NO
        Player(Index).Emoticon = -1
        Player(Index).InTrade = 0
        Player(Index).TradePlayer = 0
        Player(Index).TradeOk = 0
        Player(Index).TradeItemMax = 0
        Player(Index).TradeItemMax2 = 0
        For n = 1 To MAX_PLAYER_TRADES
            Player(Index).Trading(n).InvName = ""
            Player(Index).Trading(n).InvNum = 0
        Next n
        Player(Index).ChatPlayer = 0
    Next I
End Sub

Sub ClearChar(ByVal Index As Long, ByVal CharNum As Long)
Dim n As Long
    
    Player(Index).Char(CharNum).Name = ""
    Player(Index).Char(CharNum).Class = 0
    Player(Index).Char(CharNum).Sprite = 0
    Player(Index).Char(CharNum).Level = 0
    Player(Index).Char(CharNum).Exp = 0
    Player(Index).Char(CharNum).Access = 0
    Player(Index).Char(CharNum).PK = NO
    Player(Index).Char(CharNum).POINTS = 0
    Player(Index).Char(CharNum).Guild = ""
    
    Player(Index).Char(CharNum).HP = 0
    Player(Index).Char(CharNum).MP = 0
    Player(Index).Char(CharNum).SP = 0
    
    Player(Index).Char(CharNum).STR = 0
    Player(Index).Char(CharNum).DEF = 0
    Player(Index).Char(CharNum).Luck = 0
    Player(Index).Char(CharNum).Magi = 0
    
    For n = 1 To MAX_INV
        Player(Index).Char(CharNum).Inv(n).num = 0
        Player(Index).Char(CharNum).Inv(n).Value = 0
        Player(Index).Char(CharNum).Inv(n).Dur = 0
    Next n
    
    For n = 1 To MAX_PLAYER_SPELLS
        Player(Index).Char(CharNum).Spell(n) = 0
    Next n
    
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
    Item(Index).Name = ""
    Item(Index).Desc = ""
    
    Item(Index).Type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
    Item(Index).StrReq = 0
    Item(Index).DefReq = 0
    Item(Index).LuckReq = 0
    Item(Index).ClassReq = -1
    Item(Index).AccessReq = 0
    
    Item(Index).AddHP = 0
    Item(Index).AddMP = 0
    Item(Index).AddSP = 0
    Item(Index).AddStr = 0
    Item(Index).AddDef = 0
    Item(Index).AddMagi = 0
    Item(Index).AddLuck = 0
    Item(Index).AddEXP = 0
    Item(Index).AttackSpeed = 1000
End Sub

Sub ClearItems()
Dim I As Long

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next I
End Sub

Sub ClearNpc(ByVal Index As Long)
Dim I As Long
    Npc(Index).Name = ""
    Npc(Index).AttackSay = ""
    Npc(Index).Sprite = 0
    Npc(Index).SpawnSecs = 0
    Npc(Index).Behavior = 0
    Npc(Index).Range = 0
    Npc(Index).STR = 0
    Npc(Index).DEF = 0
    Npc(Index).Luck = 0
    Npc(Index).Magi = 0
    Npc(Index).Big = 0
    Npc(Index).MaxHp = 0
    Npc(Index).Exp = 0
    Npc(Index).SpawnTime = 0
    For I = 1 To MAX_NPC_DROPS
        Npc(Index).ItemNPC(I).Chance = 0
        Npc(Index).ItemNPC(I).ItemNum = 0
        Npc(Index).ItemNPC(I).ItemValue = 0
    Next I
End Sub

Sub ClearNpcs()
Dim I As Long

    For I = 1 To MAX_NPCS
        Call ClearNpc(I)
    Next I
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    MapItem(MapNum, Index).num = 0
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
    MapNpc(MapNum, Index).num = 0
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

    Map(MapNum).Name = ""
    Map(MapNum).Revision = 0
    Map(MapNum).Moral = 0
    Map(MapNum).Up = 0
    Map(MapNum).Down = 0
    Map(MapNum).Left = 0
    Map(MapNum).Right = 0
    Map(MapNum).Indoors = 0
        
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map(MapNum).Tile(X, Y).Ground = 0
            Map(MapNum).Tile(X, Y).Mask = 0
            Map(MapNum).Tile(X, Y).Anim = 0
            Map(MapNum).Tile(X, Y).Mask2 = 0
            Map(MapNum).Tile(X, Y).M2Anim = 0
            Map(MapNum).Tile(X, Y).Fringe = 0
            Map(MapNum).Tile(X, Y).FAnim = 0
            Map(MapNum).Tile(X, Y).Fringe2 = 0
            Map(MapNum).Tile(X, Y).F2Anim = 0
            Map(MapNum).Tile(X, Y).Type = 0
            Map(MapNum).Tile(X, Y).Data1 = 0
            Map(MapNum).Tile(X, Y).Data2 = 0
            Map(MapNum).Tile(X, Y).Data3 = 0
            Map(MapNum).Tile(X, Y).String1 = ""
            Map(MapNum).Tile(X, Y).String2 = ""
            Map(MapNum).Tile(X, Y).String3 = ""
            Map(MapNum).Tile(X, Y).Light = 0
            Map(MapNum).Tile(X, Y).GroundSet = 0
            Map(MapNum).Tile(X, Y).MaskSet = 0
            Map(MapNum).Tile(X, Y).AnimSet = 0
            Map(MapNum).Tile(X, Y).Mask2Set = 0
            Map(MapNum).Tile(X, Y).M2AnimSet = 0
            Map(MapNum).Tile(X, Y).FringeSet = 0
            Map(MapNum).Tile(X, Y).FAnimSet = 0
            Map(MapNum).Tile(X, Y).Fringe2Set = 0
            Map(MapNum).Tile(X, Y).F2AnimSet = 0
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
Dim z As Long

    Shop(Index).Name = ""
    Shop(Index).JoinSay = ""
    Shop(Index).LeaveSay = ""
    
    For z = 1 To 6
        For I = 1 To MAX_TRADES
            Shop(Index).TradeItem(z).Value(I).GiveItem = 0
            Shop(Index).TradeItem(z).Value(I).GiveValue = 0
            Shop(Index).TradeItem(z).Value(I).GetItem = 0
            Shop(Index).TradeItem(z).Value(I).GetValue = 0
        Next I
    Next z
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
    Spell(Index).MPCost = 0
    Spell(Index).Sound = 0
    Spell(Index).Range = 0
    
    Spell(Index).SpellAnim = 0
    Spell(Index).SpellTime = 40
    Spell(Index).SpellDone = 1
    
    Spell(Index).AE = 0
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

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Char(Player(Index).CharNum).Guild = Guild
End Sub

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
Dim I As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddHP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddHP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddHP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddHP
    End If

    CharNum = Player(Index).CharNum
    'GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSTR(index) / 2) + Class(Player(index).Char(CharNum).Class).STR) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(Index) * AddHP.Level) + (GetPlayerSTR(Index) * AddHP.STR) + (GetPlayerDEF(Index) * AddHP.DEF) + (GetPlayerMAGI(Index) * AddHP.Magi) + (GetPlayerLUCK(Index) * AddHP.Luck) + add
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddMP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddMP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddMP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddMP
    End If

    CharNum = Player(Index).CharNum
    'GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(Index) * AddMP.Level) + (GetPlayerSTR(Index) * AddMP.STR) + (GetPlayerDEF(Index) * AddMP.DEF) + (GetPlayerMAGI(Index) * AddMP.Magi) + (GetPlayerLUCK(Index) * AddMP.Luck) + add
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddSP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddSP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddSP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddSP
    End If

    CharNum = Player(Index).CharNum
    'GetPlayerMaxSP = ((Player(index).Char(CharNum).Level + Int(GetPlayerLUCK(index) / 2) + Class(Player(index).Char(CharNum).Class).Luck) * 2) + add
    GetPlayerMaxSP = (GetPlayerLevel(Index) * AddSP.Level) + (GetPlayerSTR(Index) * AddSP.STR) + (GetPlayerDEF(Index) * AddSP.DEF) + (GetPlayerMAGI(Index) * AddSP.Magi) + (GetPlayerLUCK(Index) * AddSP.Luck) + add
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim(Class(ClassNum).Name)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(Class(ClassNum).STR / 2) + Class(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(Class(ClassNum).Magi / 2) + Class(ClassNum).Magi) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(Class(ClassNum).Luck / 2) + Class(ClassNum).Luck) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = Class(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = Class(ClassNum).DEF
End Function

Function GetClassLUCK(ByVal ClassNum As Long) As Long
    GetClassLUCK = Class(ClassNum).Luck
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = Class(ClassNum).Magi
End Function

Function GetPlayerSTR(ByVal Index As Long) As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddStr
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddStr
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddStr
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddStr
    End If
    GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).STR + add
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddDef
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddDef
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddDef
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddDef
    End If
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF + add
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub

Function GetPlayerLUCK(ByVal Index As Long) As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddLuck
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddLuck
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddLuck
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddLuck
    End If
    GetPlayerLUCK = Player(Index).Char(Player(Index).CharNum).Luck + add
End Function

Sub SetPlayerLUCK(ByVal Index As Long, ByVal Luck As Long)
    Player(Index).Char(Player(Index).CharNum).Luck = Luck
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddMagi
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddMagi
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddMagi
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddMagi
    End If
    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).Magi + add
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

Sub BattleMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Byte)
    Call SendDataTo(Index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR)
End Sub

Function Rand(ByVal High As Long, ByVal Low As Long)
Randomize
High = High + 1
Do Until Rand >= Low
    Rand = Int(Rnd * High)
Loop
End Function

Function GetPlayerColor(ByVal Index As Long)
    If Index < 0 Or Index > MAX_PLAYERS Then
        GetPlayerColor = 0
        Exit Function
    End If
    GetPlayerColor = Player(Index).Color
End Function

Sub SetPlayerColor(ByVal Index As Long, ByVal Color As String)
    If Color < 0 Or Color > 16 Or Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Color = Color
End Sub

Function GetPlayerGhost(ByVal Index As Long)
    GetPlayerGhost = Player(Index).Ghost
End Function

Sub SetPlayerGhost(ByVal Index As Long, ByVal Inv As Boolean)
    Player(Index).Ghost = Inv
End Sub
