Attribute VB_Name = "modTypes"

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit
Public Const MAX_PARTIES = 500
Public TRADESKILL_TIMER As Long
Public PLAYER_CORPSES As Byte
Public NPC_CORPSES As Byte
Public SIZE_X As Integer
Public SIZE_Y As Integer
Public HPRegen As Byte
Public MPRegen As Byte
Public SPRegen As Byte
Public MOVEMENT_TIREDNESS As Byte
Global PlayerI As Byte
Public GAME_PORT As Long
Public PAPERDOLL As Integer
Public SPRITESIZE As Integer
Public WordList As Double
Public Wordfilter() As String
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
Public SCRIPTING As Long
Public Const MAX_PARTY_MEMBERS = 4
Public Const MAX_PARTY_INV_SLOTS = 10
Public MAX_SPEECH As Long
Public MAX_ELEMENTS As Long
Public Const MAX_ARROWS = 100
Public Const MAX_SPEECH_OPTIONS = 20
Public Const MAX_INV = 24
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 66
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10
Public Const MAX_FRIENDS = 20
Public Const MAX_BANK = 50
Public Const MAX_QUESTS = 500
Public Const NO = 0
Public Const YES = 1
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1
Public MAX_MAPX As Variant
Public MAX_MAPY As Variant
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_NO_PENALTY = 2
Public Const MAP_MORAL_HOUSE = 3
Public Const PIC_X = 32
Public Const PIC_Y = 32
Public Const MON_X = 32
Public Const MON_Y = 32
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
Public Const TILE_TYPE_NONE = 20
Public Const TILE_TYPE_BANK = 23
Public Const TILE_TYPE_HOUSE_BUY = 24
Public Const TILE_TYPE_HOUSE = 25
Public Const TILE_TYPE_FURNITURE = 26
Public Const TILE_TYPE_ROOF = 27
Public Const TILE_TYPE_ROOFBLOCK = 28
Public Const TILE_TYPE_SPAWNGATE = 29
Public Const TILE_TYPE_FISH = 30
Public Const TILE_TYPE_MINE = 31
Public Const TILE_TYPE_LJACKING = 32
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
Public Const ITEM_TYPE_PET = 14
Public Const ITEM_TYPE_FURNITURE = 15
Public Const ITEM_TYPE_SCRIPTED = 16
Public Const ITEM_TYPE_LEGS = 17
Public Const ITEM_TYPE_BOOTS = 18
Public Const ITEM_TYPE_GLOVES = 19
Public Const ITEM_TYPE_RING1 = 20
Public Const ITEM_TYPE_RING2 = 21
Public Const ITEM_TYPE_AMULET = 22
Public Const ITEM_TYPE_GUILDDEED = 23
Public Const ITEM_TYPE_HOUSEKEY = 24
Public Const ITEM_TYPE_FOOD = 25
Public Const ITEM_TYPE_ARROWS = 26
Public Const DIR_UP = 0
Public Const DIR_DOWN = 1
Public Const DIR_LEFT = 2
Public Const DIR_RIGHT = 3
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2
Public Const WEATHER_NONE = 0
Public Const WEATHER_RAINING = 1
Public Const WEATHER_SNOWING = 2
Public Const WEATHER_THUNDER = 3
Public Const TIME_DAY = 0
Public Const TIME_NIGHT = 1
Public Const ADMIN_MONITER = 1
Public Const ADMIN_MAPPER = 2
Public Const ADMIN_DEVELOPER = 3
Public Const ADMIN_CREATOR = 4
Public Const NPC_BEHAVIOR_ATTACKONSIGHT = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED = 1
Public Const NPC_BEHAVIOR_FRIENDLY = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER = 3
Public Const NPC_BEHAVIOR_GUARD = 4
Public Const NPC_BEHAVIOR_SCRIPTED = 5
Public Const NPC_BEHAVIOR_QUEST = 6
Public Const NPC_BEHAVIOR_BANKER = 7
Public Const NPC_BEHAVIOR_SPELLCASTER = 8
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_PET = 6
Public Const SPELL_TYPE_SCRIPTED = 7
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1
Public Const TARGET_TYPE_LOCATION = 2
Public Const TARGET_TYPE_PET = 3
Public Const EMOTICON_TYPE_IMAGE = 0
Public Const EMOTICON_TYPE_SOUND = 1
Public Const EMOTICON_TYPE_BOTH = 2

Type QuestRec
Name As String '
LevelIsReq As Byte '
ClassIsReq As Byte '
StartOn As Byte '
LevelReq As Integer '
ClassReq As Integer '

StartItem As Long '
Startval As Long '
ItemReq As Long '
ItemVal As Long '
RewardNum As Long '
RewardVal As Long '
Start As String '
End As String '
During As String '
NotHasItem As String '
Before As String '
After As String '
QuestExpReward As Long
End Type

Type ElementRec
    Name As String * NAME_LENGTH
    Strong As Long
    Weak As Long
End Type

Type BankRec
    num As Long
    Value As Long
    Dur As Long
End Type

Type PlayerInvRec
    num As Long
    Value As Long
    Dur As Long
End Type

Type PetRec
    Sprite As Long
    Alive As Byte
    Map As Long
    x As Long
    y As Long
    Dir As Byte
    Level As Long
    HP As Long
    MapToGo As Long
    XToGo As Long
    YToGo As Long
    Target As Long
    TargetType As Byte
    AttackTimer As Long
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
    Fp As Long

    ' Stats
    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    LegsSlot As Long
    BootsSlot As Long
    GlovesSlot As Long
    Ring1Slot As Long
    Ring2Slot As Long
    AmuletSlot As Long

    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    Bank(1 To MAX_BANK) As BankRec

    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    Hands As Long
    Friends(1 To MAX_FRIENDS) As String
    Alignment As Long
    FishExp As Long
    MineExp As Long
    LJackingExp As Long
    LargeBladesExp As Long
    SmallBladesExp As Long
    BluntWeaponsExp As Long
    PolesExp As Long
    AxesExp As Long
    ThrownExp As Long
    XbowsExp As Long
    BowsExp As Long
    FishLevel As Long
    MineLevel As Long
    LJackingLevel As Long
    LargeBladesLevel As Long
    SmallBladesLevel As Long
    BluntWeaponsLevel As Long
    PolesLevel As Long
    AxesLevel As Long
    ThrownLevel As Long
    XbowsLevel As Long
    BowsLevel As Long
    Race As Long
    SpawnGateMap As Long
    SpawnGateX As Long
    SpawnGateY As Long
    ArrowsAmount As Long
    QuestFlags(1 To MAX_QUESTS) As Long
    Poisoned As Byte
    Diseased As Byte
    AilmentInterval As Long
    AilmentMS As Long
    TradeSkillMS As Long
    InParty As Byte
    LookingForParty As Byte
    PartyInvitedTo As Byte
    PartyInvitedToBy As String
    Party As Byte
    ShieldLogin As Long
    LegsLogin As Long
    ArmorLogin As Long
    WeaponLogin As Long
    HelmetLogin As Long
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
    InvVal As Long
End Type

Type PartyRec
Leader As Byte
Member(1 To MAX_PARTY_MEMBERS) As Byte
Created As Boolean
TimeCreated As Double
End Type

Type AccountRec

    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    Email As String * NAME_LENGTH
    Vault As String * NAME_LENGTH

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
    PartyID As Long
    InParty As Byte
    PartyPlayer As Long
    PartyStarter As Byte
    Invited As Long
    TargetType As Byte
    Target As Long
    CastedSpell As Byte
    SpellVar As Long
    SpellDone As Long
    spellnum As Long
    GettingMap As Byte
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
    Pet As PetRec
    CorpseMap As Integer
    CorpseX As Byte
    CorpseY As Byte
    CorpseLoot(1 To 4) As PlayerInvRec
    CorpseTimer As Long
    OnlineTime As Currency
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
    GroundSet As Long
    MaskSet As Long
    AnimSet As Long
    Mask2Set As Long
    M2AnimSet As Long
    FringeSet As Long
    FAnimSet As Long
    Fringe2Set As Long
    F2AnimSet As Long
End Type

Type LocRec
    Used As Byte
    x As Long
    y As Long
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
    NpcSpawn(1 To MAX_MAP_NPCS) As LocRec
    Owner As String
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
    Speed As Long
    Magi As Long
    Map As Long
    x As Byte
    y As Byte
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
    SpeedReq As Long
    MagicReq As Long
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
    Price As Long
    Stackable As Long
    Bound As Long
    LevelReq As Long
    Element As Long
    StamRemove As Long
    Rarity As String * 11
    BowsReq As Long
    LargeBladesReq As Long
    SmallBladesReq As Long
    BluntWeaponsReq As Long
    PoleArmsReq As Long
    AxesReq As Long
    ThrownReq As Long
    XbowsReq As Long
    LBA As Long
    SBA As Long
    BWA As Long
    PAA As Long
    AA As Long
    TWA As Long
    XBA As Long
    BA As Long
    Poison As Long
    Disease As Long
    AilmentDamage As Long
    AilmentInterval As Long
    AilmentMS As Long
End Type

Type MapGridRec
    Blocked As Boolean
End Type

Type GridRec
    Loc() As MapGridRec
End Type

Type MapItemRec
    num As Long
    Value As Long
    Dur As Long
    x As Byte
    y As Byte
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
    Magi As Long
    Big As Long
    MaxHp As Long
    Exp As Long
    SpawnTime As Long
    Speech As Long
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
    Element As Long
    Poison As Long
    AP As Long
    Disease As Long
    Quest As Long
    NpcDIR As Byte
    AilmentDamage As Long
    AilmentInterval As Long
    AilmentMS As Long
    Spell As Long
End Type

Type MapNpcRec
    num As Long
    TargetType As Byte
    Target As Long
    HP As Long
    MP As Long
    SP As Long
    x As Byte
    y As Byte
    Dir As Integer

    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    LastAttack As Long
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
    sound As Long
    Type As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Range As Byte
    SpellAnim As Long
    SpellTime As Long
    SpellDone As Long
    AE As Long
    Pic As Long
    Element As Long
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
    sound As String
    Command As String
    Type As Byte
End Type

Type CMRec
    Title As String
    Message As String
End Type

Type OptionRec
    text As String
    GoTo As Long
    Exit As Byte
End Type

Type InvSpeechRec
    Exit As Byte
    text As String
    SaidBy As Byte
    Respond As Byte
    Script As Long
    Responces(1 To 3) As OptionRec
End Type

Type SpeechRec
    Name As String
    num(0 To MAX_SPEECH_OPTIONS) As InvSpeechRec
End Type

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1
Public NEXT_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte
Public Map() As MapRec
Public TempTile() As TempTileRec
Public PlayersOnMap() As Long
Public Player() As AccountRec
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc() As MapNpcRec
Public Grid() As GridRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Guild() As GuildRec
Public Emoticons() As EmoRec
Public Element() As ElementRec
Public Experience() As Long
Public CMessages(1 To 6) As CMRec
Public Speech() As SpeechRec
Public Quest(1 To MAX_QUESTS) As QuestRec
Public Party(1 To MAX_PARTIES) As PartyRec

Type ArrowRec
    Name As String
    Pic As Long
    Range As Byte
End Type

Public Arrows(1 To MAX_ARROWS) As ArrowRec
Type StatRec
    Level As Long
    STR As Long
    DEF As Long
    Magi As Long
    Speed As Long
End Type

Public AddHP As StatRec
Public AddMP As StatRec
Public AddSP As StatRec

Sub BattleMsg(ByVal Index As Long, _
   ByVal Msg As String, _
   ByVal Color As Byte, _
   ByVal Side As Byte)
    Call SendDataTo(Index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR)
End Sub

Sub ClearChar(ByVal Index As Long, _
   ByVal CharNum As Long)
Dim N As Long

    Player(Index).Char(CharNum).Name = ""
    Player(Index).Char(CharNum).Class = 1
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
    Player(Index).Char(CharNum).Fp = 0
    Player(Index).Char(CharNum).STR = 0
    Player(Index).Char(CharNum).DEF = 0
    Player(Index).Char(CharNum).Speed = 0
    Player(Index).Char(CharNum).Magi = 0
    Player(Index).Char(CharNum).Alignment = 0
    Player(Index).Char(CharNum).LargeBladesExp = 0
    Player(Index).Char(CharNum).SmallBladesExp = 0
    Player(Index).Char(CharNum).BluntWeaponsExp = 0
    Player(Index).Char(CharNum).PolesExp = 0
    Player(Index).Char(CharNum).AxesExp = 0
    Player(Index).Char(CharNum).ThrownExp = 0
    Player(Index).Char(CharNum).XbowsExp = 0
    Player(Index).Char(CharNum).BowsExp = 0
    Player(Index).Char(CharNum).Race = 1
    Player(Index).Char(CharNum).SpawnGateMap = 1
    Player(Index).Char(CharNum).SpawnGateX = 16
    Player(Index).Char(CharNum).SpawnGateY = 18
    Player(Index).Char(CharNum).ArrowsAmount = 0
    Player(Index).Char(CharNum).FishExp = 0
    Player(Index).Char(CharNum).MineExp = 0
    Player(Index).Char(CharNum).LJackingExp = 0
    Player(Index).Char(CharNum).FishLevel = 0
    Player(Index).Char(CharNum).MineLevel = 0
    Player(Index).Char(CharNum).LJackingLevel = 0
    Player(Index).Char(CharNum).Poisoned = 0
    Player(Index).Char(CharNum).Diseased = 0
    Player(Index).Char(CharNum).PartyInvitedTo = 0
    Player(Index).Char(CharNum).PartyInvitedToBy = 0
    Player(Index).Char(CharNum).LookingForParty = 0
    Player(Index).Char(CharNum).InParty = 0
    Player(Index).Char(CharNum).Party = 0
    Player(Index).Char(CharNum).HelmetLogin = 0
    Player(Index).Char(CharNum).ShieldLogin = 0
    Player(Index).Char(CharNum).ArmorLogin = 0
    Player(Index).Char(CharNum).LegsLogin = 0
    Player(Index).Char(CharNum).WeaponLogin = 0
    For N = 1 To MAX_INV
        Player(Index).Char(CharNum).Inv(N).num = 0
        Player(Index).Char(CharNum).Inv(N).Value = 0
        Player(Index).Char(CharNum).Inv(N).Dur = 0
    Next
    For N = 1 To MAX_PLAYER_SPELLS
        Player(Index).Char(CharNum).Spell(N) = 0
    Next
    For N = 1 To MAX_QUESTS
        Player(Index).Char(CharNum).QuestFlags(N) = 0
    Next
    For N = 1 To MAX_BANK
        Player(Index).Char(CharNum).Bank(N).num = 0
        Player(Index).Char(CharNum).Bank(N).Value = 0
        Player(Index).Char(CharNum).Bank(N).Dur = 0
    Next N
    Player(Index).Char(CharNum).ArmorSlot = 0
    Player(Index).Char(CharNum).WeaponSlot = 0
    Player(Index).Char(CharNum).HelmetSlot = 0
    Player(Index).Char(CharNum).ShieldSlot = 0
    Player(Index).Char(CharNum).LegsSlot = 0
    Player(Index).Char(CharNum).BootsSlot = 0
    Player(Index).Char(CharNum).GlovesSlot = 0
    Player(Index).Char(CharNum).Ring1Slot = 0
    Player(Index).Char(CharNum).Ring2Slot = 0
    Player(Index).Char(CharNum).AmuletSlot = 0
    Player(Index).Char(CharNum).Map = 0
    Player(Index).Char(CharNum).x = 0
    Player(Index).Char(CharNum).y = 0
    Player(Index).Char(CharNum).Dir = 0
End Sub

Sub ClearClasses()
Dim I As Long

    For I = 1 To Max_Classes
        Class(I).Name = ""
        Class(I).AdvanceFrom = 0
        Class(I).LevelReq = 0
        Class(I).Type = 1
        Class(I).STR = 0
        Class(I).DEF = 0
        Class(I).Speed = 0
        Class(I).Magi = 0
        Class(I).FemaleSprite = 0
        Class(I).MaleSprite = 0
        Class(I).Map = 0
        Class(I).x = 0
        Class(I).y = 0
    Next
End Sub

Sub ClearGrid()
Dim I As Long, y As Long, x As Long

    For I = 1 To MAX_MAPS
        For x = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY
                Grid(I).Loc(x, y).Blocked = False
            Next
        Next
    Next
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
    Item(Index).SpeedReq = 0
    Item(Index).MagicReq = 0
    Item(Index).ClassReq = 0
    Item(Index).AccessReq = 0
    Item(Index).AddHP = 0
    Item(Index).AddMP = 0
    Item(Index).AddSP = 0
    Item(Index).AddStr = 0
    Item(Index).AddDef = 0
    Item(Index).AddMagi = 0
    Item(Index).AddSpeed = 0
    Item(Index).AddEXP = 0
    Item(Index).AttackSpeed = 0
    Item(Index).Price = 0
    Item(Index).Stackable = 0
    Item(Index).Bound = 0
    Item(Index).LevelReq = 0
    Item(Index).Element = 0
    Item(Index).StamRemove = 0
    Item(Index).Rarity = "&HFFFFFF"
    Item(Index).BowsReq = 0
    Item(Index).LargeBladesReq = 0
    Item(Index).SmallBladesReq = 0
    Item(Index).BluntWeaponsReq = 0
    Item(Index).PoleArmsReq = 0
    Item(Index).AxesReq = 0
    Item(Index).ThrownReq = 0
    Item(Index).XbowsReq = 0
    Item(Index).LBA = 0
    Item(Index).SBA = 0
    Item(Index).BWA = 0
    Item(Index).PAA = 0
    Item(Index).AA = 0
    Item(Index).TWA = 0
    Item(Index).XBA = 0
    Item(Index).BA = 0
    Item(Index).Poison = 0
    Item(Index).Disease = 0
End Sub

Sub ClearItems()
Dim I As Long

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next
End Sub

Sub ClearMap(ByVal MapNum As Long)
Dim x As Long
Dim y As Long

    Map(MapNum).Name = ""
    Map(MapNum).Revision = 0
    Map(MapNum).Moral = 0
    Map(MapNum).Up = 0
    Map(MapNum).Down = 0
    Map(MapNum).Left = 0
    Map(MapNum).Right = 0
    Map(MapNum).Indoors = 0
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
            Map(MapNum).Tile(x, y).String1 = ""
            Map(MapNum).Tile(x, y).String2 = ""
            Map(MapNum).Tile(x, y).String3 = ""
            Map(MapNum).Tile(x, y).Light = 0
            Map(MapNum).Tile(x, y).GroundSet = -1
            Map(MapNum).Tile(x, y).MaskSet = -1
            Map(MapNum).Tile(x, y).AnimSet = -1
            Map(MapNum).Tile(x, y).Mask2Set = -1
            Map(MapNum).Tile(x, y).M2AnimSet = -1
            Map(MapNum).Tile(x, y).FringeSet = -1
            Map(MapNum).Tile(x, y).FAnimSet = -1
            Map(MapNum).Tile(x, y).Fringe2Set = -1
            Map(MapNum).Tile(x, y).F2AnimSet = -1
        Next
    Next

    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
End Sub

Sub ClearMapItem(ByVal Index As Long, _
   ByVal MapNum As Long)
    MapItem(MapNum, Index).num = 0
    MapItem(MapNum, Index).Value = 0
    MapItem(MapNum, Index).Dur = 0
    MapItem(MapNum, Index).x = 0
    MapItem(MapNum, Index).y = 0
End Sub

Sub ClearMapItems()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next
End Sub

Sub ClearMapNpc(ByVal Index As Long, _
   ByVal MapNum As Long)
    MapNpc(MapNum, Index).num = 0
    MapNpc(MapNum, Index).TargetType = 0
    MapNpc(MapNum, Index).Target = 0
    MapNpc(MapNum, Index).HP = 0
    MapNpc(MapNum, Index).MP = 0
    MapNpc(MapNum, Index).SP = 0
    MapNpc(MapNum, Index).x = 0
    MapNpc(MapNum, Index).y = 0
    MapNpc(MapNum, Index).Dir = 0

    ' Server use only
    MapNpc(MapNum, Index).SpawnWait = 0
    MapNpc(MapNum, Index).AttackTimer = 0
    MapNpc(MapNum, Index).LastAttack = 0
End Sub

Sub ClearMapNpcs()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next
End Sub

Sub ClearMaps()
Dim I As Long

    For I = 1 To MAX_MAPS
        Call ClearMap(I)
    Next
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
    Npc(Index).Speed = 0
    Npc(Index).Magi = 0
    Npc(Index).Big = 0
    Npc(Index).MaxHp = 0
    Npc(Index).Exp = 0
    Npc(Index).SpawnTime = 0
    Npc(Index).Speech = 0
    Npc(Index).Element = 0
    Npc(Index).Poison = 0
    Npc(Index).AP = 0
    Npc(Index).Disease = 0
    Npc(Index).Quest = 1
    Npc(Index).NpcDIR = 0
    Npc(Index).AilmentDamage = 0
    Npc(Index).AilmentInterval = 0
    Npc(Index).AilmentMS = 0
    Npc(Index).Spell = 0
    For I = 1 To MAX_NPC_DROPS
        Npc(Index).ItemNPC(I).Chance = 0
        Npc(Index).ItemNPC(I).ItemNum = 0
        Npc(Index).ItemNPC(I).ItemValue = 0
    Next
End Sub

Sub ClearNpcs()
Dim I As Long

    For I = 1 To MAX_NPCS
        Call ClearNpc(I)
    Next
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim I As Long
Dim N As Long

    Player(Index).Login = ""
    Player(Index).Password = ""
    
    Player(Index).CorpseMap = 0
    Player(Index).CorpseX = 0
    Player(Index).CorpseY = 0
    For I = 1 To 4
    Player(Index).CorpseLoot(I).Dur = 0
    Player(Index).CorpseLoot(I).num = 0
    Player(Index).CorpseLoot(I).Value = 0
    Next I
    
    For I = 1 To MAX_CHARS
        Player(Index).Char(I).Name = ""
        Player(Index).Char(I).Class = 1
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
        Player(Index).Char(I).Fp = 0
        Player(Index).Char(I).STR = 0
        Player(Index).Char(I).DEF = 0
        Player(Index).Char(I).Speed = 0
        Player(Index).Char(I).Magi = 0
        Player(Index).Char(I).Alignment = 0
        Player(Index).Char(I).LargeBladesExp = 0
        Player(Index).Char(I).SmallBladesExp = 0
        Player(Index).Char(I).BluntWeaponsExp = 0
        Player(Index).Char(I).PolesExp = 0
        Player(Index).Char(I).AxesExp = 0
        Player(Index).Char(I).ThrownExp = 0
        Player(Index).Char(I).XbowsExp = 0
        Player(Index).Char(I).BowsExp = 0
        Player(Index).Char(I).Race = 1
        Player(Index).Char(I).SpawnGateMap = 1
        Player(Index).Char(I).SpawnGateY = 16
        Player(Index).Char(I).SpawnGateX = 18
        Player(Index).Char(I).ArrowsAmount = 0
        Player(Index).Char(I).FishExp = 0
        Player(Index).Char(I).MineExp = 0
        Player(Index).Char(I).LJackingExp = 0
        Player(Index).Char(I).MineLevel = 0
        Player(Index).Char(I).LJackingLevel = 0
        Player(Index).Char(I).Poisoned = 0
        Player(Index).Char(I).Diseased = 0
        Player(Index).Char(I).PartyInvitedTo = 0
        Player(Index).Char(I).PartyInvitedToBy = 0
        Player(Index).Char(I).LookingForParty = 0
        Player(Index).Char(I).InParty = 0
        Player(Index).Char(I).Party = 0
        Player(Index).Char(I).HelmetLogin = 0
        Player(Index).Char(I).LegsLogin = 0
        Player(Index).Char(I).ArmorLogin = 0
        Player(Index).Char(I).ShieldLogin = 0
        Player(Index).Char(I).WeaponLogin = 0
        For N = 1 To MAX_INV
            Player(Index).Char(I).Inv(N).num = 0
            Player(Index).Char(I).Inv(N).Value = 0
            Player(Index).Char(I).Inv(N).Dur = 0
        Next
        For N = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(I).Spell(N) = 0
        Next
        For N = 1 To MAX_QUESTS
            Player(Index).Char(I).QuestFlags(N) = 0
        Next
        For N = 1 To MAX_BANK
            Player(Index).Char(I).Bank(N).num = 0
            Player(Index).Char(I).Bank(N).Value = 0
            Player(Index).Char(I).Bank(N).Dur = 0
        Next N
        Player(Index).Char(I).ArmorSlot = 0
        Player(Index).Char(I).WeaponSlot = 0
        Player(Index).Char(I).HelmetSlot = 0
        Player(Index).Char(I).ShieldSlot = 0
        Player(Index).Char(I).LegsSlot = 0
        Player(Index).Char(I).BootsSlot = 0
        Player(Index).Char(I).GlovesSlot = 0
        Player(Index).Char(I).Ring1Slot = 0
        Player(Index).Char(I).Ring2Slot = 0
        Player(Index).Char(I).AmuletSlot = 0
        Player(Index).Char(I).Map = 0
        Player(Index).Char(I).x = 0
        Player(Index).Char(I).y = 0
        Player(Index).Char(I).Dir = 0
        For N = 1 To MAX_FRIENDS
            Player(Index).Char(I).Friends(N) = ""
        Next
    Next
    Player(Index).Pet.Alive = NO

    ' Temporary vars
    Player(Index).Buffer = ""
    Player(Index).IncBuffer = ""
    Player(Index).CharNum = 0
    Player(Index).InGame = False
    Player(Index).AttackTimer = 0
    Player(Index).DataTimer = 0
    Player(Index).DataBytes = 0
    Player(Index).DataPackets = 0
    Player(Index).PartyID = 0
    Player(Index).InParty = 0
    Player(Index).Invited = 0
    Player(Index).Target = 0
    Player(Index).TargetType = 0
    Player(Index).CastedSpell = NO
    Player(Index).GettingMap = NO
    Player(Index).Emoticon = -1
    Player(Index).InTrade = 0
    Player(Index).TradePlayer = 0
    Player(Index).TradeOk = 0
    Player(Index).TradeItemMax = 0
    Player(Index).TradeItemMax2 = 0
     Player(Index).PartyPlayer = 0
        Player(Index).InParty = 0
        Player(Index).PartyStarter = NO
        Player(Index).PartyPlayer = 0
    For N = 1 To MAX_PLAYER_TRADES
        Player(Index).Trading(N).InvName = ""
        Player(Index).Trading(N).InvNum = 0
    Next
    Player(Index).ChatPlayer = 0
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
        Next
    Next
End Sub

Sub ClearShops()
Dim I As Long

    For I = 1 To MAX_SHOPS
        Call ClearShop(I)
    Next
End Sub

Sub ClearSpeech(ByVal Index As Long)
Dim I As Long
Dim o As Long

    Speech(Index).Name = ""
    For o = 0 To MAX_SPEECH_OPTIONS
        Speech(Index).num(o).Exit = 0
        Speech(Index).num(o).Respond = 0
        Speech(Index).num(o).SaidBy = 0
        Speech(Index).num(o).text = "Write what you want to be said here."
        Speech(Index).num(o).Script = 0
        For I = 1 To 3
            Speech(Index).num(o).Responces(I).Exit = 0
            Speech(Index).num(o).Responces(I).GoTo = 0
            Speech(Index).num(o).Responces(I).text = "Write a responce here."
        Next
    Next
End Sub

Sub ClearSpeeches()
Dim I As Long

    For I = 1 To MAX_SPEECH
        Call ClearSpeech(I)
    Next
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
    Spell(Index).sound = 0
    Spell(Index).Range = 0
    Spell(Index).SpellAnim = 0
    Spell(Index).SpellTime = 40
    Spell(Index).SpellDone = 1
    Spell(Index).AE = 0
    Spell(Index).Pic = 0
    Spell(Index).Element = 0
End Sub

Sub ClearSpells()
Dim I As Long

    For I = 1 To MAX_SPELLS
        Call ClearSpell(I)
    Next
End Sub

Sub ClearTempTile()
Dim I As Long, y As Long, x As Long

    For I = 1 To MAX_MAPS
        TempTile(I).DoorTimer = 0
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                TempTile(I).DoorOpen(x, y) = NO
            Next
        Next
    Next
End Sub

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(Class(ClassNum).STR / 2) + Class(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(Class(ClassNum).Magi / 2) + Class(ClassNum).Magi) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(Class(ClassNum).Speed / 2) + Class(ClassNum).Speed) * 2
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
End Function

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).Char(Player(Index).CharNum).ArmorSlot
End Function

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Function GetPlayerDEF(ByVal Index As Long) As Long
Dim Add As Long

    Add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddDef
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddDef
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddDef
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddDef
    End If
    
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddDef
    End If
    
    If GetPlayerBootsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerBootsSlot(Index))).AddDef
    End If
    
    If GetPlayerGlovesSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerGlovesSlot(Index))).AddDef
    End If
    
    If GetPlayerRing1Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing1Slot(Index))).AddDef
    End If
    
    If GetPlayerRing2Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing2Slot(Index))).AddDef
    End If
    
    If GetPlayerAmuletSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerAmuletSlot(Index))).AddDef
    End If
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF + Add
End Function
Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Char(Player(Index).CharNum).Guildaccess
End Function

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).Char(Player(Index).CharNum).HelmetSlot
End Function

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
End Function

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num
End Function

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value
End Function

Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).Level
End Function

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Function GetPlayerMAGI(ByVal Index As Long) As Long
Dim Add As Long

    Add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddMagi
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddMagi
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddMagi
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddMagi
    End If
    
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddMagi
    End If
    
    If GetPlayerBootsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerBootsSlot(Index))).AddMagi
    End If
    
    If GetPlayerGlovesSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerGlovesSlot(Index))).AddMagi
    End If
    
    If GetPlayerRing1Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing1Slot(Index))).AddMagi
    End If
    
    If GetPlayerRing2Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing2Slot(Index))).AddMagi
    End If
    
    If GetPlayerAmuletSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerAmuletSlot(Index))).AddMagi
    End If
    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).Magi + Add
End Function

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
End Function

Function GetPlayerMaxHP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim Add As Long

    Add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddHP
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddHP
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddHP
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddHP
    End If
    
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddHP
    End If
    
    If GetPlayerBootsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerBootsSlot(Index))).AddHP
    End If
    
    If GetPlayerGlovesSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerGlovesSlot(Index))).AddHP
    End If
    
    If GetPlayerRing1Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing1Slot(Index))).AddHP
    End If
    
    If GetPlayerRing2Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing2Slot(Index))).AddHP
    End If
    
    If GetPlayerAmuletSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerAmuletSlot(Index))).AddHP
    End If
    CharNum = Player(Index).CharNum

    'GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerstr(index) / 2) + Class(Player(index).Char(CharNum).Class).str) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(Index) * AddHP.Level) + (GetPlayerstr(Index) * AddHP.STR) + (GetPlayerDEF(Index) * AddHP.DEF) + (GetPlayerMAGI(Index) * AddHP.Magi) + (GetPlayerSPEED(Index) * AddHP.Speed) + Add
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim Add As Long

    Add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddMP
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddMP
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddMP
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddMP
    End If
    
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddMP
    End If
    
    If GetPlayerBootsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerBootsSlot(Index))).AddMP
    End If
    
    If GetPlayerGlovesSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerGlovesSlot(Index))).AddMP
    End If
    
    If GetPlayerRing1Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing1Slot(Index))).AddMP
    End If
    
    If GetPlayerRing2Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing2Slot(Index))).AddMP
    End If
    
    If GetPlayerAmuletSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerAmuletSlot(Index))).AddMP
    End If
    CharNum = Player(Index).CharNum

    'GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(Index) * AddMP.Level) + (GetPlayerstr(Index) * AddMP.STR) + (GetPlayerDEF(Index) * AddMP.DEF) + (GetPlayerMAGI(Index) * AddMP.Magi) + (GetPlayerSPEED(Index) * AddMP.Speed) + Add
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim Add As Long

    Add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddSP
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddSP
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddSP
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddSP
    End If
    
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddSP
    End If
    
    If GetPlayerBootsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerBootsSlot(Index))).AddSP
    End If
    
    If GetPlayerGlovesSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerGlovesSlot(Index))).AddSP
    End If
    
    If GetPlayerRing1Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing1Slot(Index))).AddSP
    End If
    
    If GetPlayerRing2Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing2Slot(Index))).AddSP
    End If
    
    If GetPlayerAmuletSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerAmuletSlot(Index))).AddSP
    End If
    CharNum = Player(Index).CharNum

    'GetPlayerMaxSP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSPEED(index) / 2) + Class(Player(index).Char(CharNum).Class).SPEED) * 2) + add
    GetPlayerMaxSP = (GetPlayerLevel(Index) * AddSP.Level) + (GetPlayerstr(Index) * AddSP.STR) + (GetPlayerDEF(Index) * AddSP.DEF) + (GetPlayerMAGI(Index) * AddSP.Magi) + (GetPlayerSPEED(Index) * AddSP.Speed) + Add
End Function

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
End Function

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
End Function

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
End Function

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
End Function

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).Char(Player(Index).CharNum).ShieldSlot
End Function

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Function GetPlayerSPEED(ByVal Index As Long) As Long
Dim Add As Long

    Add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddSpeed
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddSpeed
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddSpeed
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddSpeed
    End If
    
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddSpeed
    End If
    
    If GetPlayerBootsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerBootsSlot(Index))).AddSpeed
    End If
    
    If GetPlayerGlovesSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerGlovesSlot(Index))).AddSpeed
    End If
    
    If GetPlayerRing1Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing1Slot(Index))).AddSpeed
    End If
    
    If GetPlayerRing2Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing2Slot(Index))).AddSpeed
    End If
    
    If GetPlayerAmuletSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerAmuletSlot(Index))).AddSpeed
    End If
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).Speed + Add
End Function

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot)
End Function

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).Sprite
End Function

Function GetPlayerstr(ByVal Index As Long) As Long
Dim Add As Long

    Add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddStr
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddStr
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddStr
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddStr
    End If
    
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddStr
    End If
    
    If GetPlayerBootsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerBootsSlot(Index))).AddStr
    End If
    
    If GetPlayerGlovesSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerGlovesSlot(Index))).AddStr
    End If
    
    If GetPlayerRing1Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing1Slot(Index))).AddStr
    End If
    
    If GetPlayerRing2Slot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing2Slot(Index))).AddStr
    End If
    
    If GetPlayerAmuletSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerAmuletSlot(Index))).AddStr
    End If
    GetPlayerstr = Player(Index).Char(Player(Index).CharNum).STR + Add
End Function

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).Char(Player(Index).CharNum).WeaponSlot
End Function

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).x
End Function

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).y
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(Player(Index).CharNum).Access = Access
End Sub

Sub SetPlayerArmorSlot(ByVal Index As Long, _
   InvNum As Long)
   
    If InvNum > 0 Then
    Player(Index).Char(Player(Index).CharNum).ArmorLogin = GetPlayerInvItemNum(Index, InvNum)
    Else
    Player(Index).Char(Player(Index).CharNum).ArmorLogin = 0
    End If
    'MsgBox "item #" & Player(Index).Char(Player(Index).CharNum).ArmorLogin
    Player(Index).Char(Player(Index).CharNum).ArmorSlot = InvNum
End Sub

Sub SetPlayerClass(ByVal Index As Long, _
   ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

Sub SetPlayerDEF(ByVal Index As Long, _
   ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub
Sub SetPlayerDir(ByVal Index As Long, _
   ByVal Dir As Long)
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
End Sub

Sub SetPlayerExp(ByVal Index As Long, _
   ByVal Exp As Long)
    Player(Index).Char(Player(Index).CharNum).Exp = Exp
End Sub

Sub SetPlayerGuild(ByVal Index As Long, _
   ByVal Guild As String)
    Player(Index).Char(Player(Index).CharNum).Guild = Guild
End Sub

Sub SetPlayerGuildAccess(ByVal Index As Long, _
   ByVal Guildaccess As Long)
    Player(Index).Char(Player(Index).CharNum).Guildaccess = Guildaccess
End Sub

Sub SetPlayerHelmetSlot(ByVal Index As Long, _
   InvNum As Long)
    If InvNum > 0 Then
    Player(Index).Char(Player(Index).CharNum).HelmetLogin = GetPlayerInvItemNum(Index, InvNum)
    'Call PlayerMsg(Index, "Helmet Entry for Login is " & Player(Index).Char(Player(Index).CharNum).HelmetLogin & " !", Green)
    Else
    Player(Index).Char(Player(Index).CharNum).HelmetLogin = 0
    End If
   
    Player(Index).Char(Player(Index).CharNum).HelmetSlot = InvNum
End Sub

Sub SetPlayerHP(ByVal Index As Long, _
   ByVal HP As Long)
    Player(Index).Char(Player(Index).CharNum).HP = HP

    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    End If

    If GetPlayerHP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).HP = 0
    End If
    Call SendStats(Index)
    
    If GetPlayerParty(Index) > 0 Then
    Dim I As Long, N As Long, StatPercent As Byte, Packet As String
    For I = 1 To MAX_PARTY_MEMBERS
    If Party(GetPlayerParty(Index)).Member(I) = Index Then
    StatPercent = Val(((GetPlayerHP(Index) / 100) / (GetPlayerMaxHP(Index) / 100)) * 42)
    Packet = "k" & SEP_CHAR & I & SEP_CHAR & StatPercent & SEP_CHAR & END_CHAR
    Call SendDataToParty(GetPlayerParty(Index), Packet)
    End If
    Next I
    End If
End Sub

Sub SetPlayerInvItemDur(ByVal Index As Long, _
   ByVal InvSlot As Long, _
   ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Sub SetPlayerInvItemNum(ByVal Index As Long, _
   ByVal InvSlot As Long, _
   ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num = ItemNum
End Sub

Sub SetPlayerInvItemValue(ByVal Index As Long, _
   ByVal InvSlot As Long, _
   ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Sub SetPlayerLevel(ByVal Index As Long, _
   ByVal Level As Long)
    Player(Index).Char(Player(Index).CharNum).Level = Level
   ' Call SendLevel(Index)
    
    If GetPlayerParty(Index) > 0 Then
    Dim I As Long, N As Long, StatPercent As Long, Packet As String
    For I = 1 To MAX_PARTY_MEMBERS
    If Party(GetPlayerParty(Index)).Member(I) = Index Then
    StatPercent = GetPlayerLevel(Index)
    Packet = "l" & SEP_CHAR & I & SEP_CHAR & StatPercent & SEP_CHAR & END_CHAR
    Call SendDataToParty(GetPlayerParty(Index), Packet)
    End If
    Next I
    End If
End Sub

Sub SetPlayerMAGI(ByVal Index As Long, _
   ByVal Magi As Long)
    Player(Index).Char(Player(Index).CharNum).Magi = Magi
End Sub

Sub SetPlayerMap(ByVal Index As Long, _
   ByVal MapNum As Long)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(Player(Index).CharNum).Map = MapNum
    End If
End Sub

Sub SetPlayerMP(ByVal Index As Long, _
   ByVal MP As Long)
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    End If

    If GetPlayerMP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).MP = 0
    End If
    
    If GetPlayerParty(Index) > 0 Then
    Dim I As Long, N As Long, StatPercent As Byte, Packet As String
    For I = 1 To MAX_PARTY_MEMBERS
    If Party(GetPlayerParty(Index)).Member(I) = Index Then
    StatPercent = Val(((GetPlayerMP(Index) / 100) / (GetPlayerMaxMP(Index) / 100)) * 42)
    Packet = "m" & SEP_CHAR & I & SEP_CHAR & StatPercent & SEP_CHAR & END_CHAR
    Call SendDataToParty(GetPlayerParty(Index), Packet)
    End If
    Next I
    End If
End Sub

Sub SetPlayerPK(ByVal Index As Long, _
   ByVal PK As Long)
    Player(Index).Char(Player(Index).CharNum).PK = PK
End Sub

Sub SetPlayerPOINTS(ByVal Index As Long, _
   ByVal POINTS As Long)
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
End Sub

Sub SetPlayerShieldSlot(ByVal Index As Long, _
   InvNum As Long)
   
   If InvNum > 0 Then
    Player(Index).Char(Player(Index).CharNum).ShieldLogin = GetPlayerInvItemNum(Index, InvNum)
    'Call PlayerMsg(Index, "Helmet Entry for Login is " & Player(Index).Char(Player(Index).CharNum).HelmetLogin & " !", Green)
    Else
    Player(Index).Char(Player(Index).CharNum).ShieldLogin = 0
    End If
   
    Player(Index).Char(Player(Index).CharNum).ShieldSlot = InvNum
End Sub

Sub SetPlayerSP(ByVal Index As Long, _
   ByVal SP As Long)
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    End If

    If GetPlayerSP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).SP = 0
    End If
End Sub

Sub SetPlayerSPEED(ByVal Index As Long, _
   ByVal Speed As Long)
    Player(Index).Char(Player(Index).CharNum).Speed = Speed
End Sub

Sub SetPlayerSpell(ByVal Index As Long, _
   ByVal SpellSlot As Long, _
   ByVal spellnum As Long)
    Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot) = spellnum
End Sub

Sub SetPlayerSprite(ByVal Index As Long, _
   ByVal Sprite As Long)
    Player(Index).Char(Player(Index).CharNum).Sprite = Sprite
End Sub

Sub SetPlayerstr(ByVal Index As Long, _
   ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Sub SetPlayerWeaponSlot(ByVal Index As Long, _
   InvNum As Long)
   
   If InvNum > 0 Then
    Player(Index).Char(Player(Index).CharNum).WeaponLogin = GetPlayerInvItemNum(Index, InvNum)
    'Call PlayerMsg(Index, "Helmet Entry for Login is " & Player(Index).Char(Player(Index).CharNum).HelmetLogin & " !", Green)
    Else
    Player(Index).Char(Player(Index).CharNum).WeaponLogin = 0
    End If
   
    Player(Index).Char(Player(Index).CharNum).WeaponSlot = InvNum
End Sub

Sub SetPlayerX(ByVal Index As Long, _
   ByVal x As Long)
    Player(Index).Char(Player(Index).CharNum).x = x
End Sub

Sub SetPlayerY(ByVal Index As Long, _
   ByVal y As Long)
    Player(Index).Char(Player(Index).CharNum).y = y
End Sub

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

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
End Sub

Sub SetPlayerAlignment(ByVal Index As Long, ByVal num As Long)
   Player(Index).Char(Player(Index).CharNum).Alignment = num
End Sub

Function GetPlayerAlignment(ByVal Index As Long) As Long
    GetPlayerAlignment = Player(Index).Char(Player(Index).CharNum).Alignment
End Function

Function Rand(ByVal High As Long, ByVal Low As Long)
Randomize
High = High + 1
Do Until Rand >= Low
    Rand = Int(Rnd * High)
Loop
End Function

Function GetPlayerLegsSlot(ByVal Index As Long) As Long
    GetPlayerLegsSlot = Player(Index).Char(Player(Index).CharNum).LegsSlot
End Function

Function GetPlayerBootsSlot(ByVal Index As Long) As Long
    GetPlayerBootsSlot = Player(Index).Char(Player(Index).CharNum).BootsSlot
End Function

Function GetPlayerGlovesSlot(ByVal Index As Long) As Long
    GetPlayerGlovesSlot = Player(Index).Char(Player(Index).CharNum).GlovesSlot
End Function

Function GetPlayerRing1Slot(ByVal Index As Long) As Long
    GetPlayerRing1Slot = Player(Index).Char(Player(Index).CharNum).Ring1Slot
End Function

Function GetPlayerRing2Slot(ByVal Index As Long) As Long
    GetPlayerRing2Slot = Player(Index).Char(Player(Index).CharNum).Ring2Slot
End Function

Function GetPlayerAmuletSlot(ByVal Index As Long) As Long
    GetPlayerAmuletSlot = Player(Index).Char(Player(Index).CharNum).AmuletSlot
End Function

Sub SetPlayerLegsSlot(ByVal Index As Long, _
   InvNum As Long)
   
    If InvNum > 0 Then
    Player(Index).Char(Player(Index).CharNum).LegsLogin = GetPlayerInvItemNum(Index, InvNum)
    Else
    Player(Index).Char(Player(Index).CharNum).LegsLogin = 0
    End If
   
    Player(Index).Char(Player(Index).CharNum).LegsSlot = InvNum
End Sub

Sub SetPlayerBootsSlot(ByVal Index As Long, _
   InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).BootsSlot = InvNum
End Sub

Sub SetPlayerGlovesSlot(ByVal Index As Long, _
   InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).GlovesSlot = InvNum
End Sub

Sub SetPlayerRing1Slot(ByVal Index As Long, _
   InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).Ring1Slot = InvNum
End Sub

Sub SetPlayerRing2Slot(ByVal Index As Long, _
   InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).Ring2Slot = InvNum
End Sub

Sub SetPlayerAmuletSlot(ByVal Index As Long, _
   InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).AmuletSlot = InvNum
End Sub

Function GetPlayerRace(ByVal Index As Long) As Long
    GetPlayerRace = Player(Index).Char(Player(Index).CharNum).Race
End Function

Sub SetPlayerRace(ByVal Index As Long, _
   ByVal RaceNum As Long)
    Player(Index).Char(Player(Index).CharNum).Race = RaceNum
End Sub

Function GetPlayerLargeBladesExp(ByVal Index As Long) As Long
    GetPlayerLargeBladesExp = Player(Index).Char(Player(Index).CharNum).LargeBladesExp
End Function

Function GetPlayerLargeBladesLevel(ByVal Index As Long) As Long
    GetPlayerLargeBladesLevel = Player(Index).Char(Player(Index).CharNum).LargeBladesLevel
End Function

Function GetPlayerNextLargeBladesLevel(ByVal Index As Long) As Long
    GetPlayerNextLargeBladesLevel = Experience(GetPlayerLargeBladesLevel(Index))
End Function

Sub SetPlayerLargeBladesLevel(ByVal Index As Long, _
   ByVal LargeBladesLevel As Long)
    Player(Index).Char(Player(Index).CharNum).LargeBladesLevel = LargeBladesLevel
End Sub

Sub SetPlayerLargeBladesExp(ByVal Index As Long, _
   ByVal LargeBladesExp As Long)
    Player(Index).Char(Player(Index).CharNum).LargeBladesExp = LargeBladesExp
End Sub

Function GetPlayerSmallBladesExp(ByVal Index As Long) As Long
    GetPlayerSmallBladesExp = Player(Index).Char(Player(Index).CharNum).SmallBladesExp
End Function

Function GetPlayerSmallBladesLevel(ByVal Index As Long) As Long
    GetPlayerSmallBladesLevel = Player(Index).Char(Player(Index).CharNum).SmallBladesLevel
End Function

Function GetPlayerNextSmallBladesLevel(ByVal Index As Long) As Long
    GetPlayerNextSmallBladesLevel = Experience(GetPlayerSmallBladesLevel(Index))
End Function

Sub SetPlayerSmallBladesLevel(ByVal Index As Long, _
   ByVal SmallBladesLevel As Long)
    Player(Index).Char(Player(Index).CharNum).SmallBladesLevel = SmallBladesLevel
End Sub

Sub SetPlayerSmallBladesExp(ByVal Index As Long, _
   ByVal SmallBladesExp As Long)
    Player(Index).Char(Player(Index).CharNum).SmallBladesExp = SmallBladesExp
End Sub

Function GetPlayerBluntWeaponsExp(ByVal Index As Long) As Long
    GetPlayerBluntWeaponsExp = Player(Index).Char(Player(Index).CharNum).BluntWeaponsExp
End Function

Function GetPlayerBluntWeaponsLevel(ByVal Index As Long) As Long
    GetPlayerBluntWeaponsLevel = Player(Index).Char(Player(Index).CharNum).BluntWeaponsLevel
End Function

Function GetPlayerNextBluntWeaponsLevel(ByVal Index As Long) As Long
    GetPlayerNextBluntWeaponsLevel = Experience(GetPlayerBluntWeaponsLevel(Index))
End Function

Sub SetPlayerBluntWeaponsLevel(ByVal Index As Long, _
   ByVal BluntWeaponsLevel As Long)
    Player(Index).Char(Player(Index).CharNum).BluntWeaponsLevel = BluntWeaponsLevel
End Sub

Sub SetPlayerBluntWeaponsExp(ByVal Index As Long, _
   ByVal BluntWeaponsExp As Long)
    Player(Index).Char(Player(Index).CharNum).BluntWeaponsExp = BluntWeaponsExp
End Sub

Function GetPlayerPolesExp(ByVal Index As Long) As Long
    GetPlayerPolesExp = Player(Index).Char(Player(Index).CharNum).PolesExp
End Function

Function GetPlayerPolesLevel(ByVal Index As Long) As Long
    GetPlayerPolesLevel = Player(Index).Char(Player(Index).CharNum).PolesLevel
End Function

Function GetPlayerNextPolesLevel(ByVal Index As Long) As Long
    GetPlayerNextPolesLevel = Experience(GetPlayerPolesLevel(Index))
End Function

Sub SetPlayerPolesLevel(ByVal Index As Long, _
   ByVal PolesLevel As Long)
    Player(Index).Char(Player(Index).CharNum).PolesLevel = PolesLevel
End Sub

Sub SetPlayerPolesExp(ByVal Index As Long, _
   ByVal PolesExp As Long)
    Player(Index).Char(Player(Index).CharNum).PolesExp = PolesExp
End Sub

Function GetPlayerAxesExp(ByVal Index As Long) As Long
    GetPlayerAxesExp = Player(Index).Char(Player(Index).CharNum).AxesExp
End Function

Function GetPlayerAxesLevel(ByVal Index As Long) As Long
    GetPlayerAxesLevel = Player(Index).Char(Player(Index).CharNum).AxesLevel
End Function

Function GetPlayerNextAxesLevel(ByVal Index As Long) As Long
    GetPlayerNextAxesLevel = Experience(GetPlayerAxesLevel(Index))
End Function

Sub SetPlayerAxesLevel(ByVal Index As Long, _
   ByVal AxesLevel As Long)
    Player(Index).Char(Player(Index).CharNum).AxesLevel = AxesLevel
End Sub

Sub SetPlayerAxesExp(ByVal Index As Long, _
   ByVal AxesExp As Long)
    Player(Index).Char(Player(Index).CharNum).AxesExp = AxesExp
End Sub

Function GetPlayerThrownExp(ByVal Index As Long) As Long
    GetPlayerThrownExp = Player(Index).Char(Player(Index).CharNum).ThrownExp
End Function

Function GetPlayerThrownLevel(ByVal Index As Long) As Long
    GetPlayerThrownLevel = Player(Index).Char(Player(Index).CharNum).ThrownLevel
End Function

Function GetPlayerNextThrownLevel(ByVal Index As Long) As Long
    GetPlayerNextThrownLevel = Experience(GetPlayerThrownLevel(Index))
End Function

Sub SetPlayerThrownLevel(ByVal Index As Long, _
   ByVal ThrownLevel As Long)
    Player(Index).Char(Player(Index).CharNum).ThrownLevel = ThrownLevel
End Sub

Sub SetPlayerThrownExp(ByVal Index As Long, _
   ByVal ThrownExp As Long)
    Player(Index).Char(Player(Index).CharNum).ThrownExp = ThrownExp
End Sub

Function GetPlayerXbowsExp(ByVal Index As Long) As Long
    GetPlayerXbowsExp = Player(Index).Char(Player(Index).CharNum).XbowsExp
End Function

Function GetPlayerXbowsLevel(ByVal Index As Long) As Long
    GetPlayerXbowsLevel = Player(Index).Char(Player(Index).CharNum).XbowsLevel
End Function

Function GetPlayerNextXbowsLevel(ByVal Index As Long) As Long
    GetPlayerNextXbowsLevel = Experience(GetPlayerXbowsLevel(Index))
End Function

Sub SetPlayerXbowsLevel(ByVal Index As Long, _
   ByVal XbowsLevel As Long)
    Player(Index).Char(Player(Index).CharNum).XbowsLevel = XbowsLevel
End Sub

Sub SetPlayerXbowsExp(ByVal Index As Long, _
   ByVal XbowsExp As Long)
    Player(Index).Char(Player(Index).CharNum).XbowsExp = XbowsExp
End Sub

Function GetPlayerBowsExp(ByVal Index As Long) As Long
    GetPlayerBowsExp = Player(Index).Char(Player(Index).CharNum).BowsExp
End Function

Function GetPlayerBowsLevel(ByVal Index As Long) As Long
    GetPlayerBowsLevel = Player(Index).Char(Player(Index).CharNum).BowsLevel
End Function

Function GetPlayerNextBowsLevel(ByVal Index As Long) As Long
    GetPlayerNextBowsLevel = Experience(GetPlayerBowsLevel(Index))
End Function

Sub SetPlayerBowsLevel(ByVal Index As Long, _
   ByVal BowsLevel As Long)
    Player(Index).Char(Player(Index).CharNum).BowsLevel = BowsLevel
End Sub

Sub SetPlayerBowsExp(ByVal Index As Long, _
   ByVal BowsExp As Long)
    Player(Index).Char(Player(Index).CharNum).BowsExp = BowsExp
End Sub

Function GetPlayerSpawnGateX(ByVal Index As Long) As Long
    GetPlayerSpawnGateX = Player(Index).Char(Player(Index).CharNum).SpawnGateX
End Function

Function GetPlayerSpawnGateY(ByVal Index As Long) As Long
    GetPlayerSpawnGateY = Player(Index).Char(Player(Index).CharNum).SpawnGateY
End Function

Function GetPlayerSpawnGateMap(ByVal Index As Long) As Long
    GetPlayerSpawnGateMap = Player(Index).Char(Player(Index).CharNum).SpawnGateMap
End Function

Sub SetPlayerSpawnGateX(ByVal Index As Long, _
   ByVal x As Long)
    Player(Index).Char(Player(Index).CharNum).SpawnGateX = x
End Sub

Sub SetPlayerSpawnGateY(ByVal Index As Long, _
   ByVal y As Long)
    Player(Index).Char(Player(Index).CharNum).SpawnGateY = y
End Sub

Sub SetPlayerSpawnGateMap(ByVal Index As Long, _
   ByVal MapNum As Long)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(Player(Index).CharNum).SpawnGateMap = MapNum
    End If
End Sub

Function GetPlayerFP(ByVal Index As Long) As Long
    GetPlayerFP = Player(Index).Char(Player(Index).CharNum).Fp
End Function

Sub SetPlayerFP(ByVal Index As Long, _
   ByVal Fp As Long)
    Player(Index).Char(Player(Index).CharNum).Fp = Fp

    If GetPlayerFP(Index) > GetPlayerMaxFP(Index) Then
        Player(Index).Char(Player(Index).CharNum).Fp = GetPlayerMaxFP(Index)
    End If

    If GetPlayerFP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).Fp = 0
    End If
    Call SendFP(Index)
End Sub

Function GetPlayerMaxFP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim Add As Long

    Add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        'Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddHP
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        'Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddHP
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        'Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddHP
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        'Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddHP
    End If
    CharNum = Player(Index).CharNum
    
    If GetPlayerLegsSlot(Index) > 0 Then
        'Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddHP
    End If
    
    If GetPlayerBootsSlot(Index) > 0 Then
        'Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerBootsSlot(Index))).AddHP
    End If
    
    If GetPlayerGlovesSlot(Index) > 0 Then
        'Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerGlovesSlot(Index))).AddHP
    End If
    
    If GetPlayerRing1Slot(Index) > 0 Then
        'Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing1Slot(Index))).AddHP
    End If
    
    If GetPlayerRing2Slot(Index) > 0 Then
       ' Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRing2Slot(Index))).AddHP
    End If
    
    If GetPlayerAmuletSlot(Index) > 0 Then
        'Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerAmuletSlot(Index))).AddHP
    End If
    

    'GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerstr(index) / 2) + Class(Player(index).Char(CharNum).Class).str) * 2) + add
    GetPlayerMaxFP = 100
    '(GetPlayerLevel(Index) * AddHP.Level) + (GetPlayerstr(Index) * AddHP.STR) + (GetPlayerDEF(Index) * AddHP.DEF) + (GetPlayerMAGI(Index) * AddHP.Magi) + (GetPlayerSPEED(Index) * AddHP.Speed) + Add
End Function

Sub SetPlayerArrowsAmount(ByVal Index As Long, ByVal ArrowsAmount As Long)
   Player(Index).Char(Player(Index).CharNum).ArrowsAmount = ArrowsAmount
End Sub

Function GetPlayerArrowsAmount(ByVal Index As Long) As Long
    GetPlayerArrowsAmount = Player(Index).Char(Player(Index).CharNum).ArrowsAmount
End Function

Function GetPlayerFishExp(ByVal Index As Long) As Long
    GetPlayerFishExp = Player(Index).Char(Player(Index).CharNum).FishExp
End Function

Function GetPlayerMineExp(ByVal Index As Long) As Long
    GetPlayerMineExp = Player(Index).Char(Player(Index).CharNum).MineExp
End Function

Function GetPlayerFishLevel(ByVal Index As Long) As Long
    GetPlayerFishLevel = Player(Index).Char(Player(Index).CharNum).FishLevel
End Function

Function GetPlayerMineLevel(ByVal Index As Long) As Long
    GetPlayerMineLevel = Player(Index).Char(Player(Index).CharNum).MineLevel
End Function

Function GetPlayerNextFishLevel(ByVal Index As Long) As Long
    GetPlayerNextFishLevel = Experience(GetPlayerFishLevel(Index))
End Function

Function GetPlayerNextMineLevel(ByVal Index As Long) As Long
    GetPlayerNextMineLevel = Experience(GetPlayerMineLevel(Index))
End Function

Sub SetPlayerFishLevel(ByVal Index As Long, _
   ByVal FishLevel As Long)
    Player(Index).Char(Player(Index).CharNum).FishLevel = FishLevel
End Sub

Sub SetPlayerMineLevel(ByVal Index As Long, _
   ByVal MineLevel As Long)
    Player(Index).Char(Player(Index).CharNum).MineLevel = MineLevel
End Sub

Sub SetPlayerFishExp(ByVal Index As Long, _
   ByVal FishExp As Long)
    Player(Index).Char(Player(Index).CharNum).FishExp = FishExp
End Sub

Sub SetPlayerMineExp(ByVal Index As Long, _
   ByVal MineExp As Long)
    Player(Index).Char(Player(Index).CharNum).MineExp = MineExp
End Sub

Function GetPlayerLJackingExp(ByVal Index As Long) As Long
    GetPlayerLJackingExp = Player(Index).Char(Player(Index).CharNum).LJackingExp
End Function

Function GetPlayerLJackingLevel(ByVal Index As Long) As Long
    GetPlayerLJackingLevel = Player(Index).Char(Player(Index).CharNum).LJackingLevel
End Function

Function GetPlayerNextLJackingLevel(ByVal Index As Long) As Long
    GetPlayerNextLJackingLevel = Experience(GetPlayerLJackingLevel(Index))
End Function

Sub SetPlayerLJackingExp(ByVal Index As Long, _
   ByVal LJackingExp As Long)
    Player(Index).Char(Player(Index).CharNum).LJackingExp = LJackingExp
End Sub

Sub SetPlayerLJackingLevel(ByVal Index As Long, _
   ByVal LJackingLevel As Long)
    Player(Index).Char(Player(Index).CharNum).LJackingLevel = LJackingLevel
End Sub

Function GetPlayerVaultCode(ByVal Index As Long) As String
    GetPlayerVaultCode = Trim$(Player(Index).Vault)
End Function

Function GetPlayerQuestFlag(ByVal Index As Long, ByVal QuestFlagSlot As Long) As Long
    GetPlayerQuestFlag = Player(Index).Char(Player(Index).CharNum).QuestFlags(QuestFlagSlot)
End Function

Sub SetPlayerQuestFlag(ByVal Index As Long, _
   ByVal QuestFlagSlot As Long, _
   ByVal QuestFlagnum As Long)
    Player(Index).Char(Player(Index).CharNum).QuestFlags(QuestFlagSlot) = QuestFlagnum
    Call SendPlayerQuestFlags(Index)
End Sub

Sub SendPlayerQuestFlags(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "QUESTFLAGS" & SEP_CHAR
    For I = 1 To MAX_QUESTS
        Packet = Packet & GetPlayerQuestFlag(Index, I) & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SetPlayerPoisoned(ByVal Index As Long, _
   ByVal PoisonedNum As Long)
        Player(Index).Char(Player(Index).CharNum).Poisoned = PoisonedNum
End Sub

Function GetPlayerPoisoned(ByVal Index As Long) As Long
    GetPlayerPoisoned = Player(Index).Char(Player(Index).CharNum).Poisoned
End Function

Sub SetPlayerDiseased(ByVal Index As Long, _
   ByVal DiseasedNum As Long)
        Player(Index).Char(Player(Index).CharNum).Diseased = DiseasedNum
End Sub

Function GetPlayerDiseased(ByVal Index As Long) As Long
    GetPlayerDiseased = Player(Index).Char(Player(Index).CharNum).Diseased
End Function

Sub SetPlayerAilmentInterval(ByVal Index As Long, _
   ByVal AilmentIntervalNum As Long)
        Player(Index).Char(Player(Index).CharNum).AilmentInterval = AilmentIntervalNum
End Sub

Function GetPlayerAilmentInterval(ByVal Index As Long) As Long
    GetPlayerAilmentInterval = Player(Index).Char(Player(Index).CharNum).AilmentInterval
End Function

Sub SetPlayerAilmentMS(ByVal Index As Long, _
   ByVal AilmentMSNum As Long)
        Player(Index).Char(Player(Index).CharNum).AilmentMS = AilmentMSNum
End Sub

Function GetPlayerAilmentMS(ByVal Index As Long) As Long
    GetPlayerAilmentMS = Player(Index).Char(Player(Index).CharNum).AilmentMS
End Function

Sub SetPlayerTradeskillMS(ByVal Index As Long, _
   ByVal TradeskillMSNum As Long)
        Player(Index).Char(Player(Index).CharNum).TradeSkillMS = TradeskillMSNum
End Sub

Function GetPlayerTradeskillMS(ByVal Index As Long) As Long
    GetPlayerTradeskillMS = Player(Index).Char(Player(Index).CharNum).TradeSkillMS
End Function
