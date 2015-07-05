Attribute VB_Name = "modTypes"
Option Explicit
Public PlayerI As Byte

' Winsock globals
Public GAME_PORT As Long

' Map Control
Public IS_SCROLLING As Long

' General constants
Public GAME_NAME As String
Public MAX_PLAYERS As Integer
Public MAX_SPELLS As Integer
Public MAX_ELEMENTS As Integer
Public MAX_MAPS As Integer
Public MAX_SHOPS As Integer
Public MAX_SKILLS As Integer
Public MAX_QUESTS As Integer
Public MAX_ITEMS As Integer
Public MAX_NPCS As Integer
Public MAX_MAP_ITEMS As Integer
Public MAX_GUILDS As Integer
Public MAX_GUILD_MEMBERS As Integer
Public MAX_EMOTICONS As Integer
Public MAX_LEVEL As Integer
Public Scripting As Byte
Public MAX_PARTY_MEMBERS As Integer
Public Paperdoll As Byte
Public Spritesize As Byte
Public CUSTOM_SPRITE As Integer
Public MAX_SCRIPTSPELLS As Integer
Public ENCRYPT_PASS As String
Public ENCRYPT_TYPE As String
Public STAT1 As String
Public STAT2 As String
Public STAT3 As String
Public STAT4 As String

'ASGARD
Public WordList As Double
Public Wordfilter() As String

Public Const MAX_ARROWS = 100

Public Const MAX_INV = 24
Public Const MAX_BANK = 50
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 66
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10
Public Const MAX_SKILLS_SHEETS = 10
Public Const MAX_SKILL_LEVEL = 100
Public Const MAX_QUEST_LENGHT = 10
Public Const MAX_SHOP_ITEMS = 25

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 = "t4AZuz7e8daxs81BM3Gcy5WSKfu4528I3u9X58Ob1YjXvFxRa9Et44E1pQ2gzr34oh5Gg8sxpWV6UZr52H4qrcOV234JMUg2gn37b74Sw2g33iYtq49bqwl9"
Public Const SEC_CODE2 = "x5P3Nmfi76GYD8C9OHtbEntFbb9imD2xnE1v6zc63x713WZwjQ9w3Q3JRMt2wJI31YuziSRTKWbmui4UJj17fY5P14Wy5Kgu9q6L6DYpLVwj26c5BIuD9NqPx"
Public Const SEC_CODE3 = "XW8qUJ6J786I6p42MXXO98rMJKaMc5c3Q825yVkk4QP39H5lv1E19hi898fcIyY77Q1IQkJfaXJv5O93fX962WJD5uV6FQUWjLLz4rWAKJbkk6S2F74qO7csu"
Public Const SEC_CODE4 = "68164GVUt5P73KUD36c63D468kfT712415l7LDx3jvB17tPnN7USAgaCuzS7uVMk7cFg5qA6k8TvX2OmCgb6soZqCrw89je7nB2S52pgeR48IoluCGznv7bhf"

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
Public Const MAP_MORAL_HOUSE = 3

' Image constants
Public Const PIC_X = 32
Public Const PIC_Y = 32

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
Public Const TILE_TYPE_CHEST = 17
Public Const TILE_TYPE_CLASS_CHANGE = 18
Public Const TILE_TYPE_SCRIPTED = 19
Public Const TILE_TYPE_NPC_SPAWN = 20
Public Const TILE_TYPE_HOUSE = 21
Public Const TILE_TYPE_CANON = 22
Public Const TILE_TYPE_BANK = 23
Public Const TILE_TYPE_SKILL = 24
Public Const TILE_TYPE_GUILDBLOCK = 25
Public Const TILE_TYPE_HOOKSHOT = 26
Public Const TILE_TYPE_WALKTHRU = 27
Public Const TILE_TYPE_ROOF = 28
Public Const TILE_TYPE_ROOFBLOCK = 29
Public Const TILE_TYPE_ONCLICK = 30
Public Const TILE_TYPE_LOWER_STAT = 31

' Item constants
Public Const ITEM_TYPE_NONE = 0
Public Const ITEM_TYPE_WEAPON = 1
Public Const ITEM_TYPE_ARMOR = 2
Public Const ITEM_TYPE_HELMET = 3
Public Const ITEM_TYPE_SHIELD = 4
Public Const ITEM_TYPE_LEGS = 5
Public Const ITEM_TYPE_RING = 6
Public Const ITEM_TYPE_NECKLACE = 7
Public Const ITEM_TYPE_POTIONADDHP = 8
Public Const ITEM_TYPE_POTIONADDMP = 9
Public Const ITEM_TYPE_POTIONADDSP = 10
Public Const ITEM_TYPE_POTIONSUBHP = 11
Public Const ITEM_TYPE_POTIONSUBMP = 12
Public Const ITEM_TYPE_POTIONSUBSP = 13
Public Const ITEM_TYPE_KEY = 14
Public Const ITEM_TYPE_CURRENCY = 15
Public Const ITEM_TYPE_SPELL = 16
Public Const ITEM_TYPE_SCRIPTED = 17

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
Public Const NPC_BEHAVIOR_SCRIPTED = 5

' Spell constants
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_GIVEITEM = 6
Public Const SPELL_TYPE_SCRIPTED = 6

' Target type constants
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1
Public Const TARGET_TYPE_ATTRIBUTE_NPC = 2

'---NOTE to future developers!----------
' When loading, types ARE order-sensitive!
' This means do not change the order of variables in between
' versions, and add new variables to the end. This way, we can
' just load the old files! I learned that the hard way :D
'            -Pickle

Type PlayerInvRec
    num As Integer
    Value As Long
    Dur As Integer
End Type

Type BankRec
    num As Integer
    Value As Long
    Dur As Integer
End Type

Type ElementRec
    Name As String * NAME_LENGTH
    Strong As Integer
    Weak As Integer
End Type

Type QuestRec
    Name As String * NAME_LENGTH
    Pictop As Integer
    Picleft As Integer
    Map(0 To MAX_QUEST_LENGHT) As Integer
    X(0 To MAX_QUEST_LENGHT) As Integer
    Y(0 To MAX_QUEST_LENGHT) As Integer
    Npc(0 To MAX_QUEST_LENGHT) As Integer
    Script(0 To MAX_QUEST_LENGHT) As Integer
    ItemTake1num(0 To MAX_QUEST_LENGHT) As Integer
    ItemTake2num(0 To MAX_QUEST_LENGHT) As Integer
    ItemTake1val(0 To MAX_QUEST_LENGHT) As Integer
    ItemTake2val(0 To MAX_QUEST_LENGHT) As Integer
    ItemGive1num(0 To MAX_QUEST_LENGHT) As Integer
    ItemGive2num(0 To MAX_QUEST_LENGHT) As Integer
    ItemGive1val(0 To MAX_QUEST_LENGHT) As Integer
    ItemGive2val(0 To MAX_QUEST_LENGHT) As Integer
    ExpGiven(0 To MAX_QUEST_LENGHT) As Integer
End Type

Type SkillRec
    Name As String * NAME_LENGTH
    Action As String
    Fail As String
    Succes As String
    Pictop As Long
    Picleft As Long
    ItemTake1num(1 To MAX_SKILLS_SHEETS) As Integer
    ItemTake2num(1 To MAX_SKILLS_SHEETS) As Integer
    ItemGive1num(1 To MAX_SKILLS_SHEETS) As Integer
    ItemGive2num(1 To MAX_SKILLS_SHEETS) As Integer
    minlevel(1 To MAX_SKILLS_SHEETS) As Integer
    ExpGiven(1 To MAX_SKILLS_SHEETS) As Integer
    base_chance(1 To MAX_SKILLS_SHEETS) As Integer
    ItemTake1val(1 To MAX_SKILLS_SHEETS) As Integer
    ItemTake2val(1 To MAX_SKILLS_SHEETS) As Integer
    ItemGive1val(1 To MAX_SKILLS_SHEETS) As Integer
    ItemGive2val(1 To MAX_SKILLS_SHEETS) As Integer
    itemequiped(1 To MAX_SKILLS_SHEETS) As Integer
    'Added to bottom for compatiblity! What is displayed when the skill is attempted
    AttemptName As String
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Sex As Byte
    Class As Integer
    Sprite As Long
    Level As Integer
    Exp As Long
    access As Byte
    PK As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long

    ' Stats
    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Integer
    WeaponSlot As Integer
    HelmetSlot As Integer
    ShieldSlot As Integer
    LegsSlot As Integer
    RingSlot As Integer
    NecklaceSlot As Integer

    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Integer
    Bank(1 To MAX_BANK) As BankRec

    ' Position and movement
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte

    targetnpc As Integer

    head As Integer
    body As Integer
    leg As Integer

    SkillLvl() As Integer
    SkillExp() As Long

    Paperdoll As Byte
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
    locked As Boolean
    lockedspells As Boolean
    lockeditems As Boolean
    lockedattack As Boolean
    targetnpc As Long

    pet As Long
    HookShotX As Byte
    HookShotY As Byte

    'MENUS
    custom_msg As String
    custom_title As String

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
    light As Long
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

Type Temp_TileRec
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
    light As Long
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
    left As Long
    Right As Long
    music As String
    BootMap As Long
    BootX As Byte
    BootY As Byte
    Shop As Long
    Indoors As Byte
    tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Integer
    owner As String
    scrolling As Byte
    Weather As Integer
End Type

Type Temp_MapRec
    Name As String * 40
    Revision As Long
    Moral As Byte
    Up As Long
    Down As Long
    left As Long
    Right As Long
    music As String
    BootMap As Long
    BootX As Byte
    BootY As Byte
    Shop As Long
    Indoors As Byte
    tile() As Temp_TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    owner As String
    scrolling As Long
    Weather As Long
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    
    AdvanceFrom As Long
    LevelReq As Long
    
    Type As Long
    locked As Long
    
    MaleSprite As Long
    FemaleSprite As Long
    
    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long
    
    Map As Long
    X As Byte
    Y As Byte
    
    ' Description
    Desc As String
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
    Stackable As Byte
    Bound As Byte
    
    'Moved back to bottom... I suck :P -Pickle
    TwoHanded As Long
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
    Speed As Long
    Magi As Long
    Big As Long
    MaxHp As Long
    Exp As Long
    SpawnTime As Long

    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec

    Element As Long

    Spritesize As Byte
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
    owner As Long

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

Type ShopItemRec
    ItemNum As Integer
    Price As Integer
    Amount As Integer
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    FixesItems As Byte 'Does the shop fix items?
    BuysItems As Byte 'Does the shop buy items?
    ShowInfo As Byte   'Popup box with item info?
    ShopItem(1 To MAX_SHOP_ITEMS) As ShopItemRec  'The items
    currencyItem As Integer 'The item needed to buy the items
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
    Big As Long
    
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
    Command As String
End Type

Type CMRec
    Title As String
    Message As String
End Type

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Calender stuff
Public year As Long
Public month As Long
Public day As Long
Public weekday As Long

' Maximum classes
Public MAX_CLASSES As Byte

Public Map() As MapRec
Public MapPackets() As String
Public Temp_Map() As Temp_MapRec
Public TempTile() As TempTileRec
Public PlayersOnMap() As Long
Public Player() As AccountRec
Public Class() As ClassRec
Public Class2() As ClassRec
Public Class3() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public skill() As SkillRec
Public Quest() As QuestRec
Public MapItem() As MapItemRec
Public MapNpc() As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Guild() As GuildRec
Public Emoticons() As EmoRec
Public Element() As ElementRec
Public Experience() As Long
Public CMessages(1 To 6) As CMRec
Public CTimers As Collection

Type ArrowRec
    Name As String
    Pic As Long
    Range As Byte
    Amount As Integer
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

Public Const MAX_ATTRIBUTE_NPCS = 25

Type MapAttributeNpcRec
    num As Long

    Target As Long

    HP As Long
    MP As Long
    SP As Long

    X As Byte
    Y As Byte
    Dir As Integer
    owner As Long

    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
End Type

Public MapAttributeNpc() As MapAttributeNpcRec


Sub BattleMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Byte)

    Call SendDataTo(index, PacketID.DamageDisplay & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR)

End Sub

Sub ClearChar(ByVal index As Long, ByVal CharNum As Long)

  Dim n As Long

    Player(index).Char(CharNum).Name = vbNullString
    Player(index).Char(CharNum).Class = 0
    Player(index).Char(CharNum).Sprite = 0
    Player(index).Char(CharNum).Level = 0
    Player(index).Char(CharNum).Exp = 0
    Player(index).Char(CharNum).access = 0
    Player(index).Char(CharNum).PK = NO
    Player(index).Char(CharNum).POINTS = 0
    Player(index).Char(CharNum).Guild = vbNullString

    Player(index).Char(CharNum).HP = 0
    Player(index).Char(CharNum).MP = 0
    Player(index).Char(CharNum).SP = 0

    Player(index).Char(CharNum).STR = 0
    Player(index).Char(CharNum).DEF = 0
    Player(index).Char(CharNum).Speed = 0
    Player(index).Char(CharNum).Magi = 0

    For n = 1 To MAX_INV
        Player(index).Char(CharNum).Inv(n).num = 0
        Player(index).Char(CharNum).Inv(n).Value = 0
        Player(index).Char(CharNum).Inv(n).Dur = 0
    Next n

    For n = 1 To MAX_BANK
        Player(index).Char(CharNum).Bank(n).num = 0
        Player(index).Char(CharNum).Bank(n).Value = 0
        Player(index).Char(CharNum).Bank(n).Dur = 0
    Next n

    For n = 1 To MAX_PLAYER_SPELLS
        Player(index).Char(CharNum).Spell(n) = 0
    Next n

    Player(index).Char(CharNum).ArmorSlot = 0
    Player(index).Char(CharNum).WeaponSlot = 0
    Player(index).Char(CharNum).HelmetSlot = 0
    Player(index).Char(CharNum).ShieldSlot = 0
    Player(index).Char(CharNum).LegsSlot = 0
    Player(index).Char(CharNum).RingSlot = 0
    Player(index).Char(CharNum).NecklaceSlot = 0

    Player(index).Char(CharNum).Map = 0
    Player(index).Char(CharNum).X = 0
    Player(index).Char(CharNum).Y = 0
    Player(index).Char(CharNum).Dir = 0

End Sub

Sub ClearClasses()

  Dim i As Long

    For i = 0 To MAX_CLASSES
        Class(i).Name = ""
        Class(i).AdvanceFrom = 0
        Class(i).LevelReq = 0
        Class(i).Type = 1
        Class(i).STR = 0
        Class(i).DEF = 0
        Class(i).Speed = 0
        Class(i).Magi = 0
        Class(i).FemaleSprite = 0
        Class(i).MaleSprite = 0
        Class(i).Map = 0
        Class(i).X = 0
        Class(i).Y = 0
    Next i

End Sub

Sub ClearClasses2()

  Dim i As Long

    For i = 0 To MAX_CLASSES
        Class2(i).Name = ""
        Class2(i).AdvanceFrom = 0
        Class2(i).LevelReq = 0
        Class2(i).Type = 2
        Class2(i).STR = 0
        Class2(i).DEF = 0
        Class2(i).Speed = 0
        Class2(i).Magi = 0
        Class2(i).FemaleSprite = 0
        Class2(i).MaleSprite = 0
        Class2(i).Map = 0
        Class2(i).X = 0
        Class2(i).Y = 0
    Next i

End Sub

Sub ClearClasses3()

  Dim i As Long

    For i = 0 To MAX_CLASSES
        Class3(i).Name = ""
        Class3(i).AdvanceFrom = 0
        Class3(i).LevelReq = 0
        Class3(i).Type = 3
        Class3(i).STR = 0
        Class3(i).DEF = 0
        Class3(i).Speed = 0
        Class3(i).Magi = 0
        Class3(i).FemaleSprite = 0
        Class3(i).MaleSprite = 0
        Class3(i).Map = 0
        Class3(i).X = 0
        Class3(i).Y = 0
    Next i

End Sub

Sub ClearItem(ByVal index As Long)

    Item(index).Name = vbNullString
    Item(index).Desc = vbNullString

    Item(index).Type = 0
    Item(index).Data1 = 0
    Item(index).Data2 = 0
    Item(index).Data3 = 0
    Item(index).StrReq = 0
    Item(index).DefReq = 0
    Item(index).SpeedReq = 0
    Item(index).ClassReq = -1
    Item(index).AccessReq = 0

    Item(index).AddHP = 0
    Item(index).AddMP = 0
    Item(index).AddSP = 0
    Item(index).AddStr = 0
    Item(index).AddDef = 0
    Item(index).AddMagi = 0
    Item(index).AddSpeed = 0
    Item(index).AddEXP = 0
    Item(index).AttackSpeed = 1000
    Item(index).Price = 0
    Item(index).Stackable = 0
    Item(index).Bound = 0

End Sub

Sub ClearItems()

  Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i

End Sub

Sub ClearMap(ByVal MapNum As Long)

  Dim X As Long
  Dim Y As Long

    Map(MapNum).Name = vbNullString
    Map(MapNum).Revision = 0
    Map(MapNum).Moral = 0
    Map(MapNum).Up = 0
    Map(MapNum).Down = 0
    Map(MapNum).left = 0
    Map(MapNum).Right = 0
    Map(MapNum).Indoors = 0
    Map(MapNum).scrolling = GetVar(App.Path & "\Data.ini", "CONFIG", "Scrolling") + 1
    Map(MapNum).Weather = 0

    For X = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(X) = 0
    Next X

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map(MapNum).tile(X, Y).Ground = 0
            Map(MapNum).tile(X, Y).Mask = 0
            Map(MapNum).tile(X, Y).Anim = 0
            Map(MapNum).tile(X, Y).Mask2 = 0
            Map(MapNum).tile(X, Y).M2Anim = 0
            Map(MapNum).tile(X, Y).Fringe = 0
            Map(MapNum).tile(X, Y).FAnim = 0
            Map(MapNum).tile(X, Y).Fringe2 = 0
            Map(MapNum).tile(X, Y).F2Anim = 0
            Map(MapNum).tile(X, Y).Type = 0
            Map(MapNum).tile(X, Y).Data1 = 0
            Map(MapNum).tile(X, Y).Data2 = 0
            Map(MapNum).tile(X, Y).Data3 = 0
            Map(MapNum).tile(X, Y).String1 = vbNullString
            Map(MapNum).tile(X, Y).String2 = vbNullString
            Map(MapNum).tile(X, Y).String3 = vbNullString
            Map(MapNum).tile(X, Y).light = 0
            Map(MapNum).tile(X, Y).GroundSet = 0
            Map(MapNum).tile(X, Y).MaskSet = 0
            Map(MapNum).tile(X, Y).AnimSet = 0
            Map(MapNum).tile(X, Y).Mask2Set = 0
            Map(MapNum).tile(X, Y).M2AnimSet = 0
            Map(MapNum).tile(X, Y).FringeSet = 0
            Map(MapNum).tile(X, Y).FAnimSet = 0
            Map(MapNum).tile(X, Y).Fringe2Set = 0
            Map(MapNum).tile(X, Y).F2AnimSet = 0
        Next X

    Next Y

    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO

End Sub

Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Long)

    MapItem(MapNum, index).num = 0
    MapItem(MapNum, index).Value = 0
    MapItem(MapNum, index).Dur = 0
    MapItem(MapNum, index).X = 0
    MapItem(MapNum, index).Y = 0

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

Sub ClearMapNpc(ByVal index As Long, ByVal MapNum As Long)

    MapNpc(MapNum, index).num = 0
    MapNpc(MapNum, index).Target = 0
    MapNpc(MapNum, index).HP = 0
    MapNpc(MapNum, index).MP = 0
    MapNpc(MapNum, index).SP = 0
    MapNpc(MapNum, index).X = 0
    MapNpc(MapNum, index).Y = 0
    MapNpc(MapNum, index).Dir = 0

    ' Server use only
    MapNpc(MapNum, index).SpawnWait = 0
    MapNpc(MapNum, index).AttackTimer = 0

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

Sub ClearMaps()

  Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next i

End Sub

Sub ClearMapScroll(ByVal MapNum As Long)

  Dim X As Long
  Dim Y As Long

    Map(MapNum).scrolling = GetVar(App.Path & "\Data.ini", "CONFIG", "Scrolling") + 1
    Temp_Map(MapNum).BootMap = Map(MapNum).BootMap
    Temp_Map(MapNum).BootX = Map(MapNum).BootX
    Temp_Map(MapNum).BootY = Map(MapNum).BootY
    Temp_Map(MapNum).Down = Map(MapNum).Down
    Temp_Map(MapNum).Indoors = Map(MapNum).Indoors
    Temp_Map(MapNum).left = Map(MapNum).left
    Temp_Map(MapNum).Moral = Map(MapNum).Moral
    Temp_Map(MapNum).music = Map(MapNum).music
    Temp_Map(MapNum).Name = Map(MapNum).Name
    Temp_Map(MapNum).owner = Map(MapNum).owner
    Temp_Map(MapNum).Revision = Map(MapNum).Revision
    Temp_Map(MapNum).Right = Map(MapNum).Right
    Temp_Map(MapNum).Shop = Map(MapNum).Shop
    Temp_Map(MapNum).Up = Map(MapNum).Up
    Temp_Map(MapNum).Weather = Map(MapNum).Weather

    For X = 1 To MAX_MAP_NPCS
        Temp_Map(MapNum).Npc(X) = Map(MapNum).Npc(X)
    Next X

    For X = 0 To 19
        For Y = 0 To 14
            Temp_Map(MapNum).tile(X, Y).Ground = Map(MapNum).tile(X, Y).Ground
            Temp_Map(MapNum).tile(X, Y).Mask = Map(MapNum).tile(X, Y).Mask
            Temp_Map(MapNum).tile(X, Y).Anim = Map(MapNum).tile(X, Y).Anim
            Temp_Map(MapNum).tile(X, Y).Mask2 = Map(MapNum).tile(X, Y).Mask2
            Temp_Map(MapNum).tile(X, Y).M2Anim = Map(MapNum).tile(X, Y).M2Anim
            Temp_Map(MapNum).tile(X, Y).Fringe = Map(MapNum).tile(X, Y).Fringe
            Temp_Map(MapNum).tile(X, Y).FAnim = Map(MapNum).tile(X, Y).FAnim
            Temp_Map(MapNum).tile(X, Y).Fringe2 = Map(MapNum).tile(X, Y).Fringe2
            Temp_Map(MapNum).tile(X, Y).F2Anim = Map(MapNum).tile(X, Y).F2Anim
            Temp_Map(MapNum).tile(X, Y).Type = Map(MapNum).tile(X, Y).Type
            Temp_Map(MapNum).tile(X, Y).Data1 = Map(MapNum).tile(X, Y).Data1
            Temp_Map(MapNum).tile(X, Y).Data2 = Map(MapNum).tile(X, Y).Data2
            Temp_Map(MapNum).tile(X, Y).Data3 = Map(MapNum).tile(X, Y).Data3
            Temp_Map(MapNum).tile(X, Y).String1 = Map(MapNum).tile(X, Y).String1
            Temp_Map(MapNum).tile(X, Y).String2 = Map(MapNum).tile(X, Y).String2
            Temp_Map(MapNum).tile(X, Y).String3 = Map(MapNum).tile(X, Y).String3
            Temp_Map(MapNum).tile(X, Y).light = Map(MapNum).tile(X, Y).light
            Temp_Map(MapNum).tile(X, Y).GroundSet = Map(MapNum).tile(X, Y).GroundSet
            Temp_Map(MapNum).tile(X, Y).MaskSet = Map(MapNum).tile(X, Y).MaskSet
            Temp_Map(MapNum).tile(X, Y).AnimSet = Map(MapNum).tile(X, Y).AnimSet
            Temp_Map(MapNum).tile(X, Y).Mask2Set = Map(MapNum).tile(X, Y).Mask2Set
            Temp_Map(MapNum).tile(X, Y).M2AnimSet = Map(MapNum).tile(X, Y).M2AnimSet
            Temp_Map(MapNum).tile(X, Y).FringeSet = Map(MapNum).tile(X, Y).FringeSet
            Temp_Map(MapNum).tile(X, Y).FAnimSet = Map(MapNum).tile(X, Y).FAnimSet
            Temp_Map(MapNum).tile(X, Y).Fringe2Set = Map(MapNum).tile(X, Y).Fringe2Set
            Temp_Map(MapNum).tile(X, Y).F2AnimSet = Map(MapNum).tile(X, Y).F2AnimSet
        Next Y

    Next X

    Kill (App.Path & "\maps\map" & MapNum & ".dat")
    Call CheckMaps

    Map(MapNum).BootMap = Temp_Map(MapNum).BootMap
    Map(MapNum).BootX = Temp_Map(MapNum).BootX
    Map(MapNum).BootY = Temp_Map(MapNum).BootY
    Map(MapNum).Down = Temp_Map(MapNum).Down
    Map(MapNum).Indoors = Temp_Map(MapNum).Indoors
    Map(MapNum).left = Temp_Map(MapNum).left
    Map(MapNum).Moral = Temp_Map(MapNum).Moral
    Map(MapNum).music = Temp_Map(MapNum).music
    Map(MapNum).Name = Temp_Map(MapNum).Name
    Map(MapNum).owner = Temp_Map(MapNum).owner
    Map(MapNum).Revision = Temp_Map(MapNum).Revision
    Map(MapNum).Right = Temp_Map(MapNum).Right
    Map(MapNum).Shop = Temp_Map(MapNum).Shop
    Map(MapNum).Up = Temp_Map(MapNum).Up
    Map(MapNum).Weather = Temp_Map(MapNum).Weather

    For X = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(X) = Temp_Map(MapNum).Npc(X)
    Next X

    For X = 0 To 19
        For Y = 0 To 14
            Map(MapNum).tile(X, Y).Ground = Temp_Map(MapNum).tile(X, Y).Ground
            Map(MapNum).tile(X, Y).Mask = Temp_Map(MapNum).tile(X, Y).Mask
            Map(MapNum).tile(X, Y).Anim = Temp_Map(MapNum).tile(X, Y).Anim
            Map(MapNum).tile(X, Y).Mask2 = Temp_Map(MapNum).tile(X, Y).Mask2
            Map(MapNum).tile(X, Y).M2Anim = Temp_Map(MapNum).tile(X, Y).M2Anim
            Map(MapNum).tile(X, Y).Fringe = Temp_Map(MapNum).tile(X, Y).Fringe
            Map(MapNum).tile(X, Y).FAnim = Temp_Map(MapNum).tile(X, Y).FAnim
            Map(MapNum).tile(X, Y).Fringe2 = Temp_Map(MapNum).tile(X, Y).Fringe2
            Map(MapNum).tile(X, Y).F2Anim = Temp_Map(MapNum).tile(X, Y).F2Anim
            Map(MapNum).tile(X, Y).Type = Temp_Map(MapNum).tile(X, Y).Type
            Map(MapNum).tile(X, Y).Data1 = Temp_Map(MapNum).tile(X, Y).Data1
            Map(MapNum).tile(X, Y).Data2 = Temp_Map(MapNum).tile(X, Y).Data2
            Map(MapNum).tile(X, Y).Data3 = Temp_Map(MapNum).tile(X, Y).Data3
            Map(MapNum).tile(X, Y).String1 = Temp_Map(MapNum).tile(X, Y).String1
            Map(MapNum).tile(X, Y).String2 = Temp_Map(MapNum).tile(X, Y).String2
            Map(MapNum).tile(X, Y).String3 = Temp_Map(MapNum).tile(X, Y).String3
            Map(MapNum).tile(X, Y).light = Temp_Map(MapNum).tile(X, Y).light
            Map(MapNum).tile(X, Y).GroundSet = Temp_Map(MapNum).tile(X, Y).GroundSet
            Map(MapNum).tile(X, Y).MaskSet = Temp_Map(MapNum).tile(X, Y).MaskSet
            Map(MapNum).tile(X, Y).AnimSet = Temp_Map(MapNum).tile(X, Y).AnimSet
            Map(MapNum).tile(X, Y).Mask2Set = Temp_Map(MapNum).tile(X, Y).Mask2Set
            Map(MapNum).tile(X, Y).M2AnimSet = Temp_Map(MapNum).tile(X, Y).M2AnimSet
            Map(MapNum).tile(X, Y).FringeSet = Temp_Map(MapNum).tile(X, Y).FringeSet
            Map(MapNum).tile(X, Y).FAnimSet = Temp_Map(MapNum).tile(X, Y).FAnimSet
            Map(MapNum).tile(X, Y).Fringe2Set = Temp_Map(MapNum).tile(X, Y).Fringe2Set
            Map(MapNum).tile(X, Y).F2AnimSet = Temp_Map(MapNum).tile(X, Y).F2AnimSet
        Next Y

    Next X

    Call SaveMap(MapNum)

    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO

End Sub

Sub ClearNpc(ByVal index As Long)

  Dim i As Long

    Npc(index).Name = vbNullString
    Npc(index).AttackSay = vbNullString
    Npc(index).Sprite = 0
    Npc(index).SpawnSecs = 0
    Npc(index).Behavior = 0
    Npc(index).Range = 0
    Npc(index).STR = 0
    Npc(index).DEF = 0
    Npc(index).Speed = 0
    Npc(index).Magi = 0
    Npc(index).Big = 0
    Npc(index).MaxHp = 0
    Npc(index).Exp = 0
    Npc(index).SpawnTime = 0
    Npc(index).Element = 0

    For i = 1 To MAX_NPC_DROPS
        Npc(index).ItemNPC(i).Chance = 0
        Npc(index).ItemNPC(i).ItemNum = 0
        Npc(index).ItemNPC(i).ItemValue = 0
    Next i

End Sub

Sub ClearNpcs()

  Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next i

End Sub

Sub ClearPlayer(ByVal index As Long)

  Dim i As Long
  Dim n As Long

    Player(index).Login = vbNullString
    Player(index).Password = vbNullString

    For i = 1 To MAX_CHARS
        Player(index).Char(i).Name = vbNullString
        Player(index).Char(i).Class = 0
        Player(index).Char(i).Level = 0
        Player(index).Char(i).Sprite = 0
        Player(index).Char(i).Exp = 0
        Player(index).Char(i).access = 0
        Player(index).Char(i).PK = NO
        Player(index).Char(i).POINTS = 0
        Player(index).Char(i).Guild = vbNullString

        Player(index).Char(i).HP = 0
        Player(index).Char(i).MP = 0
        Player(index).Char(i).SP = 0

        Player(index).Char(i).STR = 0
        Player(index).Char(i).DEF = 0
        Player(index).Char(i).Speed = 0
        Player(index).Char(i).Magi = 0

        For n = 1 To MAX_INV
            Player(index).Char(i).Inv(n).num = 0
            Player(index).Char(i).Inv(n).Value = 0
            Player(index).Char(i).Inv(n).Dur = 0
        Next n

        For n = 1 To MAX_BANK
            Player(index).Char(i).Bank(n).num = 0
            Player(index).Char(i).Bank(n).Value = 0
            Player(index).Char(i).Bank(n).Dur = 0
        Next n

        For n = 1 To MAX_PLAYER_SPELLS
            Player(index).Char(i).Spell(n) = 0
        Next n

        Player(index).Char(i).ArmorSlot = 0
        Player(index).Char(i).WeaponSlot = 0
        Player(index).Char(i).HelmetSlot = 0
        Player(index).Char(i).ShieldSlot = 0
        Player(index).Char(i).LegsSlot = 0
        Player(index).Char(i).RingSlot = 0
        Player(index).Char(i).NecklaceSlot = 0

        Player(index).Char(i).Map = 0
        Player(index).Char(i).X = 0
        Player(index).Char(i).Y = 0
        Player(index).Char(i).Dir = 0

        Player(index).locked = False
        Player(index).lockedspells = False
        Player(index).lockeditems = False
        Player(index).lockedattack = False

        ' Temporary vars
        Player(index).Buffer = vbNullString
        Player(index).IncBuffer = vbNullString
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
        Player(index).Emoticon = -1
        Player(index).InTrade = 0
        Player(index).TradePlayer = 0
        Player(index).TradeOk = 0
        Player(index).TradeItemMax = 0
        Player(index).TradeItemMax2 = 0

        For n = 1 To MAX_PLAYER_TRADES
            Player(index).Trading(n).InvName = vbNullString
            Player(index).Trading(n).InvNum = 0
        Next n

        Player(index).ChatPlayer = 0
    Next i

End Sub

Sub ClearQuest(ByVal index As Long)

  Dim j As Long

    Quest(index).Name = vbNullString
    Quest(index).Pictop = 0
    Quest(index).Picleft = 0

    For j = 0 To MAX_QUEST_LENGHT
        Quest(index).Map(j) = 0
        Quest(index).X(j) = 0
        Quest(index).Y(j) = 0
        Quest(index).Npc(j) = 0
        Quest(index).Script(j) = 0
        Quest(index).ItemTake1num(j) = 0
        Quest(index).ItemTake2num(j) = 0
        Quest(index).ItemGive1num(j) = 0
        Quest(index).ItemGive2num(j) = 0
        Quest(index).ItemTake1val(j) = 0
        Quest(index).ItemTake2val(j) = 0
        Quest(index).ItemGive1val(j) = 0
        Quest(index).ItemGive2val(j) = 0
        Quest(index).ExpGiven(j) = 0
    Next j

End Sub

Sub ClearQuests()

  Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next i

End Sub

Sub ClearShop(ByVal index As Long)

  Dim i As Long

    Shop(index).Name = vbNullString
    Shop(index).currencyItem = 1
    Shop(index).FixesItems = 0
    Shop(index).ShowInfo = 0

    For i = 1 To MAX_SHOP_ITEMS
        Shop(index).ShopItem(i).ItemNum = 0
        Shop(index).ShopItem(i).Amount = 0
        Shop(index).ShopItem(i).Price = 0
    Next i

End Sub

Sub ClearShops()

  Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next i

End Sub

Sub ClearSkill(ByVal index As Long)

  Dim j As Long

    skill(index).Name = vbNullString
    skill(index).Action = vbNullString
    skill(index).Fail = vbNullString
    skill(index).Succes = vbNullString
    skill(index).AttemptName = vbNullString
    skill(index).Pictop = 0
    skill(index).Picleft = 0

    For j = 1 To MAX_SKILLS_SHEETS
        skill(index).ItemTake1num(j) = 0
        skill(index).ItemTake2num(j) = 0
        skill(index).ItemGive1num(j) = 0
        skill(index).ItemGive2num(j) = 0
        skill(index).minlevel(j) = 0
        skill(index).ExpGiven(j) = 0
        skill(index).ItemTake1val(j) = 0
        skill(index).ItemTake2val(j) = 0
        skill(index).ItemGive1val(j) = 0
        skill(index).ItemGive2val(j) = 0
        skill(index).itemequiped(j) = 0
    Next j

End Sub

Sub ClearSkills()

  Dim i As Long

    For i = 1 To MAX_SKILLS
        Call ClearSkill(i)
    Next i

End Sub

Sub ClearSpell(ByVal index As Long)

    Spell(index).Name = vbNullString
    Spell(index).ClassReq = 0
    Spell(index).LevelReq = 0
    Spell(index).Type = 0
    Spell(index).Data1 = 0
    Spell(index).Data2 = 0
    Spell(index).Data3 = 0
    Spell(index).MPCost = 0
    Spell(index).Sound = 0
    Spell(index).Range = 0

    Spell(index).SpellAnim = 0
    Spell(index).SpellTime = 40
    Spell(index).SpellDone = 1

    Spell(index).AE = 0
    Spell(index).Big = 0

    Spell(index).Element = 0

End Sub

Sub ClearSpells()

  Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next i

End Sub

Sub ClearTempTile()

  Dim i As Long
  Dim Y As Long
  Dim X As Long

    For i = 1 To MAX_MAPS
        TempTile(i).DoorTimer = 0

        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                TempTile(i).DoorOpen(X, Y) = NO
            Next X

        Next Y
    Next i

End Sub

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

Function GetClassName(ByVal ClassNum As Long) As String

    GetClassName = Trim(Class(ClassNum).Name)

End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long

    GetClassSPEED = Class(ClassNum).Speed

End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long

    GetClassSTR = Class(ClassNum).STR

End Function

Function GetPlayerAccess(ByVal index As Long) As Long

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

Function GetPlayerBody(ByVal index As Long) As Integer

    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerBody = Player(index).Char(Player(index).CharNum).body
    End If

End Function

Function GetPlayerClass(ByVal index As Long) As Long

    GetPlayerClass = Player(index).Char(Player(index).CharNum).Class

End Function

Function GetPlayerDEF(ByVal index As Long) As Long

  Dim add As Long

    add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddDef
    End If

    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddDef
    End If

    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddDef
    End If

    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddDef
    End If

    If GetPlayerLegsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).AddDef
    End If

    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddDef
    End If

    If GetPlayerNecklaceSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).AddDef
    End If

    GetPlayerDEF = Player(index).Char(Player(index).CharNum).DEF + add

End Function

Function GetPlayerDir(ByVal index As Long) As Long

    GetPlayerDir = Player(index).Char(Player(index).CharNum).Dir

End Function

Function GetPlayerExp(ByVal index As Long) As Long

    GetPlayerExp = Player(index).Char(Player(index).CharNum).Exp

End Function

Function GetPlayerGuild(ByVal index As Long) As String

    GetPlayerGuild = Trim(Player(index).Char(Player(index).CharNum).Guild)

End Function

Function GetPlayerGuildAccess(ByVal index As Long) As Long

    GetPlayerGuildAccess = Player(index).Char(Player(index).CharNum).Guildaccess

End Function

Function GetPlayerHead(ByVal index As Long) As Integer

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

    If InvSlot > 0 Then GetPlayerInvItemNum = Player(index).Char(Player(index).CharNum).Inv(InvSlot).num

End Function

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long

    GetPlayerInvItemValue = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value

End Function

Function GetPlayerIP(ByVal index As Long) As String

    GetPlayerIP = frmServer.Socket(index).RemoteHostIP

End Function

Function GetPlayerleg(ByVal index As Long) As Integer

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


' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String

    GetPlayerLogin = Trim(Player(index).Login)

End Function

Function GetPlayerMAGI(ByVal index As Long) As Long

  Dim add As Long

    add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddMagi
    End If

    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddMagi
    End If

    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddMagi
    End If

    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddMagi
    End If

    If GetPlayerLegsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).AddMagi
    End If

    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddMagi
    End If

    If GetPlayerNecklaceSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).AddMagi
    End If

    GetPlayerMAGI = Player(index).Char(Player(index).CharNum).Magi + add

End Function

Function GetPlayerMap(ByVal index As Long) As Long

    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerMap = Player(index).Char(Player(index).CharNum).Map
    End If

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
  Dim add As Long

    add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddSP
    End If

    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddSP
    End If

    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddSP
    End If

    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddSP
    End If

    If GetPlayerLegsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).AddSP
    End If

    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddSP
    End If

    If GetPlayerNecklaceSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).AddSP
    End If

    CharNum = Player(index).CharNum
    'GetPlayerMaxSP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSPEED(index) / 2) + Class(Player(index).Char(CharNum).Class).SPEED) * 2) + add
    GetPlayerMaxSP = (GetPlayerLevel(index) * AddSP.Level) + (GetPlayerSTR(index) * AddSP.STR) + (GetPlayerDEF(index) * AddSP.DEF) + (GetPlayerMAGI(index) * AddSP.Magi) + (GetPlayerSPEED(index) * AddSP.Speed) + add

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

    GetPlayerNextLevel = Experience(GetPlayerLevel(index))

End Function

Function GetPlayerPaperdoll(ByVal index As Long) As Byte

    If index < MAX_PLAYERS And index > 0 Then
        If Player(index).InGame Then
            GetPlayerPaperdoll = Player(index).Char(Player(index).CharNum).Paperdoll
        End If

    End If

End Function

Function GetPlayerPassword(ByVal index As Long) As String

    GetPlayerPassword = Trim(Player(index).Password)

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

Function GetPlayerSkillExp(ByVal index As Long, ByVal skill As Long) As Long

    If index > 0 And index < MAX_PLAYERS And IsPlaying(index) Then
        GetPlayerSkillExp = Player(index).Char(Player(index).CharNum).SkillExp(skill)
    End If

End Function

Function GetPlayerSkillLvl(ByVal index As Long, ByVal skill As Long) As Integer

    If index > 0 And index < MAX_PLAYERS And IsPlaying(index) Then
        GetPlayerSkillLvl = Player(index).Char(Player(index).CharNum).SkillLvl(skill)
    End If

End Function

Function GetPlayerSP(ByVal index As Long) As Long

    GetPlayerSP = Player(index).Char(Player(index).CharNum).SP

End Function

Function GetPlayerSPEED(ByVal index As Long) As Long

  Dim add As Long

    add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddSpeed
    End If

    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddSpeed
    End If

    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddSpeed
    End If

    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddSpeed
    End If

    If GetPlayerLegsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).AddSpeed
    End If

    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddSpeed
    End If

    If GetPlayerNecklaceSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).AddSpeed
    End If

    GetPlayerSPEED = Player(index).Char(Player(index).CharNum).Speed + add

End Function

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long) As Long

    GetPlayerSpell = Player(index).Char(Player(index).CharNum).Spell(SpellSlot)

End Function

Function GetPlayerSprite(ByVal index As Long) As Long

    GetPlayerSprite = Player(index).Char(Player(index).CharNum).Sprite

End Function

Function GetPlayerSTR(ByVal index As Long) As Long

  Dim add As Long

    add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddStr
    End If

    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddStr
    End If

    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddStr
    End If

    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddStr
    End If

    If GetPlayerLegsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).AddStr
    End If

    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddStr
    End If

    If GetPlayerNecklaceSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).AddStr
    End If

    GetPlayerSTR = Player(index).Char(Player(index).CharNum).STR + add

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

Sub HidePlayerPaperdoll(ByVal index As Long)

    If index < MAX_PLAYERS And index > 0 Then
        If Player(index).InGame Then
            Player(index).Char(Player(index).CharNum).Paperdoll = 0
        End If

    End If

End Sub

Function Rand(ByVal High As Long, ByVal Low As Long) As Long

    Randomize
    High = High + 1

    Do Until Rand >= Low
        Rand = Int(Rnd * High)
    Loop

End Function

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

Sub setplayerguild(ByVal index As Long, ByVal Guild As String)

    Player(index).Char(Player(index).CharNum).Guild = Guild

End Sub

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal Guildaccess As Long)

    Player(index).Char(Player(index).CharNum).Guildaccess = Guildaccess

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

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)

    Player(index).Login = Login

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

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)

    Player(index).Password = Password

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

    If Y >= 0 And Y <= MAX_MAPY Then Player(index).Char(Player(index).CharNum).Y = Y

End Sub

Sub ShowPlayerPaperdoll(ByVal index As Long)

    If index < MAX_PLAYERS And index > 0 Then
        If Player(index).InGame Then
            Player(index).Char(Player(index).CharNum).Paperdoll = 1
        End If

    End If

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

