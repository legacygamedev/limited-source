Attribute VB_Name = "modTypes"
Option Explicit

' General constants
Public GAME_NAME As String
Public WEBSITE As String
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_ELEMENTS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_SKILLS As Long
Public MAX_QUESTS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_EMOTICONS As Long
Public MAX_SPELL_ANIM As Long
Public MAX_BLT_LINE As Long
Public paperdoll As Long
Public Spritesize As Long
Public MAX_SCRIPTSPELLS As Long
Public ENCRYPT_PASS As String
Public ENCRYPT_TYPE As String
Public customplayers As Long
Public CUSTOM_TITLE As String
Public CUSTOM_IS_CLOSABLE As Long
Public MAX_PARTY_MEMBERS As Long
Public temp As Long
Public lvl As Long
Public STAT1 As String
Public STAT2 As String
Public STAT3 As String
Public STAT4 As String

Public Anim1Data As Long
Public Anim2Data As Long
Public M2AnimData As Long
Public FAnimData As Long
Public F2AnimData As Long

Public Screen_RESIZED As Long
Public screen_hplen As Long
Public screen_xplen As Long
Public screen_mplen As Long

Public Const MAX_ARROWS = 100
Public Const MAX_PLAYER_ARROWS = 100
Public Const MAX_BUBBLES = 20
Public Const MAX_BANK = 50
Public Const MAX_INV = 24
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

' OnClick tile info
Public ClickScript As Integer

' Map constants
'Public Const MAX_MAPX = 30
'Public Const MAX_MAPY = 30
Public MAX_MAPX As Integer
Public MAX_MAPY As Integer
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_NO_PENALTY = 2
Public Const MAP_MORAL_HOUSE = 3

'skill data
Public skill1
Public skill2
Public currentsheet As Long
Public ItemTake1num(1 To MAX_SKILLS_SHEETS) As Long
Public ItemTake2num(1 To MAX_SKILLS_SHEETS) As Long
Public ItemGive1num(1 To MAX_SKILLS_SHEETS) As Long
Public ItemGive2num(1 To MAX_SKILLS_SHEETS) As Long
Public minlevel(1 To MAX_SKILLS_SHEETS) As Long
Public ExpGiven(1 To MAX_SKILLS_SHEETS) As Long
Public base_chance(1 To MAX_SKILLS_SHEETS) As Long
Public ItemTake1val(1 To MAX_SKILLS_SHEETS) As Long
Public ItemTake2val(1 To MAX_SKILLS_SHEETS) As Long
Public ItemGive1val(1 To MAX_SKILLS_SHEETS) As Long
Public ItemGive2val(1 To MAX_SKILLS_SHEETS) As Long
Public itemequiped(1 To MAX_SKILLS_SHEETS) As Long

'Quest Data
Public Q_Map(0 To MAX_QUEST_LENGHT) As Long
Public Q_X(0 To MAX_QUEST_LENGHT) As Long
Public Q_Y(0 To MAX_QUEST_LENGHT) As Long
Public Q_Npc(0 To MAX_QUEST_LENGHT) As Long
Public Q_Script(0 To MAX_QUEST_LENGHT) As Long
Public Q_ItemTake1num(0 To MAX_QUEST_LENGHT) As Long
Public Q_ItemTake2num(0 To MAX_QUEST_LENGHT) As Long
Public Q_ItemGive1num(0 To MAX_QUEST_LENGHT) As Long
Public Q_ItemGive2num(0 To MAX_QUEST_LENGHT) As Long
Public Q_ItemTake1val(0 To MAX_QUEST_LENGHT) As Long
Public Q_ItemTake2val(0 To MAX_QUEST_LENGHT) As Long
Public Q_ItemGive1val(0 To MAX_QUEST_LENGHT) As Long
Public Q_ItemGive2val(0 To MAX_QUEST_LENGHT) As Long
Public Q_ExpGiven(0 To MAX_QUEST_LENGHT) As Long

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
Public Const ITEM_TYPE_THROW = 18
Public Const ITEM_TYPE_WARP = 19

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
Public Const SPELL_TYPE_SCRIPTED = 6
Public Const SPELL_TYPE_TEMP = 7

'Canon stuff
Public CanonItem As Long
Public CanonDamage As Long
Public CanonDirection As Long
Public CanonUsed As Long
Public CanonX As Long
Public CanonY As Long

'Minus Stat values
Public MinusHp As Integer
Public MinusMp As Integer
Public MinusSp As Integer
Public MessageMinus As String

'Music values

Type musictracker
    Song As String
End Type

Type ChatBubble
    Text As String
    Created As Long
End Type

Type ScriptBubble
    Text As String
    Created As Long
    Map As Long
    x As Long
    y As Long
    Colour As Long
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

Type ElementRec
    Name As String * NAME_LENGTH
    Strong As Long
    Weak As Long
End Type

Type SpellAnimRec
    CastedSpell As Byte

    SpellTime As Long
    SpellVar As Long
    SpellDone As Long

    Target As Long
    TargetType As Long
End Type

Type ScriptSpellAnimRec
    CastedSpell As Byte

    SpellTime As Long
    SpellVar As Long
    SpellDone As Long

    SpellNum As Long
    x As Long
    y As Long
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

Type PartyRec
    Leader As Byte
    Member() As Byte
    ShareExp As Boolean
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Class As Long
    Sprite As Long
    Level As Long
    Exp As Long
    Access As Byte
    PK As Byte
    input As Byte
    iso As Byte
    Canon As Long
    Party As PartyRec

    ' Vitals
    HP As Long
    MP As Long
    SP As Long

    ' Stats
    STR As Long
    DEF As Long
    speed As Long
    MAGI As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    LegsSlot As Long
    RingSlot As Long
    NecklaceSlot As Long

    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    Bank(1 To MAX_BANK) As BankRec

    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte

    ' Client use only
    MaxHp As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    MovingH As Integer
    MovingV As Integer
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

    SkilLvl() As Long
    SkilExp() As Long

    Armor As Long
    Helmet As Long
    Shield As Long
    Weapon As Long
    legs As Long
    Ring As Long
    Necklace As Long
    color As Long
    pet As Long

    head As Long
    body As Long
    leg As Long

    HookShotX As Long
    HookShotY As Long
    HookShotSucces As Long
    HookShotAnim As Long
    HookShotTime As Long
    HookShotToX As Long
    HookShotToY As Long
    HookShotDir As Long

    paperdoll As Byte
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
    Floor As Byte
    FloorSet As Byte
End Type

Type MapRec
    Name As String * 40
    Revision As Long
    Moral As Byte
    Up As Long
    Down As Long
    Left As Long
    right As Long
    Music As String
    BootMap As Long
    BootX As Byte
    BootY As Byte
    Shop As Long
    Indoors As Byte
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    Owner As String
    Weather As Long
    Lights As Long
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    MaleSprite As Long
    FemaleSprite As Long

    Locked As Long

    STR As Long
    DEF As Long
    speed As Long
    MAGI As Long

    ' For client use
    HP As Long
    MP As Long
    SP As Long

    ' Description
    desc As String
End Type

Type QuestRec
    Name As String * NAME_LENGTH
    Pictop As Long
    Picleft As Long
    Map(0 To MAX_QUEST_LENGHT) As Long
    x(0 To MAX_QUEST_LENGHT) As Long
    y(0 To MAX_QUEST_LENGHT) As Long
    Npc(0 To MAX_QUEST_LENGHT) As Long
    Script(0 To MAX_QUEST_LENGHT) As Long
    ItemTake1num(0 To MAX_QUEST_LENGHT) As Long
    ItemTake2num(0 To MAX_QUEST_LENGHT) As Long
    ItemTake1val(0 To MAX_QUEST_LENGHT) As Long
    ItemTake2val(0 To MAX_QUEST_LENGHT) As Long
    ItemGive1num(0 To MAX_QUEST_LENGHT) As Long
    ItemGive2num(0 To MAX_QUEST_LENGHT) As Long
    ItemGive1val(0 To MAX_QUEST_LENGHT) As Long
    ItemGive2val(0 To MAX_QUEST_LENGHT) As Long
    ExpGiven(0 To MAX_QUEST_LENGHT) As Long
End Type

Type SkillRec
    Name As String * NAME_LENGTH
    Action As String
    Fail As String
    Succes As String
    Pictop As Long
    Picleft As Long
    ItemTake1num(1 To MAX_SKILLS_SHEETS) As Long
    ItemTake2num(1 To MAX_SKILLS_SHEETS) As Long
    ItemGive1num(1 To MAX_SKILLS_SHEETS) As Long
    ItemGive2num(1 To MAX_SKILLS_SHEETS) As Long
    minlevel(1 To MAX_SKILLS_SHEETS) As Long
    ExpGiven(1 To MAX_SKILLS_SHEETS) As Long
    base_chance(1 To MAX_SKILLS_SHEETS) As Long
    ItemTake1val(1 To MAX_SKILLS_SHEETS) As Long
    ItemTake2val(1 To MAX_SKILLS_SHEETS) As Long
    ItemGive1val(1 To MAX_SKILLS_SHEETS) As Long
    ItemGive2val(1 To MAX_SKILLS_SHEETS) As Long
    itemequiped(1 To MAX_SKILLS_SHEETS) As Long
    'Added to bottom for compatiblity! What is displayed when the skill is attempted
    AttemptName As String
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
    Price As Long

    Stackable As Long
    Bound As Long
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
    chance As Long
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100

    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    Spritesize As Long

    STR  As Long
    DEF As Long
    speed As Long
    MAGI As Long
    Big As Long
    MaxHp As Long
    Exp As Long
    SpawnTime As Long
    Spell As Long

    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec

    Element As Long
End Type

Type MapNpcRec
    num As Long

    Target As Long

    HP As Long
    MaxHp As Long
    MP As Long
    SP As Long

    Map As Long
    x As Byte
    y As Byte
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
    Big As Long
    
    Element As Long
    reload As Long
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
    x As Long
    y As Long
    Randomized As Boolean
    speed As Byte
End Type

' Bubble thing
Public Bubble() As ChatBubble
Public ScriptBubble() As ScriptBubble

' Calender stuff
Public year As Long
Public month As Long
Public day As Long
Public weekday As Long

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

Public Map() As MapRec
Public TempTile() As TempTileRec
Public Player() As PlayerRec
Public Class() As ClassRec
Public item() As ItemRec
Public skill() As SkillRec
Public Quest() As QuestRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Element() As ElementRec
Public Emoticons() As EmoRec
Public MapReport() As MapRec
Public ScriptSpell() As ScriptSpellAnimRec

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

Public Trade(1 To 7) As TradeRec

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
    color As Byte
    Time As Long
    Done As Byte
    y As Long
End Type

Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Type ItemDurRec
    item As Long
    Dur As Long
    Done As Byte
End Type

Public ItemDur(1 To 4) As ItemDurRec

Public Inventory As Long
Public slot As Long

Public Const MAX_ATTRIBUTE_NPCS = 25
Public MapAttributeNpc() As MapNpcRec
Public SaveMapAttributeNpc() As MapNpcRec
Public Direct As Long
Public GuildBlock As String

Sub ClearItem(ByVal Index As Long)

    item(Index).Name = vbNullString
    item(Index).desc = vbNullString

    item(Index).Type = 0
    item(Index).Data1 = 0
    item(Index).Data2 = 0
    item(Index).Data3 = 0
    item(Index).StrReq = 0
    item(Index).DefReq = 0
    item(Index).SpeedReq = 0
    item(Index).ClassReq = -1
    item(Index).AccessReq = 0

    item(Index).AddHP = 0
    item(Index).AddMP = 0
    item(Index).AddSP = 0
    item(Index).AddStr = 0
    item(Index).AddDef = 0
    item(Index).AddMagi = 0
    item(Index).AddSpeed = 0
    item(Index).AddEXP = 0
    item(Index).AttackSpeed = 1000
    item(Index).Stackable = 0

End Sub

Sub ClearItems()

  Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i

End Sub

Sub ClearMap()

  Dim i As Long
  Dim x As Long
  Dim y As Long

    For i = 1 To MAX_MAPS
        Map(i).Name = vbNullString
        Map(i).Revision = 0
        Map(i).Moral = 0
        Map(i).Up = 0
        Map(i).Down = 0
        Map(i).Left = 0
        Map(i).right = 0
        Map(i).Indoors = 0

        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(i).Tile(x, y).Ground = 0
                Map(i).Tile(x, y).Mask = 0
                Map(i).Tile(x, y).Anim = 0
                Map(i).Tile(x, y).Mask2 = 0
                Map(i).Tile(x, y).M2Anim = 0
                Map(i).Tile(x, y).Fringe = 0
                Map(i).Tile(x, y).FAnim = 0
                Map(i).Tile(x, y).Fringe2 = 0
                Map(i).Tile(x, y).F2Anim = 0
                Map(i).Tile(x, y).Type = 0
                Map(i).Tile(x, y).Data1 = 0
                Map(i).Tile(x, y).Data2 = 0
                Map(i).Tile(x, y).Data3 = 0
                Map(i).Tile(x, y).String1 = vbNullString
                Map(i).Tile(x, y).String2 = vbNullString
                Map(i).Tile(x, y).String3 = vbNullString
                Map(i).Tile(x, y).light = 0
                Map(i).Tile(x, y).GroundSet = 0
                Map(i).Tile(x, y).MaskSet = 0
                Map(i).Tile(x, y).AnimSet = 0
                Map(i).Tile(x, y).Mask2Set = 0
                Map(i).Tile(x, y).M2AnimSet = 0
                Map(i).Tile(x, y).FringeSet = 0
                Map(i).Tile(x, y).FAnimSet = 0
                Map(i).Tile(x, y).Fringe2Set = 0
                Map(i).Tile(x, y).F2AnimSet = 0
                Map(i).Tile(x, y).Floor = 0
                Map(i).Tile(x, y).FloorSet = 0
            Next x

        Next y
    Next i

End Sub

Sub ClearMapAttributeNpc(ByVal Index As Long, ByVal x As Long, ByVal y As Long)

    MapAttributeNpc(Index, x, y).num = 0
    MapAttributeNpc(Index, x, y).Target = 0
    MapAttributeNpc(Index, x, y).HP = 0
    MapAttributeNpc(Index, x, y).MP = 0
    MapAttributeNpc(Index, x, y).SP = 0
    MapAttributeNpc(Index, x, y).Map = 0
    MapAttributeNpc(Index, x, y).x = 0
    MapAttributeNpc(Index, x, y).y = 0
    MapAttributeNpc(Index, x, y).Dir = 0

    ' Client use only
    MapAttributeNpc(Index, x, y).XOffset = 0
    MapAttributeNpc(Index, x, y).YOffset = 0
    MapAttributeNpc(Index, x, y).Moving = 0
    MapAttributeNpc(Index, x, y).Attacking = 0
    MapAttributeNpc(Index, x, y).AttackTimer = 0

End Sub

Sub ClearMapAttributeNpcs()

  Dim i As Long
  Dim x As Long
  Dim y As Long

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            For i = 1 To MAX_ATTRIBUTE_NPCS
                Call ClearMapAttributeNpc(i, x, y)
            Next i

        Next x
    Next y

End Sub

Sub ClearMapItem(ByVal Index As Long)

    MapItem(Index).num = 0
    MapItem(Index).Value = 0
    MapItem(Index).Dur = 0
    MapItem(Index).x = 0
    MapItem(Index).y = 0

End Sub

Sub ClearMapItems()

  Dim x As Long

    For x = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(x)
    Next x

End Sub

Sub ClearMapNpc(ByVal Index As Long)

    MapNpc(Index).num = 0
    MapNpc(Index).Target = 0
    MapNpc(Index).HP = 0
    MapNpc(Index).MP = 0
    MapNpc(Index).SP = 0
    MapNpc(Index).Map = 0
    MapNpc(Index).x = 0
    MapNpc(Index).y = 0
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
    Player(Index).x = 0
    Player(Index).y = 0
    Player(Index).Dir = 0

    ' Client use only
    Player(Index).MaxHp = 0
    Player(Index).MaxMP = 0
    Player(Index).MaxSP = 0
    Player(Index).XOffset = 0
    Player(Index).YOffset = 0
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

Sub ClearTempTile()

  Dim x As Long
  Dim y As Long

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            TempTile(x, y).DoorOpen = NO
        Next x

    Next y

End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long

    GetPlayerAccess = Player(Index).Access

End Function

Function GetPlayerArmorSlot(ByVal Index As Long) As Long

    GetPlayerArmorSlot = Player(Index).ArmorSlot

End Function

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long) As Long

    GetPlayerBankItemDur = Player(Index).Bank(BankSlot).Dur

End Function

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long

    If BankSlot > MAX_BANK Then Exit Function
    GetPlayerBankItemNum = Player(Index).Bank(BankSlot).num

End Function

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long

    GetPlayerBankItemValue = Player(Index).Bank(BankSlot).Value

End Function

Function GetPlayerBody(ByVal Index As Long)

    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerBody = Player(Index).body
    End If

End Function

Function GetPlayerClass(ByVal Index As Long) As Long

    GetPlayerClass = Player(Index).Class

End Function

Function GetPlayerDEF(ByVal Index As Long) As Long

    GetPlayerDEF = Player(Index).DEF

End Function

Function GetPlayerDir(ByVal Index As Long) As Long

    GetPlayerDir = Player(Index).Dir

End Function

Function GetPlayerExp(ByVal Index As Long) As Long

    GetPlayerExp = Player(Index).Exp

End Function

Function GetPlayerGuild(ByVal Index As Long) As String

    GetPlayerGuild = Trim$(Player(Index).Guild)

End Function

Function GetPlayerGuildAccess(ByVal Index As Long) As Long

    GetPlayerGuildAccess = Player(Index).Guildaccess

End Function

Function GetPlayerHead(ByVal Index As Long)

    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerHead = Player(Index).head
    End If

End Function

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long

    GetPlayerHelmetSlot = Player(Index).HelmetSlot

End Function

Function GetPlayerHP(ByVal Index As Long) As Long

    GetPlayerHP = Player(Index).HP

End Function

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long

    GetPlayerInvItemDur = Player(Index).Inv(InvSlot).Dur

End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long

    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).num

End Function

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long

    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).Value

End Function

Function GetPlayerleg(ByVal Index As Long)

    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerleg = Player(Index).leg
    End If

End Function

Function GetPlayerLegsSlot(ByVal Index As Long) As Long

    GetPlayerLegsSlot = Player(Index).LegsSlot

End Function

Function GetPlayerLevel(ByVal Index As Long) As Long

    GetPlayerLevel = Player(Index).Level

End Function

Function GetPlayerMAGI(ByVal Index As Long) As Long

    GetPlayerMAGI = Player(Index).MAGI

End Function

Function GetPlayerMap(ByVal Index As Long) As Long

    If Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map

End Function

Function GetPlayerMaxHP(ByVal Index As Long) As Long

    GetPlayerMaxHP = Player(Index).MaxHp

End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long

    GetPlayerMaxMP = Player(Index).MaxMP

End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long

    GetPlayerMaxSP = Player(Index).MaxSP

End Function

Function GetPlayerMP(ByVal Index As Long) As Long

    GetPlayerMP = Player(Index).MP

End Function

Function GetPlayerName(ByVal Index As Long) As String

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)

End Function

Function GetPlayerNecklaceSlot(ByVal Index As Long) As Long

    GetPlayerNecklaceSlot = Player(Index).NecklaceSlot

End Function

Function GetPlayerPaperdoll(ByVal Index As Long) As Byte

    If Index < MAX_PLAYERS And Index > 0 Then
        If IsPlaying(Index) Then
            GetPlayerPaperdoll = Player(Index).paperdoll
        End If

    End If

End Function

Function GetPlayerPK(ByVal Index As Long) As Long

    GetPlayerPK = Player(Index).PK

End Function

Function GetPlayerPOINTS(ByVal Index As Long) As Long

    GetPlayerPOINTS = Player(Index).POINTS

End Function

Function GetPlayerRingSlot(ByVal Index As Long) As Long

    GetPlayerRingSlot = Player(Index).RingSlot

End Function

Function GetPlayerShieldSlot(ByVal Index As Long) As Long

    GetPlayerShieldSlot = Player(Index).ShieldSlot

End Function

Function GetPlayerSkillExp(ByVal Index As Long, ByVal skill As Long)

    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerSkillExp = Player(Index).SkilExp(skill)
    End If

End Function

Function GetPlayerSkillLvl(ByVal Index As Long, ByVal skill As Long)

    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerSkillLvl = Player(Index).SkilLvl(skill)
    End If

End Function

Function GetPlayerSP(ByVal Index As Long) As Long

    GetPlayerSP = Player(Index).SP

End Function

Function GetPlayerSPEED(ByVal Index As Long) As Long

    GetPlayerSPEED = Player(Index).speed

End Function

Function GetPlayerSprite(ByVal Index As Long) As Long

    GetPlayerSprite = Player(Index).Sprite

End Function

Function GetPlayerSTR(ByVal Index As Long) As Long

    GetPlayerSTR = Player(Index).STR

End Function

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long

    GetPlayerWeaponSlot = Player(Index).WeaponSlot

End Function

Function GetPlayerX(ByVal Index As Long) As Long

    GetPlayerX = Player(Index).x

End Function

Function GetPlayerY(ByVal Index As Long) As Long

    GetPlayerY = Player(Index).y

End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)

    Player(Index).Access = Access

End Sub

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)

    Player(Index).ArmorSlot = InvNum

End Sub

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)

    Player(Index).Bank(BankSlot).Dur = ItemDur

End Sub

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)

    Player(Index).Bank(BankSlot).num = ItemNum

End Sub

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)

    Player(Index).Bank(BankSlot).Value = ItemValue

End Sub

Function SetPlayerBody(ByVal Index As Long, ByVal body As Long)

    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).body = body
    End If

End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)

    Player(Index).Class = ClassNum

End Sub

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)

    Player(Index).DEF = DEF

End Sub

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)

    Player(Index).Dir = Dir

End Sub

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)

    Player(Index).Exp = Exp

End Sub

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)

    Player(Index).Guild = Guild

End Sub

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Guildaccess As Long)

    Player(Index).Guildaccess = Guildaccess

End Sub

Function SetPlayerHead(ByVal Index As Long, ByVal head As Long)

    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).head = head
    End If

End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)

    Player(Index).HelmetSlot = InvNum

End Sub

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)

    Player(Index).HP = HP

    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).HP = GetPlayerMaxHP(Index)
    End If

End Sub

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)

    Player(Index).Inv(InvSlot).Dur = ItemDur

End Sub

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)

    Player(Index).Inv(InvSlot).num = ItemNum

End Sub

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)

    Player(Index).Inv(InvSlot).Value = ItemValue

End Sub

Function SetPlayerLeg(ByVal Index As Long, ByVal leg As Long)

    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).leg = leg
    End If

End Function

Sub SetPlayerLegsSlot(ByVal Index As Long, InvNum As Long)

    Player(Index).LegsSlot = InvNum

End Sub

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)

    Player(Index).Level = Level

End Sub

Sub SetPlayerLoc(ByVal Index As Long, ByVal x As Long, ByVal y As Long)

    Player(Index).x = x
    Player(Index).y = y

End Sub

Sub SetPlayerMAGI(ByVal Index As Long, ByVal MAGI As Long)

    Player(Index).MAGI = MAGI

End Sub

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)

    Player(Index).Map = MapNum

End Sub

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)

    Player(Index).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).MP = GetPlayerMaxMP(Index)
    End If

End Sub

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)

    Player(Index).Name = Name

End Sub

Sub SetPlayerNecklaceSlot(ByVal Index As Long, InvNum As Long)

    Player(Index).NecklaceSlot = InvNum

End Sub

Sub SetPlayerPaperdoll(ByVal Index As Long, ByVal paperdoll As Byte)

    If Index < MAX_PLAYERS And Index > 0 Then
        If IsPlaying(Index) Then
            Player(Index).paperdoll = paperdoll
        End If

    End If

End Sub

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)

    Player(Index).PK = PK

End Sub

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)

    Player(Index).POINTS = POINTS

End Sub

Sub SetPlayerRingSlot(ByVal Index As Long, InvNum As Long)

    Player(Index).RingSlot = InvNum

End Sub

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)

    Player(Index).ShieldSlot = InvNum

End Sub

Sub SetPlayerSkillExp(ByVal Index As Long, ByVal skill As Long, ByVal lvl As Long)

    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).SkilExp(skill) = lvl
    End If

End Sub

Sub SetPlayerSkillLvl(ByVal Index As Long, ByVal skill As Long, ByVal lvl As Long)

    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).SkilLvl(skill) = lvl
    End If

End Sub

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)

    Player(Index).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).SP = GetPlayerMaxSP(Index)
    End If

End Sub

Sub SetPlayerSPEED(ByVal Index As Long, ByVal speed As Long)

    Player(Index).speed = speed

End Sub

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)

    Player(Index).Sprite = Sprite

End Sub

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)

    Player(Index).STR = STR

End Sub

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)

    Player(Index).WeaponSlot = InvNum

End Sub

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)

    If x >= 0 And x <= MAX_MAPX Then Player(Index).x = x

End Sub

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)

    Player(Index).y = y

End Sub

