Attribute VB_Name = "modTypes"
Option Explicit
Global PlayerI As Byte

Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long
    

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
Public Const ITEM_TYPE_TWO_HAND = 2
Public Const ITEM_TYPE_ARMOR = 3
Public Const ITEM_TYPE_HELMET = 4
Public Const ITEM_TYPE_SHIELD = 5
Public Const ITEM_TYPE_LEGS = 6
Public Const ITEM_TYPE_RING = 7
Public Const ITEM_TYPE_NECKLACE = 8
Public Const ITEM_TYPE_POTIONADDHP = 9
Public Const ITEM_TYPE_POTIONADDMP = 10
Public Const ITEM_TYPE_POTIONADDSP = 11
Public Const ITEM_TYPE_POTIONSUBHP = 12
Public Const ITEM_TYPE_POTIONSUBMP = 13
Public Const ITEM_TYPE_POTIONSUBSP = 14
Public Const ITEM_TYPE_KEY = 15
Public Const ITEM_TYPE_CURRENCY = 16
Public Const ITEM_TYPE_SPELL = 17
Public Const ITEM_TYPE_SCRIPTED = 18

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
    x(0 To MAX_QUEST_LENGHT) As Integer
    y(0 To MAX_QUEST_LENGHT) As Integer
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

Public Type PlayerRec
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
    x As Byte
    y As Byte
    Dir As Byte
    
    targetnpc As Integer
    
    head As Integer
    body As Integer
    leg As Integer
        
    SkillLvl() As Integer
    SkillExp() As Long
    
    Paperdoll As Byte
    
    MAXHP As Long
    MAXMP As Long
    MAXSP As Long
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
    
Public Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    Email As String

    ' Characters (we use 0 to prevent a crash that still needs to be figured out)
    Char(0 To MAX_CHARS) As PlayerRec
    
    ' None saved local vars
    Buffer As String
    IncBuffer As String
    charnum As Byte
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
    x As Byte
    y As Byte
    
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
    
    addHP As Long
    addMP As Long
    addSP As Long
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
    MAXHP As Long
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
        
    x As Byte
    y As Byte
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
Public Temp_Map() As Temp_MapRec
Public TempTile() As TempTileRec
Public PlayersOnMap() As Long
Public player() As AccountRec
Public tempplayer As AccountRec
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
Public addHP As StatRec
Public addMP As StatRec
Public addSP As StatRec

Public Const MAX_ATTRIBUTE_NPCS = 25
Type MapAttributeNpcRec
    num As Long
    
    Target As Long
    
    HP As Long
    MP As Long
    SP As Long
        
    x As Byte
    y As Byte
    Dir As Integer
    owner As Long
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
End Type
Public MapAttributeNpc() As MapAttributeNpcRec

Sub ClearTempTile()
Dim I As Long, y As Long, x As Long

    For I = 1 To MAX_MAPS
        TempTile(I).DoorTimer = 0
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                TempTile(I).DoorOpen(x, y) = NO
            Next x
        Next y
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
        Class(I).Speed = 0
        Class(I).Magi = 0
        Class(I).FemaleSprite = 0
        Class(I).MaleSprite = 0
        Class(I).Map = 0
        Class(I).x = 0
        Class(I).y = 0
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
        Class2(I).Speed = 0
        Class2(I).Magi = 0
        Class2(I).FemaleSprite = 0
        Class2(I).MaleSprite = 0
        Class2(I).Map = 0
        Class2(I).x = 0
        Class2(I).y = 0
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
        Class3(I).Speed = 0
        Class3(I).Magi = 0
        Class3(I).FemaleSprite = 0
        Class3(I).MaleSprite = 0
        Class3(I).Map = 0
        Class3(I).x = 0
        Class3(I).y = 0
    Next I
End Sub

Sub CleartempPlayer()
Dim I As Long
Dim N As Long

    tempplayer.Login = vbNullString
    tempplayer.Password = vbNullString
    For I = 1 To MAX_CHARS
        tempplayer.Char(I).Name = vbNullString
        tempplayer.Char(I).Class = 0
        tempplayer.Char(I).Level = 0
        tempplayer.Char(I).Sprite = 0
        tempplayer.Char(I).Exp = 0
        tempplayer.Char(I).access = 0
        tempplayer.Char(I).PK = NO
        tempplayer.Char(I).POINTS = 0
        tempplayer.Char(I).Guild = vbNullString
        
        tempplayer.Char(I).HP = 0
        tempplayer.Char(I).MP = 0
        tempplayer.Char(I).SP = 0
        
         tempplayer.Char(I).MAXHP = 0
        tempplayer.Char(I).MAXMP = 0
        tempplayer.Char(I).MAXSP = 0
        
        tempplayer.Char(I).STR = 0
        tempplayer.Char(I).DEF = 0
        tempplayer.Char(I).Speed = 0
        tempplayer.Char(I).Magi = 0
        
        For N = 1 To MAX_INV
            tempplayer.Char(I).Inv(N).num = 0
            tempplayer.Char(I).Inv(N).Value = 0
            tempplayer.Char(I).Inv(N).Dur = 0
        Next N
            For N = 1 To MAX_BANK
       tempplayer.Char(I).Bank(N).num = 0
       tempplayer.Char(I).Bank(N).Value = 0
       tempplayer.Char(I).Bank(N).Dur = 0
   Next N
        For N = 1 To MAX_PLAYER_SPELLS
            tempplayer.Char(I).Spell(N) = 0
        Next N
        
        tempplayer.Char(I).ArmorSlot = 0
        tempplayer.Char(I).WeaponSlot = 0
        tempplayer.Char(I).HelmetSlot = 0
        tempplayer.Char(I).ShieldSlot = 0
        tempplayer.Char(I).LegsSlot = 0
        tempplayer.Char(I).RingSlot = 0
        tempplayer.Char(I).NecklaceSlot = 0
        
        tempplayer.Char(I).Map = 0
        tempplayer.Char(I).x = 0
        tempplayer.Char(I).y = 0
        tempplayer.Char(I).Dir = 0
        
        tempplayer.locked = False
        tempplayer.lockedspells = False
        tempplayer.lockeditems = False
        tempplayer.lockedattack = False
        
        ' Temporary vars
        tempplayer.Buffer = vbNullString
        tempplayer.IncBuffer = vbNullString
        tempplayer.charnum = 0
        tempplayer.InGame = False
        tempplayer.AttackTimer = 0
        tempplayer.DataTimer = 0
        tempplayer.DataBytes = 0
        tempplayer.DataPackets = 0
        tempplayer.PartyPlayer = 0
        tempplayer.InParty = 0
        tempplayer.Target = 0
        tempplayer.TargetType = 0
        tempplayer.CastedSpell = NO
        tempplayer.PartyStarter = NO
        tempplayer.GettingMap = NO
        tempplayer.Emoticon = -1
        tempplayer.InTrade = 0
        tempplayer.TradePlayer = 0
        tempplayer.TradeOk = 0
        tempplayer.TradeItemMax = 0
        tempplayer.TradeItemMax2 = 0
        For N = 1 To MAX_PLAYER_TRADES
            tempplayer.Trading(N).InvName = vbNullString
            tempplayer.Trading(N).InvNum = 0
        Next N
        tempplayer.ChatPlayer = 0
        Next I
End Sub

Sub ClearPlayer(ByVal index As Long)
Dim I As Long
Dim N As Long

    player(index).Login = vbNullString
    player(index).Password = vbNullString
    For I = 1 To MAX_CHARS
        player(index).Char(I).Name = vbNullString
        player(index).Char(I).Class = 0
        player(index).Char(I).Level = 0
        player(index).Char(I).Sprite = 0
        player(index).Char(I).Exp = 0
        player(index).Char(I).access = 0
        player(index).Char(I).PK = NO
        player(index).Char(I).POINTS = 0
        player(index).Char(I).Guild = vbNullString
        
        player(index).Char(I).HP = 0
        player(index).Char(I).MP = 0
        player(index).Char(I).SP = 0
        
        player(index).Char(I).MAXHP = 0
        player(index).Char(I).MAXMP = 0
        player(index).Char(I).MAXSP = 0
        
        player(index).Char(I).STR = 0
        player(index).Char(I).DEF = 0
        player(index).Char(I).Speed = 0
        player(index).Char(I).Magi = 0
        
        For N = 1 To MAX_INV
            player(index).Char(I).Inv(N).num = 0
            player(index).Char(I).Inv(N).Value = 0
            player(index).Char(I).Inv(N).Dur = 0
        Next N
            For N = 1 To MAX_BANK
       player(index).Char(I).Bank(N).num = 0
       player(index).Char(I).Bank(N).Value = 0
       player(index).Char(I).Bank(N).Dur = 0
   Next N
        For N = 1 To MAX_PLAYER_SPELLS
            player(index).Char(I).Spell(N) = 0
        Next N
        
        player(index).Char(I).ArmorSlot = 0
        player(index).Char(I).WeaponSlot = 0
        player(index).Char(I).HelmetSlot = 0
        player(index).Char(I).ShieldSlot = 0
        player(index).Char(I).LegsSlot = 0
        player(index).Char(I).RingSlot = 0
        player(index).Char(I).NecklaceSlot = 0
        
        player(index).Char(I).Map = 0
        player(index).Char(I).x = 0
        player(index).Char(I).y = 0
        player(index).Char(I).Dir = 0
        
        player(index).locked = False
        player(index).lockedspells = False
        player(index).lockeditems = False
        player(index).lockedattack = False
        
        ' Temporary vars
        player(index).Buffer = vbNullString
        player(index).IncBuffer = vbNullString
        player(index).charnum = 0
        player(index).InGame = False
        player(index).AttackTimer = 0
        player(index).DataTimer = 0
        player(index).DataBytes = 0
        player(index).DataPackets = 0
        player(index).PartyPlayer = 0
        player(index).InParty = 0
        player(index).Target = 0
        player(index).TargetType = 0
        player(index).CastedSpell = NO
        player(index).PartyStarter = NO
        player(index).GettingMap = NO
        player(index).Emoticon = -1
        player(index).InTrade = 0
        player(index).TradePlayer = 0
        player(index).TradeOk = 0
        player(index).TradeItemMax = 0
        player(index).TradeItemMax2 = 0
        For N = 1 To MAX_PLAYER_TRADES
            player(index).Trading(N).InvName = vbNullString
            player(index).Trading(N).InvNum = 0
        Next N
        player(index).ChatPlayer = 0
    Next I
    
End Sub

Sub ClearChar(ByVal index As Long, ByVal charnum As Long)
Dim N As Long
    
    player(index).Char(charnum).Name = vbNullString
    player(index).Char(charnum).Class = 0
    player(index).Char(charnum).Sprite = 0
    player(index).Char(charnum).Level = 0
    player(index).Char(charnum).Exp = 0
    player(index).Char(charnum).access = 0
    player(index).Char(charnum).PK = NO
    player(index).Char(charnum).POINTS = 0
    player(index).Char(charnum).Guild = vbNullString
    
    player(index).Char(charnum).HP = 0
    player(index).Char(charnum).MP = 0
    player(index).Char(charnum).SP = 0
    
    player(index).Char(charnum).MAXHP = 0
    player(index).Char(charnum).MAXMP = 0
    player(index).Char(charnum).MAXSP = 0
    
    player(index).Char(charnum).STR = 0
    player(index).Char(charnum).DEF = 0
    player(index).Char(charnum).Speed = 0
    player(index).Char(charnum).Magi = 0
    
    For N = 1 To MAX_INV
        player(index).Char(charnum).Inv(N).num = 0
        player(index).Char(charnum).Inv(N).Value = 0
        player(index).Char(charnum).Inv(N).Dur = 0
    Next N
        For N = 1 To MAX_BANK
       player(index).Char(charnum).Bank(N).num = 0
       player(index).Char(charnum).Bank(N).Value = 0
       player(index).Char(charnum).Bank(N).Dur = 0
   Next N
    For N = 1 To MAX_PLAYER_SPELLS
        player(index).Char(charnum).Spell(N) = 0
    Next N
    
    player(index).Char(charnum).ArmorSlot = 0
    player(index).Char(charnum).WeaponSlot = 0
    player(index).Char(charnum).HelmetSlot = 0
    player(index).Char(charnum).ShieldSlot = 0
    player(index).Char(charnum).LegsSlot = 0
    player(index).Char(charnum).RingSlot = 0
    player(index).Char(charnum).NecklaceSlot = 0
    
    player(index).Char(charnum).Map = 0
    player(index).Char(charnum).x = 0
    player(index).Char(charnum).y = 0
    player(index).Char(charnum).Dir = 0
End Sub

Sub CleartempChar(ByVal charnum As Long)
Dim N As Long
    
    tempplayer.Char(charnum).Name = vbNullString
    tempplayer.Char(charnum).Class = 0
    tempplayer.Char(charnum).Sprite = 0
    tempplayer.Char(charnum).Level = 0
    tempplayer.Char(charnum).Exp = 0
    tempplayer.Char(charnum).access = 0
    tempplayer.Char(charnum).PK = NO
    tempplayer.Char(charnum).POINTS = 0
    tempplayer.Char(charnum).Guild = vbNullString
    
    tempplayer.Char(charnum).HP = 0
    tempplayer.Char(charnum).MP = 0
    tempplayer.Char(charnum).SP = 0
    
    tempplayer.Char(charnum).MAXHP = 0
    tempplayer.Char(charnum).MAXMP = 0
    tempplayer.Char(charnum).MAXSP = 0
    
    tempplayer.Char(charnum).STR = 0
    tempplayer.Char(charnum).DEF = 0
    tempplayer.Char(charnum).Speed = 0
    tempplayer.Char(charnum).Magi = 0
    
    For N = 1 To MAX_INV
        tempplayer.Char(charnum).Inv(N).num = 0
        tempplayer.Char(charnum).Inv(N).Value = 0
        tempplayer.Char(charnum).Inv(N).Dur = 0
    Next N
        For N = 1 To MAX_BANK
       tempplayer.Char(charnum).Bank(N).num = 0
       tempplayer.Char(charnum).Bank(N).Value = 0
       tempplayer.Char(charnum).Bank(N).Dur = 0
   Next N
    For N = 1 To MAX_PLAYER_SPELLS
        tempplayer.Char(charnum).Spell(N) = 0
    Next N
    
    tempplayer.Char(charnum).ArmorSlot = 0
    tempplayer.Char(charnum).WeaponSlot = 0
    tempplayer.Char(charnum).HelmetSlot = 0
    tempplayer.Char(charnum).ShieldSlot = 0
    tempplayer.Char(charnum).LegsSlot = 0
    tempplayer.Char(charnum).RingSlot = 0
    tempplayer.Char(charnum).NecklaceSlot = 0
    
    tempplayer.Char(charnum).Map = 0
    tempplayer.Char(charnum).x = 0
    tempplayer.Char(charnum).y = 0
    tempplayer.Char(charnum).Dir = 0
End Sub

    
Sub ClearQuest(ByVal index As Long)
Dim j As Long
    
    Quest(index).Name = vbNullString
    Quest(index).Pictop = 0
    Quest(index).Picleft = 0
    
    For j = 0 To MAX_QUEST_LENGHT
        Quest(index).Map(j) = 0
        Quest(index).x(j) = 0
        Quest(index).y(j) = 0
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
Sub ClearQuests()
Dim I As Long

    For I = 1 To MAX_QUESTS
        Call ClearQuest(I)
    Next I
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
Dim I As Long

    For I = 1 To MAX_SKILLS
        Call ClearSkill(I)
    Next I
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
    
    Item(index).addHP = 0
    Item(index).addMP = 0
    Item(index).addSP = 0
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
Dim I As Long

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next I
End Sub

Sub ClearNpc(ByVal index As Long)
Dim I As Long
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
    Npc(index).MAXHP = 0
    Npc(index).Exp = 0
    Npc(index).SpawnTime = 0
    Npc(index).Element = 0
    
    For I = 1 To MAX_NPC_DROPS
        Npc(index).ItemNPC(I).Chance = 0
        Npc(index).ItemNPC(I).ItemNum = 0
        Npc(index).ItemNPC(I).ItemValue = 0
    Next I
    
End Sub

Sub ClearNpcs()
Dim I As Long

    For I = 1 To MAX_NPCS
        Call ClearNpc(I)
    Next I
End Sub

Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Long)
    MapItem(MapNum, index).num = 0
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
    MapNpc(MapNum, index).num = 0
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

Sub ClearMapScroll(ByVal MapNum As Long)
Dim x As Long
Dim y As Long

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
    
    For x = 1 To MAX_MAP_NPCS
        Temp_Map(MapNum).Npc(x) = Map(MapNum).Npc(x)
    Next x
    
    For x = 0 To 19
        For y = 0 To 14
            Temp_Map(MapNum).tile(x, y).Ground = Map(MapNum).tile(x, y).Ground
            Temp_Map(MapNum).tile(x, y).Mask = Map(MapNum).tile(x, y).Mask
            Temp_Map(MapNum).tile(x, y).Anim = Map(MapNum).tile(x, y).Anim
            Temp_Map(MapNum).tile(x, y).Mask2 = Map(MapNum).tile(x, y).Mask2
            Temp_Map(MapNum).tile(x, y).M2Anim = Map(MapNum).tile(x, y).M2Anim
            Temp_Map(MapNum).tile(x, y).Fringe = Map(MapNum).tile(x, y).Fringe
            Temp_Map(MapNum).tile(x, y).FAnim = Map(MapNum).tile(x, y).FAnim
            Temp_Map(MapNum).tile(x, y).Fringe2 = Map(MapNum).tile(x, y).Fringe2
            Temp_Map(MapNum).tile(x, y).F2Anim = Map(MapNum).tile(x, y).F2Anim
            Temp_Map(MapNum).tile(x, y).Type = Map(MapNum).tile(x, y).Type
            Temp_Map(MapNum).tile(x, y).Data1 = Map(MapNum).tile(x, y).Data1
            Temp_Map(MapNum).tile(x, y).Data2 = Map(MapNum).tile(x, y).Data2
            Temp_Map(MapNum).tile(x, y).Data3 = Map(MapNum).tile(x, y).Data3
            Temp_Map(MapNum).tile(x, y).String1 = Map(MapNum).tile(x, y).String1
            Temp_Map(MapNum).tile(x, y).String2 = Map(MapNum).tile(x, y).String2
            Temp_Map(MapNum).tile(x, y).String3 = Map(MapNum).tile(x, y).String3
            Temp_Map(MapNum).tile(x, y).light = Map(MapNum).tile(x, y).light
            Temp_Map(MapNum).tile(x, y).GroundSet = Map(MapNum).tile(x, y).GroundSet
            Temp_Map(MapNum).tile(x, y).MaskSet = Map(MapNum).tile(x, y).MaskSet
            Temp_Map(MapNum).tile(x, y).AnimSet = Map(MapNum).tile(x, y).AnimSet
            Temp_Map(MapNum).tile(x, y).Mask2Set = Map(MapNum).tile(x, y).Mask2Set
            Temp_Map(MapNum).tile(x, y).M2AnimSet = Map(MapNum).tile(x, y).M2AnimSet
            Temp_Map(MapNum).tile(x, y).FringeSet = Map(MapNum).tile(x, y).FringeSet
            Temp_Map(MapNum).tile(x, y).FAnimSet = Map(MapNum).tile(x, y).FAnimSet
            Temp_Map(MapNum).tile(x, y).Fringe2Set = Map(MapNum).tile(x, y).Fringe2Set
            Temp_Map(MapNum).tile(x, y).F2AnimSet = Map(MapNum).tile(x, y).F2AnimSet
        Next y
    Next x
    
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
    
    For x = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(x) = Temp_Map(MapNum).Npc(x)
    Next x

    For x = 0 To 19
        For y = 0 To 14
            Map(MapNum).tile(x, y).Ground = Temp_Map(MapNum).tile(x, y).Ground
            Map(MapNum).tile(x, y).Mask = Temp_Map(MapNum).tile(x, y).Mask
            Map(MapNum).tile(x, y).Anim = Temp_Map(MapNum).tile(x, y).Anim
            Map(MapNum).tile(x, y).Mask2 = Temp_Map(MapNum).tile(x, y).Mask2
            Map(MapNum).tile(x, y).M2Anim = Temp_Map(MapNum).tile(x, y).M2Anim
            Map(MapNum).tile(x, y).Fringe = Temp_Map(MapNum).tile(x, y).Fringe
            Map(MapNum).tile(x, y).FAnim = Temp_Map(MapNum).tile(x, y).FAnim
            Map(MapNum).tile(x, y).Fringe2 = Temp_Map(MapNum).tile(x, y).Fringe2
            Map(MapNum).tile(x, y).F2Anim = Temp_Map(MapNum).tile(x, y).F2Anim
            Map(MapNum).tile(x, y).Type = Temp_Map(MapNum).tile(x, y).Type
            Map(MapNum).tile(x, y).Data1 = Temp_Map(MapNum).tile(x, y).Data1
            Map(MapNum).tile(x, y).Data2 = Temp_Map(MapNum).tile(x, y).Data2
            Map(MapNum).tile(x, y).Data3 = Temp_Map(MapNum).tile(x, y).Data3
            Map(MapNum).tile(x, y).String1 = Temp_Map(MapNum).tile(x, y).String1
            Map(MapNum).tile(x, y).String2 = Temp_Map(MapNum).tile(x, y).String2
            Map(MapNum).tile(x, y).String3 = Temp_Map(MapNum).tile(x, y).String3
            Map(MapNum).tile(x, y).light = Temp_Map(MapNum).tile(x, y).light
            Map(MapNum).tile(x, y).GroundSet = Temp_Map(MapNum).tile(x, y).GroundSet
            Map(MapNum).tile(x, y).MaskSet = Temp_Map(MapNum).tile(x, y).MaskSet
            Map(MapNum).tile(x, y).AnimSet = Temp_Map(MapNum).tile(x, y).AnimSet
            Map(MapNum).tile(x, y).Mask2Set = Temp_Map(MapNum).tile(x, y).Mask2Set
            Map(MapNum).tile(x, y).M2AnimSet = Temp_Map(MapNum).tile(x, y).M2AnimSet
            Map(MapNum).tile(x, y).FringeSet = Temp_Map(MapNum).tile(x, y).FringeSet
            Map(MapNum).tile(x, y).FAnimSet = Temp_Map(MapNum).tile(x, y).FAnimSet
            Map(MapNum).tile(x, y).Fringe2Set = Temp_Map(MapNum).tile(x, y).Fringe2Set
            Map(MapNum).tile(x, y).F2AnimSet = Temp_Map(MapNum).tile(x, y).F2AnimSet
        Next y
    Next x

    Call SaveMap(MapNum)
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
End Sub


Sub ClearMap(ByVal MapNum As Long)
Dim x As Long
Dim y As Long

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
    
    For x = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(x) = 0
    Next x
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Map(MapNum).tile(x, y).Ground = 0
            Map(MapNum).tile(x, y).Mask = 0
            Map(MapNum).tile(x, y).Anim = 0
            Map(MapNum).tile(x, y).Mask2 = 0
            Map(MapNum).tile(x, y).M2Anim = 0
            Map(MapNum).tile(x, y).Fringe = 0
            Map(MapNum).tile(x, y).FAnim = 0
            Map(MapNum).tile(x, y).Fringe2 = 0
            Map(MapNum).tile(x, y).F2Anim = 0
            Map(MapNum).tile(x, y).Type = 0
            Map(MapNum).tile(x, y).Data1 = 0
            Map(MapNum).tile(x, y).Data2 = 0
            Map(MapNum).tile(x, y).Data3 = 0
            Map(MapNum).tile(x, y).String1 = vbNullString
            Map(MapNum).tile(x, y).String2 = vbNullString
            Map(MapNum).tile(x, y).String3 = vbNullString
            Map(MapNum).tile(x, y).light = 0
            Map(MapNum).tile(x, y).GroundSet = 0
            Map(MapNum).tile(x, y).MaskSet = 0
            Map(MapNum).tile(x, y).AnimSet = 0
            Map(MapNum).tile(x, y).Mask2Set = 0
            Map(MapNum).tile(x, y).M2AnimSet = 0
            Map(MapNum).tile(x, y).FringeSet = 0
            Map(MapNum).tile(x, y).FAnimSet = 0
            Map(MapNum).tile(x, y).Fringe2Set = 0
            Map(MapNum).tile(x, y).F2AnimSet = 0
        Next x
    Next y
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
End Sub

Sub ClearMaps()
Dim I As Long

    For I = 1 To MAX_MAPS
        Call ClearMap(I)
    Next I
End Sub

Sub ClearShop(ByVal index As Long)
Dim I As Long

    Shop(index).Name = vbNullString
    Shop(index).currencyItem = 1
    Shop(index).FixesItems = 0
    Shop(index).ShowInfo = 0
    For I = 1 To MAX_SHOP_ITEMS
        Shop(index).ShopItem(I).ItemNum = 0
        Shop(index).ShopItem(I).Amount = 0
        Shop(index).ShopItem(I).Price = 0
    Next I
    
End Sub

Sub ClearShops()
Dim I As Long

    For I = 1 To MAX_SHOPS
        Call ClearShop(I)
    Next I
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
Dim I As Long

    For I = 1 To MAX_SPELLS
        Call ClearSpell(I)
    Next I
End Sub




' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim(player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    GetPlayerPassword = Trim(player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String
    GetPlayerName = Trim(player(index).Char(player(index).charnum).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    player(index).Char(player(index).charnum).Name = Name
End Sub

Function GetPlayerGuild(ByVal index As Long) As String
    GetPlayerGuild = Trim(player(index).Char(player(index).charnum).Guild)
End Function

Sub setplayerguild(ByVal index As Long, ByVal Guild As String)
    player(index).Char(player(index).charnum).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal index As Long) As Long
    GetPlayerGuildAccess = player(index).Char(player(index).charnum).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal Guildaccess As Long)
    player(index).Char(player(index).charnum).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = player(index).Char(player(index).charnum).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    player(index).Char(player(index).charnum).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    GetPlayerSprite = player(index).Char(player(index).charnum).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    If index > 0 And index <= MAX_PLAYERS Then
        player(index).Char(player(index).charnum).Sprite = Sprite
    End If
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    GetPlayerLevel = player(index).Char(player(index).charnum).Level
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
    player(index).Char(player(index).charnum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = Experience(GetPlayerLevel(index))
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = player(index).Char(player(index).charnum).Exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)
    player(index).Char(player(index).charnum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    GetPlayerAccess = player(index).Char(player(index).charnum).access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal access As Long)
    player(index).Char(player(index).charnum).access = access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    GetPlayerPK = player(index).Char(player(index).charnum).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    player(index).Char(player(index).charnum).PK = PK
End Sub

Function GetPlayerHP(ByVal index As Long) As Long
    GetPlayerHP = player(index).Char(player(index).charnum).HP
End Function

Sub SetPlayerHP(ByVal index As Long, ByVal HP As Long)
    player(index).Char(player(index).charnum).HP = HP
    
    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then
        player(index).Char(player(index).charnum).HP = GetPlayerMaxHP(index)
    End If
    If GetPlayerHP(index) < 0 Then
        player(index).Char(player(index).charnum).HP = 0
    End If
    Call SendStats(index)
End Sub

Function GetPlayerMP(ByVal index As Long) As Long
    GetPlayerMP = player(index).Char(player(index).charnum).MP
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
    player(index).Char(player(index).charnum).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        player(index).Char(player(index).charnum).MP = GetPlayerMaxMP(index)
    End If
    If GetPlayerMP(index) < 0 Then
        player(index).Char(player(index).charnum).MP = 0
    End If
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
    GetPlayerSP = player(index).Char(player(index).charnum).SP
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
    player(index).Char(player(index).charnum).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        player(index).Char(player(index).charnum).SP = GetPlayerMaxSP(index)
    End If
    If GetPlayerSP(index) < 0 Then
        player(index).Char(player(index).charnum).SP = 0
    End If
End Sub


Function GetPlayerMaxHP(ByVal index As Long) As Long
Dim charnum As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).addHP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).addHP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).addHP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).addHP
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).addHP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).addHP
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).addHP
    End If

    charnum = player(index).charnum
    'GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSTR(index) / 2) + Class(Player(index).Char(CharNum).Class).STR) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(index) * addHP.Level) + (GetPlayerSTR(index) * addHP.STR) + (GetPlayerDEF(index) * addHP.DEF) + (GetPlayerMAGI(index) * addHP.Magi) + (GetPlayerSPEED(index) * addHP.Speed) + add
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
Dim charnum As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).addMP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).addMP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).addMP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).addMP
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).addMP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).addMP
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).addMP
    End If

    charnum = player(index).charnum
    'GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(index) * addMP.Level) + (GetPlayerSTR(index) * addMP.STR) + (GetPlayerDEF(index) * addMP.DEF) + (GetPlayerMAGI(index) * addMP.Magi) + (GetPlayerSPEED(index) * addMP.Speed) + add
End Function



Function GetPlayerMaxSP(ByVal index As Long) As Long
Dim charnum As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).addSP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).addSP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).addSP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).addSP
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).addSP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).addSP
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).addSP
    End If

    charnum = player(index).charnum
    'GetPlayerMaxSP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSPEED(index) / 2) + Class(Player(index).Char(CharNum).Class).SPEED) * 2) + add
    GetPlayerMaxSP = (GetPlayerLevel(index) * addSP.Level) + (GetPlayerSTR(index) * addSP.STR) + (GetPlayerDEF(index) * addSP.DEF) + (GetPlayerMAGI(index) * addSP.Magi) + (GetPlayerSPEED(index) * addSP.Speed) + add
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
    GetClassMaxSP = (1 + Int(Class(ClassNum).Speed / 2) + Class(ClassNum).Speed) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = Class(ClassNum).STR
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
    GetPlayerSTR = player(index).Char(player(index).charnum).STR + add
End Function

Sub SetPlayerSTR(ByVal index As Long, ByVal STR As Long)
    player(index).Char(player(index).charnum).STR = STR
End Sub

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
    GetPlayerDEF = player(index).Char(player(index).charnum).DEF + add
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal DEF As Long)
    player(index).Char(player(index).charnum).DEF = DEF
End Sub

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
    GetPlayerSPEED = player(index).Char(player(index).charnum).Speed + add
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal Speed As Long)
    player(index).Char(player(index).charnum).Speed = Speed
End Sub

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
    GetPlayerMAGI = player(index).Char(player(index).charnum).Magi + add
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal Magi As Long)
    player(index).Char(player(index).charnum).Magi = Magi
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    GetPlayerPOINTS = player(index).Char(player(index).charnum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    player(index).Char(player(index).charnum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerMap = player(index).Char(player(index).charnum).Map
    End If
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        player(index).Char(player(index).charnum).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    GetPlayerX = player(index).Char(player(index).charnum).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    player(index).Char(player(index).charnum).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    GetPlayerY = player(index).Char(player(index).charnum).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    If y >= 0 And y <= MAX_MAPY Then player(index).Char(player(index).charnum).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    GetPlayerDir = player(index).Char(player(index).charnum).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    player(index).Char(player(index).charnum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
If InvSlot > 0 Then GetPlayerInvItemNum = player(index).Char(player(index).charnum).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    player(index).Char(player(index).charnum).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = player(index).Char(player(index).charnum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    player(index).Char(player(index).charnum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = player(index).Char(player(index).charnum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    player(index).Char(player(index).charnum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = player(index).Char(player(index).charnum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    player(index).Char(player(index).charnum).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerArmorSlot(ByVal index As Long) As Long
    GetPlayerArmorSlot = player(index).Char(player(index).charnum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)
    player(index).Char(player(index).charnum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal index As Long) As Long
    GetPlayerWeaponSlot = player(index).Char(player(index).charnum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)
    player(index).Char(player(index).charnum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long
    GetPlayerHelmetSlot = player(index).Char(player(index).charnum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)
    player(index).Char(player(index).charnum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long
    GetPlayerShieldSlot = player(index).Char(player(index).charnum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)
    player(index).Char(player(index).charnum).ShieldSlot = InvNum
End Sub
Function GetPlayerLegsSlot(ByVal index As Long) As Long
    GetPlayerLegsSlot = player(index).Char(player(index).charnum).LegsSlot
End Function

Sub SetPlayerLegsSlot(ByVal index As Long, InvNum As Long)
    player(index).Char(player(index).charnum).LegsSlot = InvNum
End Sub
Function GetPlayerRingSlot(ByVal index As Long) As Long
    GetPlayerRingSlot = player(index).Char(player(index).charnum).RingSlot
End Function

Sub SetPlayerRingSlot(ByVal index As Long, InvNum As Long)
    player(index).Char(player(index).charnum).RingSlot = InvNum
End Sub
Function GetPlayerNecklaceSlot(ByVal index As Long) As Long
    GetPlayerNecklaceSlot = player(index).Char(player(index).charnum).NecklaceSlot
End Function

Sub SetPlayerNecklaceSlot(ByVal index As Long, InvNum As Long)
    player(index).Char(player(index).charnum).NecklaceSlot = InvNum
End Sub

Sub BattleMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Byte)
    Call SendDataTo(index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR)
End Sub

Function Rand(ByVal High As Long, ByVal Low As Long) As Long
Randomize
High = High + 1
Do Until Rand >= Low
    Rand = Int(Rnd * High)
Loop
End Function
Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
   GetPlayerBankItemNum = player(index).Char(player(index).charnum).Bank(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
   player(index).Char(player(index).charnum).Bank(BankSlot).num = ItemNum
   Call SendBankUpdate(index, BankSlot)
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
   GetPlayerBankItemValue = player(index).Char(player(index).charnum).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
   player(index).Char(player(index).charnum).Bank(BankSlot).Value = ItemValue
   Call SendBankUpdate(index, BankSlot)
End Sub

Function GetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long) As Long
   GetPlayerBankItemDur = player(index).Char(player(index).charnum).Bank(BankSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
   player(index).Char(player(index).charnum).Bank(BankSlot).Dur = ItemDur
End Sub

Function GetPlayerHead(ByVal index As Long) As Integer
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerHead = player(index).Char(player(index).charnum).head
    End If
End Function

Sub SetPlayerHead(ByVal index As Long, ByVal head As Long)
    If index > 0 And index < MAX_PLAYERS Then
        player(index).Char(player(index).charnum).head = head
    End If
End Sub

Function GetPlayerBody(ByVal index As Long) As Integer
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerBody = player(index).Char(player(index).charnum).body
    End If
End Function

Sub SetPlayerBody(ByVal index As Long, ByVal body As Long)
    If index > 0 And index < MAX_PLAYERS Then
        player(index).Char(player(index).charnum).body = body
    End If
End Sub


Function GetPlayerleg(ByVal index As Long) As Integer
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerleg = player(index).Char(player(index).charnum).leg
    End If
End Function

Sub SetPlayerLeg(ByVal index As Long, ByVal leg As Long)
    If index > 0 And index < MAX_PLAYERS Then
        player(index).Char(player(index).charnum).leg = leg
    End If
End Sub

Function GetPlayerSkillLvl(ByVal index As Long, ByVal skill As Long) As Integer
    If index > 0 And index < MAX_PLAYERS And IsPlaying(index) Then
        GetPlayerSkillLvl = player(index).Char(player(index).charnum).SkillLvl(skill)
    End If
End Function

Sub SetPlayerSkillLvl(ByVal index As Long, ByVal skill As Long, ByVal lvl As Long)
    If index > 0 And index < MAX_PLAYERS Then
        player(index).Char(player(index).charnum).SkillLvl(skill) = lvl
    End If
End Sub

Function GetPlayerSkillExp(ByVal index As Long, ByVal skill As Long) As Long
    If index > 0 And index < MAX_PLAYERS And IsPlaying(index) Then
        GetPlayerSkillExp = player(index).Char(player(index).charnum).SkillExp(skill)
    End If
End Function

Sub SetPlayerSkillExp(ByVal index As Long, ByVal skill As Long, ByVal lvl As Long)
    If index > 0 And index < MAX_PLAYERS Then
        player(index).Char(player(index).charnum).SkillExp(skill) = lvl
    End If
End Sub

Function GetPlayerPaperdoll(ByVal index As Long) As Byte
    If index < MAX_PLAYERS And index > 0 Then
        If player(index).InGame Then
            GetPlayerPaperdoll = player(index).Char(player(index).charnum).Paperdoll
        End If
    End If
End Function

Sub ShowPlayerPaperdoll(ByVal index As Long)
    If index < MAX_PLAYERS And index > 0 Then
        If player(index).InGame Then
            player(index).Char(player(index).charnum).Paperdoll = 1
        End If
    End If
End Sub

Sub HidePlayerPaperdoll(ByVal index As Long)
    If index < MAX_PLAYERS And index > 0 Then
        If player(index).InGame Then
            player(index).Char(player(index).charnum).Paperdoll = 0
        End If
    End If
End Sub
