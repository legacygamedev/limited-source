Attribute VB_Name = "modTypes"

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

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
Public MAX_SPEECH As Long
Public MAX_ELEMENTS As Long

Public Const MAX_ARROWS = 100
Public Const MAX_PLAYER_ARROWS = 100

Public Const MAX_INV = 24
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 66
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10
Public Const MAX_SPEECH_OPTIONS = 20
Public Const MAX_FRIENDS = 20
Public Const MAX_BANK = 50

Public Const NO = 0
Public Const YES = 1

' Version constants
Public Const CLIENT_MAJOR = 1
Public Const CLIENT_MINOR = 1
Public Const CLIENT_REVISION = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Security password
Public Const SEC_CODE = "89h89hr98hewf9wfnd3nf98b9s8enfs09fn390jnf83n"

' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map constants
'Public Const MAX_MAPX = 30
'Public Const MAX_MAPY = 30
Public MAX_MAPX As Variant
Public MAX_MAPY As Variant
Public Const SCREEN_X = 19
Public Const SCREEN_Y = 13
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_NO_PENALTY = 2
Public Const MAP_MORAL_HOUSE = 3

' Image constants
Public Const PIC_X = 32
Public Const PIC_Y = 32

' Size constants (of player sprites)
Public Const SIZE_X = 32
Public Const SIZE_Y = 64

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
Public Const TILE_TYPE_NONE = 20
Public Const TILE_TYPE_BANK = 23
Public Const TILE_TYPE_HOUSE_BUY = 24
Public Const TILE_TYPE_HOUSE = 25
Public Const TILE_TYPE_FURNITURE = 26


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
Public Const DISPLAY_BUBBLE_TIME As Long = 2000 ' In milliseconds.
Public DISPLAY_BUBBLE_WIDTH As Byte
Public Const MAX_BUBBLE_WIDTH As Byte = 6 ' In tiles. Includes corners.
Public Const MAX_LINE_LENGTH As Byte = 23 ' In characters.
Public Const MAX_LINES As Byte = 3

' Spell constants
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_PET = 6
Public Const SPELL_TYPE_SCRIPTED = 7

' Target type constants
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1
Public Const TARGET_TYPE_LOCATION = 2
Public Const TARGET_TYPE_PET = 3

' Emoticon type constants
Public Const EMOTICON_TYPE_IMAGE = 0
Public Const EMOTICON_TYPE_SOUND = 1
Public Const EMOTICON_TYPE_BOTH = 2

' Encrypted GFX Password
Public Const GFX_PASSWORD = "test"

Type ElementRec
    name As String * NAME_LENGTH
    Strong As Long
    Weak As Long
End Type

Type BankRec
    num As Long
    Value As Long
    Dur As Long
End Type

Type ChatBubble
    Text As String
    Created As Long
End Type

Type PlayerInvRec
    num As Long
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
End Type

Type PetRec
    Sprite As Long
   
    Alive As Byte
   
    HP As Long
    MaxHP As Long
   
    Map As Long
    x As Long
    y As Long
    Dir As Byte
   
    Moving As Byte
    XOffset As Long
    YOffset As Long
   
    AttackTimer As Long
    Attacking As Byte
   
    LastAttack As Long
End Type

Type PlayerRec
    ' General
    name As String * NAME_LENGTH
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
    speed As Long
    MAGI As Long
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
    
    ' Pet!
    Pet As PetRec
    
    ' Client use only
    MaxHP As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    MovingH As Integer
    MovingV As Integer
    Attacking As Byte
    AttackTimer As Long
    LastAttack As Long
    MapGetTimer As Long
    CastedSpell As Byte
    
    SpellNum As Long
    SpellAnim() As SpellAnimRec

    EmoticonNum As Long
    EmoticonSound As String
    EmoticonType As Long
    EmoticonTime As Long
    EmoticonVar As Long
    EmoticonPlayed As Boolean
    
    LevelUp As Long
    LevelUpT As Long
    
    ArmorNum As Long
    WeaponNum As Long
    ShieldNum As Long
    HelmetNum As Long
    LegsNum As Long
    BootsNum As Long
    GlovesNum As Long
    Ring1Num As Long
    Ring2Num As Long
    AmuletNum As Long

    Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec
    Hands As Long
    
    Alignment As Long
    SenseAlignment As Long
    SenseAlignmentTime As Long
    
    CorpseMap As Integer
    CorpseX As Byte
    CorpseY As Byte
    CorpseLoot(1 To 4) As PlayerInvRec
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
    name As String * 40
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
    name As String * NAME_LENGTH
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
End Type

Type ItemRec
    name As String * NAME_LENGTH
    desc As String * 150
    
    pic As Long
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
    name As String * NAME_LENGTH
    AttackSay As String * 100
    
    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
        
    STR  As Long
    DEF As Long
    speed As Long
    MAGI As Long
    Big As Long
    MaxHP As Long
    EXP As Long
    SpawnTime As Long
    
    Speech As Long
    
    Script As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
    Element As Long
End Type

Type MapNpcRec
    num As Long
    
    Target As Long
    
    HP As Long
    MaxHP As Long
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
    name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
    FixesItems As Byte
    TradeItem(1 To 6) As TradeItemsRec
End Type

Type SpellRec
    name As String * NAME_LENGTH
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
    pic As Long
    Element As Long
End Type

Type TempTileRec
    DoorOpen As Byte
    
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

Type PlayerTradeRec
    InvNum As Long
    InvName As String
    InvVal As Long
End Type
Public Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public Trading2(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

Type EmoRec
    pic As Long
    Sound As String
    Command As String
    Type As Byte
End Type

Type OptionRec
    Text As String
    GoTo As Long
    Exit As Byte
End Type

Type InvSpeechRec
    Exit As Byte
    Text As String
    SaidBy As Byte
    Respond As Byte
    Script As Long
    Responces(1 To 3) As OptionRec
End Type

Type SpeechRec
    name As String
    num(0 To MAX_SPEECH_OPTIONS) As InvSpeechRec
End Type

Type DropRainRec
    x As Long
    y As Long
    Randomized As Boolean
    speed As Byte
End Type

' Bubble thing
Public Bubble() As ChatBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1
Public NEXT_CHAR As String * 1

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
Public Element() As ElementRec
Public Emoticons() As EmoRec
Public MapReport() As MapRec
Public Speech() As SpeechRec


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
    name As String
    pic As Long
    Range As Byte
End Type
Public Arrows(1 To MAX_ARROWS) As ArrowRec

Type BattleMsgRec
    Msg As String
    Index As Byte
    Color As Byte
    Time As Long
    done As Byte
    y As Long
End Type
Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Type ItemDurRec
    Item As Long
    Dur As Long
    done As Byte
End Type
Public ItemDur(1 To 4) As ItemDurRec

Public TempNpcSpawn(1 To MAX_MAP_NPCS) As LocRec

Public Inventory As Long
Public SpellIndex As Long
Public SpellMemorized As Long
Public charselsprite(MAX_CHARS) As Double


Sub ClearTempTile()
Dim x As Long, y As Long

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            TempTile(x, y).DoorOpen = NO
            
            TempTile(x, y).Ground = 0
            TempTile(x, y).Mask = 0
            TempTile(x, y).Anim = 0
            TempTile(x, y).Mask2 = 0
            TempTile(x, y).M2Anim = 0
            TempTile(x, y).Fringe = 0
            TempTile(x, y).FAnim = 0
            TempTile(x, y).Fringe2 = 0
            TempTile(x, y).F2Anim = 0
            TempTile(x, y).Type = TILE_TYPE_NONE
            TempTile(x, y).Data1 = 0
            TempTile(x, y).Data2 = 0
            TempTile(x, y).Data3 = 0
            TempTile(x, y).String1 = ""
            TempTile(x, y).String2 = ""
            TempTile(x, y).String3 = ""
            TempTile(x, y).Light = 0
            TempTile(x, y).GroundSet = 0
            TempTile(x, y).MaskSet = 0
            TempTile(x, y).AnimSet = 0
            TempTile(x, y).Mask2Set = 0
            TempTile(x, y).M2AnimSet = 0
            TempTile(x, y).FringeSet = 0
            TempTile(x, y).FAnimSet = 0
            TempTile(x, y).Fringe2Set = 0
            TempTile(x, y).F2AnimSet = 0
        Next x
    Next y
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim i As Long
Dim n As Long

    Player(Index).name = ""
    Player(Index).Guild = ""
    Player(Index).Guildaccess = 0
    Player(Index).Class = 1
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
    Player(Index).Hands = 0
    Player(Index).LegsSlot = 0
    Player(Index).BootsSlot = 0
    Player(Index).GlovesSlot = 0
    Player(Index).Ring1Slot = 0
    Player(Index).Ring2Slot = 0
    Player(Index).AmuletSlot = 0
        
    Player(Index).Map = 0
    Player(Index).x = 0
    Player(Index).y = 0
    Player(Index).Dir = 0
    
    ' Client use only
    Player(Index).MaxHP = 0
    Player(Index).MaxMP = 0
    Player(Index).MaxSP = 0
    Player(Index).XOffset = 0
    Player(Index).YOffset = 0
    Player(Index).MovingH = 0
    Player(Index).MovingV = 0
    Player(Index).Attacking = 0
    Player(Index).AttackTimer = 0
    Player(Index).MapGetTimer = 0
    Player(Index).CastedSpell = NO
    Player(Index).EmoticonNum = -1
    Player(Index).EmoticonSound = ""
    Player(Index).EmoticonType = 0
    Player(Index).EmoticonTime = 0
    Player(Index).EmoticonVar = 0
    Player(Index).EmoticonPlayed = True
    
    For i = 1 To MAX_SPELL_ANIM
        Player(Index).SpellAnim(i).CastedSpell = NO
        Player(Index).SpellAnim(i).SpellTime = 0
        Player(Index).SpellAnim(i).SpellVar = 0
        Player(Index).SpellAnim(i).SpellDone = 0
        
        Player(Index).SpellAnim(i).Target = 0
        Player(Index).SpellAnim(i).TargetType = TARGET_TYPE_PLAYER
    Next i
    
    Player(Index).SpellNum = 0
    
    Player(Index).ArmorNum = 0
    Player(Index).WeaponNum = 0
    Player(Index).ShieldNum = 0
    Player(Index).HelmetNum = 0
    Player(Index).LegsNum = 0
    Player(Index).BootsNum = 0
    Player(Index).GlovesNum = 0
    Player(Index).Ring1Num = 0
    Player(Index).Ring2Num = 0
    Player(Index).AmuletNum = 0
    
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i
    
    Player(Index).CorpseMap = 0
    Player(Index).CorpseX = 0
    Player(Index).CorpseY = 0
    For i = 1 To 4
    Player(Index).CorpseLoot(i).Dur = 0
    Player(Index).CorpseLoot(i).num = 0
    Player(Index).CorpseLoot(i).Value = 0
    Next i
    
    If CorpseIndex = Index Then
    CorpseIndex = 0
    End If
    
    Inventory = 1
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index).name = ""
    Item(Index).desc = ""
    
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
    Item(Index).AttackSpeed = 1000
    Item(Index).Stackable = 0
    Item(Index).Bound = 0
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
    MapItem(Index).x = 0
    MapItem(Index).y = 0
End Sub

Sub ClearMaps()
Dim i As Long

For i = 1 To MAX_MAPS
    Call ClearMap(i)
Next i
End Sub

Sub ClearMap(ByVal MapNum As Long)
Dim i, x, y As Long

    i = MapNum
    Map(i).name = ""
    Map(i).Revision = 0
    Map(i).Moral = 0
    Map(i).Up = 0
    Map(i).Down = 0
    Map(i).Left = 0
    Map(i).Right = 0
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
            Map(i).Tile(x, y).String1 = ""
            Map(i).Tile(x, y).String2 = ""
            Map(i).Tile(x, y).String3 = ""
            Map(i).Tile(x, y).Light = 0
            Map(i).Tile(x, y).GroundSet = -1
            Map(i).Tile(x, y).MaskSet = -1
            Map(i).Tile(x, y).AnimSet = -1
            Map(i).Tile(x, y).Mask2Set = -1
            Map(i).Tile(x, y).M2AnimSet = -1
            Map(i).Tile(x, y).FringeSet = -1
            Map(i).Tile(x, y).FAnimSet = -1
            Map(i).Tile(x, y).Fringe2Set = -1
            Map(i).Tile(x, y).F2AnimSet = -1
        Next x
    Next y
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

Sub ClearSpeech(ByVal Index As Long)
Dim i As Long
Dim O As Long

    Speech(Index).name = ""

    For O = 0 To MAX_SPEECH_OPTIONS
        Speech(Index).num(O).Exit = 0
        Speech(Index).num(O).Respond = 0
        Speech(Index).num(O).SaidBy = 0
        Speech(Index).num(O).Text = "Write what you want to be said here."
        Speech(Index).num(O).Script = 0
    
        For i = 1 To 3
            Speech(Index).num(O).Responces(i).Exit = 0
            Speech(Index).num(O).Responces(i).GoTo = 0
            Speech(Index).num(O).Responces(i).Text = "Write a responce here."
        Next i
    Next O
End Sub

Sub ClearSpeeches()
Dim i As Long

    For i = 1 To MAX_SPEECH
        Call ClearSpeech(i)
    Next i
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim(Player(Index).name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal name As String)
    Player(Index).name = name
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
    GetPlayerMaxHP = Player(Index).MaxHP
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
If Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
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

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    If BankSlot > MAX_BANK Then Exit Function
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

Sub SetPlayerHands(ByVal Index As Long, Item As Long)
    Player(Index).Hands = Item
    Call SendData("SETHANDS" & SEP_CHAR & Player(Index).Hands & SEP_CHAR & END_CHAR)
End Sub

Sub SetPlayerAlignment(ByVal Index As Long, num As Long)
    Player(Index).Alignment = num
End Sub

Function GetPlayerAlignment(ByVal Index As Long) As Long
    GetPlayerAlignment = Player(Index).Alignment
End Function

Function GetPlayerLegsSlot(ByVal Index As Long) As Long
    GetPlayerLegsSlot = Player(Index).LegsSlot
End Function

Sub SetPlayerLegsSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).LegsSlot = InvNum
End Sub

Function GetPlayerBootsSlot(ByVal Index As Long) As Long
    GetPlayerBootsSlot = Player(Index).BootsSlot
End Function

Sub SetPlayerBootsSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).BootsSlot = InvNum
End Sub

Function GetPlayerGlovesSlot(ByVal Index As Long) As Long
    GetPlayerGlovesSlot = Player(Index).GlovesSlot
End Function

Sub SetPlayerGlovesSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).GlovesSlot = InvNum
End Sub

Function GetPlayerRing1Slot(ByVal Index As Long) As Long
    GetPlayerRing1Slot = Player(Index).Ring1Slot
End Function

Sub SetPlayerRing1Slot(ByVal Index As Long, InvNum As Long)
    Player(Index).Ring1Slot = InvNum
End Sub

Function GetPlayerRing2Slot(ByVal Index As Long) As Long
    GetPlayerRing2Slot = Player(Index).Ring2Slot
End Function

Sub SetPlayerRing2Slot(ByVal Index As Long, InvNum As Long)
    Player(Index).Ring2Slot = InvNum
End Sub

Function GetPlayerAmuletSlot(ByVal Index As Long) As Long
    GetPlayerAmuletSlot = Player(Index).AmuletSlot
End Function

Sub SetPlayerAmuletSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).AmuletSlot = InvNum
End Sub

