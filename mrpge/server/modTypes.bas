Attribute VB_Name = "modTypes"
Option Explicit


' Winsock globals
Public GAME_PORT
Public GAME_UPDATE_PORT

' General constants
Public GAME_NAME
Public Const MAX_PLAYERS = 30
Public Const MAX_ITEMS = 1020
Public Const MAX_NPCS = 255
Public Const MAX_INV = 50
Public Const MAX_BANK = 30
Public Const MAX_MAP_ITEMS = 20
Public Const MAX_MAP_NPCS = 14
Public Const MAX_SHOPS = 10
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_SPELLS = 255
Public Const MAX_TRADES = 20
Public Const MAX_GUILDS = 20
Public Const MAX_GUILD_MEMBERS = 25
Public Const MAX_SIGNS = 255
Public Const MAX_BOOSTS = 10
Public Const MAX_PETS = 30

' Summon Consts - for the ability to sumon things
Public Const MAX_SUMMONINGS = 30
Public Const MAX_SUMMONINGS_Player = 1

'Quests
Public Const MAX_QUESTS = 50

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const ITEM_NAME_LENGTH = 50
Public Const MAX_CHARS = 6

' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map constants
Public Const MAX_MAPS = 1000
Public Const MAX_MAPX = 15
Public Const MAX_MAPY = 11
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_ARENA = 2
Public Const MAP_MORAL_SAVAGE = 3

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
Public Const TILE_TYPE_WARP_LEVEL = 7
Public Const TILE_TYPE_DAMAGE = 8
Public Const TILE_TYPE_HEAL = 9
Public Const TILE_TYPE_SIGN = 10
Public Const TILE_TYPE_NPC_SPAWN = 11
Public Const TILE_TYPE_LEVEL = 12


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
Public Const ITEM_TYPE_PRAYER = 14
Public Const ITEM_TYPE_POTIONADDPP = 15

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

Public Const NPC_TYPE_NORMAL = 0
Public Const NPC_TYPE_UNDEAD = 1

' Spell constants
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_GIVEITEM = 6
Public Const SPELL_TYPE_SUMMON = 7

' Prayer Consts
Public Const PRAYER_TYPE_HEAL = 0
Public Const PRAYER_TYPE_CURE = 1
Public Const PRAYER_TYPE_ENHANCE = 2
Public Const PRAYER_TYPE_BOOST = 3

' Target type constants
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1

'Guild constants
Public Const COST_TO_CREATE_GUILD = 0
Public Const MEMBER_OF_GUILD = 1
Public Const LEADER_OF_GUILD = 2
Public Const FOUNDER_OF_GUILD = 3

Public Type NpcChatAns
    ans1 As String
    ans2 As String
    ans3 As String
    ans4 As String
End Type

Type PlayerInvRec
    num As Byte
    value As Long
    Dur As Integer
End Type

Type SignRec
    header As String
    msg As String
End Type

Type QuestRec
    ID As Long
    StartQuestMsg As String
    GetItemQuestMsg As String
    FinishQuestMessage As String
    ItemToObtain As Long
    ExpGiven As Long
    ItemGiven As Long
    ItemValGiven As Long
    requiredLevel As Long
    goldGiven As Long
End Type

Type PointBoostRec
    index As Long
    HP As Long
    MP As Long
    SP As Long
    PP As Long
    str As Byte
    intel As Byte
    dex As Byte
    con As Byte
    wiz As Byte
    cha As Byte
    isUsed As Boolean
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Sex As Byte
    Class As Byte
    sprite As Integer
    level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    Guild As Byte
    GuildAccess As Byte
    txtColour As Long
    ingameColour As Long
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    PP As Long ' prayer points :(
    lastSentHP As Long
    lastSentMP As Long
    lastSentSP As Long
    lastSentPP As Long
    
    ' Stats
    str As Byte
    intel As Byte
    dex As Byte
    con As Byte
    wiz As Byte
    cha As Byte
    'old stuff
    def As Byte
    speed As Byte
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
    Prayer(1 To MAX_PLAYER_SPELLS) As Byte
    Bank(1 To MAX_BANK) As PlayerInvRec
    ' Position
    map As Integer
    x As Byte
    y As Byte
    Dir As Byte

    'effects
    Poison As Boolean
    Poison_length As Long
    Poison_vital As Long
    
    'boosts
    Boosts(1 To MAX_BOOSTS) As PointBoostRec
    
    'quests
    CurrentQuest As Long
    QuestStatus As Long
    ingnoreBlocks As Boolean
    
    'PetId As Long
    'HasPet As Boolean
    'PetName As String
    'PetSprite As Long
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
    target As Byte
    CastedSpell As Byte
    PartyStarter As Byte
    GettingMap As Byte
    RealName As String
    Email As String
    Bio As String
End Type

Type TileRec
    Ground As Integer
    Mask As Integer
    Anim As Integer
    Fringe As Integer
    type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    Data4 As Integer
    Data5 As Integer
    'Data6 As Integer
    'Data7 As Integer
    'Data8 As Integer
    'Data9 As Integer
    'Data10 As Integer
    TileSheet_Ground As Byte
    TileSheet_Fringe As Byte
    TileSheet_Anim As Byte
    TileSheet_Mask As Byte
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
    Night As Byte
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
    Respawn As Boolean
    Bank As Boolean
End Type

Type MapRec
    Name As String * NAME_LENGTH
    street As String * NAME_LENGTH
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
    Night As Byte
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
    Respawn As Boolean
    Bank As Boolean
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    
    sprite As Integer
    
    str As Byte
    def As Byte
    speed As Byte
    MAGI As Byte
    intel As Byte
    dex As Byte
    con As Byte
    wiz As Byte
    cha As Byte
End Type

Type ItemRec
    Name As String * ITEM_NAME_LENGTH
    
    BaseDamage As Integer
    Pic As Integer
    type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    
    str As Integer
    intel As Integer
    dex As Integer
    con As Integer
    wiz As Integer
    cha As Integer
    
    Description As String
    Poisons As Boolean
    Poison_length As Long
    Poison_vital As Long
    weaponType As Long
    Boosts As PointBoostRec
End Type

Type MapItemRec
    num As Byte
    value As Long
    Dur As Integer
    
    x As Byte
    y As Byte
End Type


Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 255
    
    sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    
    str  As Byte
    def As Byte
    speed As Byte
    MAGI As Byte
    HP As Long
    
    Gold As Long
    ExpGiven As Long
    Respawn As Boolean
    Attack_with_Poison As Boolean
    Poison_length As Long
    Poison_vital As Long
    
    QuestID As Long
    
    opensShop As Boolean
    opensBank As Boolean
    
    type As Long
End Type

Type MapNpcRec
    num As Integer
    
    target As Integer
    
    HP As Long
    MP As Long
    SP As Long
        
    x As Byte
    y As Byte
    Dir As Integer
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    
    Gold As Long
    
    maxHP As Long
    
    TargetRunAway As Long
    TargetRunAwayTime As Long
    Respawn As Boolean
    Attack_with_Poison As Boolean
    Poison_length As Long
    Poison_vital As Long
End Type

Type SummonRec
    num As Integer
    owner As Integer
    target As Integer
    
    HP As Long
    MP As Long
    SP As Long
        
    x As Byte
    y As Byte
    Dir As Integer
    
    ' For server use only
    'SpawnWait As Long
    AttackTimer As Long
    
    'Gold As Long
    
    maxHP As Long
End Type

'Type PetRec
'    Name As String
'    owner As Integer
'    target As Integer
'
'    sprite As Long
'
'    HP As Long
'    MP As Long
'    SP As Long
'
'    map As Byte
'    x As Byte
'    y As Byte
'    Dir As Integer
'
'    AttackTimer As Long
'
'    maxHP As Long
'End Type

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
    type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    sound As Long
    manaUse As Long
End Type

Type PrayerRec
    Name As String * NAME_LENGTH
    ClassReq As Byte
    LevelReq As Byte
    type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    sound As Long
    manaUse As Long
    Boosts As PointBoostRec
End Type

Type TempTileRec
    DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY)  As Byte
    DoorTimer As Long
End Type

Type GuildRec
    Name As String '* NAME_LENGTH
    Founder As String '* NAME_LENGTH
    Member(1 To MAX_GUILD_MEMBERS) As String '* NAME_LENGTH
    Leaders(1 To MAX_GUILD_MEMBERS) As String
    Description As String
    InviteList As String
End Type


    

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

Public map(1 To MAX_MAPS) As MapRec
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public player(1 To MAX_PLAYERS) As AccountRec
Public Class() As ClassRec
Public Item(0 To MAX_ITEMS) As ItemRec
Public Npc(0 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Prayer(1 To MAX_SPELLS) As PrayerRec
Public Guild(1 To MAX_GUILDS) As GuildRec
Public Signs(1 To MAX_SIGNS) As SignRec
Public Summonings(1 To MAX_SUMMONINGS) As SummonRec
Public Quests(1 To MAX_QUESTS) As QuestRec
'Public Pets(1 To MAX_PETS) As PetRec


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
        Class(i).str = 0
        'Class(i).DEF = 0
        'Class(i).speed = 0
        'Class(i).MAGI = 0
        Class(i).intel = 0
        Class(i).dex = 0
        Class(i).con = 0
        Class(i).wiz = 0
        Class(i).cha = 0
    Next i
End Sub

Sub ClearPlayer(ByVal index As Long)
Dim i As Long
Dim n As Long

    player(index).Login = ""
    player(index).Password = ""
    
    For i = 1 To MAX_CHARS
        player(index).Char(i).Name = ""
        player(index).Char(i).Class = 0
        player(index).Char(i).level = 0
        player(index).Char(i).sprite = 0
        player(index).Char(i).Exp = 0
        player(index).Char(i).Access = 0
        player(index).Char(i).PK = NO
        player(index).Char(i).POINTS = 0
        player(index).Char(i).Guild = 0
        player(index).Char(i).GuildAccess = 0
        
        player(index).Char(i).HP = 0
        player(index).Char(i).MP = 0
        player(index).Char(i).SP = 0
        
        player(index).Char(i).str = 0
        player(index).Char(i).intel = 0
        player(index).Char(i).dex = 0
        player(index).Char(i).con = 0
        player(index).Char(i).wiz = 0
        player(index).Char(i).cha = 0
        'Player(index).Char(i).DEF = 0
        'Player(index).Char(i).speed = 0
        'Player(index).Char(i).MAGI = 0
        
        For n = 1 To MAX_INV
            player(index).Char(i).Inv(n).num = 0
            player(index).Char(i).Inv(n).value = 0
            player(index).Char(i).Inv(n).Dur = 0
        Next n
        For n = 1 To MAX_BANK
            player(index).Char(i).Bank(n).num = 0
            player(index).Char(i).Bank(n).value = 0
            player(index).Char(i).Bank(n).Dur = 0
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            player(index).Char(i).Spell(n) = 0
        Next n
        
        player(index).Char(i).ArmorSlot = 0
        player(index).Char(i).WeaponSlot = 0
        player(index).Char(i).HelmetSlot = 0
        player(index).Char(i).ShieldSlot = 0
        
        player(index).Char(i).map = 0
        player(index).Char(i).x = 0
        player(index).Char(i).y = 0
        player(index).Char(i).Dir = 0
        
        ' Temporary vars
        player(index).Buffer = ""
        player(index).IncBuffer = ""
        player(index).CharNum = 0
        player(index).InGame = False
        player(index).AttackTimer = 0
        player(index).DataTimer = 0
        player(index).DataBytes = 0
        player(index).DataPackets = 0
        player(index).PartyPlayer = 0
        player(index).InParty = 0
        player(index).target = 0
        player(index).TargetType = 0
        player(index).CastedSpell = NO
        player(index).PartyStarter = NO
        player(index).GettingMap = NO
    Next i
End Sub

Sub ClearChar(ByVal index As Long, ByVal CharNum As Long)
Dim n As Long
    
    player(index).Char(CharNum).Name = ""
    player(index).Char(CharNum).Class = 0
    player(index).Char(CharNum).sprite = 0
    player(index).Char(CharNum).level = 0
    player(index).Char(CharNum).Exp = 0
    player(index).Char(CharNum).Access = 0
    player(index).Char(CharNum).PK = NO
    player(index).Char(CharNum).POINTS = 0
    player(index).Char(CharNum).Guild = 0
    
    player(index).Char(CharNum).HP = 0
    player(index).Char(CharNum).MP = 0
    player(index).Char(CharNum).SP = 0
    
    player(index).Char(CharNum).str = 0
    player(index).Char(CharNum).intel = 0
    player(index).Char(CharNum).dex = 0
    player(index).Char(CharNum).con = 0
    player(index).Char(CharNum).wiz = 0
    player(index).Char(CharNum).cha = 0
    'Player(index).Char(CharNum).DEF = 0
    'Player(index).Char(CharNum).speed = 0
    'Player(index).Char(CharNum).MAGI = 0
    
    For n = 1 To MAX_INV
        player(index).Char(CharNum).Inv(n).num = 0
        player(index).Char(CharNum).Inv(n).value = 0
        player(index).Char(CharNum).Inv(n).Dur = 0
    Next n
    
    For n = 1 To MAX_PLAYER_SPELLS
        player(index).Char(CharNum).Spell(n) = 0
    Next n
    
    player(index).Char(CharNum).ArmorSlot = 0
    player(index).Char(CharNum).WeaponSlot = 0
    player(index).Char(CharNum).HelmetSlot = 0
    player(index).Char(CharNum).ShieldSlot = 0
    
    player(index).Char(CharNum).map = 0
    player(index).Char(CharNum).x = 0
    player(index).Char(CharNum).y = 0
    player(index).Char(CharNum).Dir = 0
End Sub
    
Sub ClearItem(ByVal index As Long)
    Item(index).Name = ""
    
    Item(index).type = 0
    Item(index).Data1 = 0
    Item(index).Data2 = 0
    Item(index).Data3 = 0
    Item(index).BaseDamage = 0
    Item(index).cha = 0
    Item(index).con = 0
    Item(index).dex = 0
    Item(index).intel = 0
    Item(index).str = 0
    Item(index).wiz = 0
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
    Npc(index).sprite = 0
    Npc(index).SpawnSecs = 0
    Npc(index).Behavior = 0
    Npc(index).Range = 0
    Npc(index).DropChance = 0
    Npc(index).DropItem = 0
    Npc(index).DropItemValue = 0
    Npc(index).str = 0
    Npc(index).def = 0
    Npc(index).speed = 0
    Npc(index).MAGI = 0
End Sub

Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next i
End Sub

Sub ClearMapItem(ByVal index As Long, ByVal mapNum As Long)
    MapItem(mapNum, index).num = 0
    MapItem(mapNum, index).value = 0
    MapItem(mapNum, index).Dur = 0
    MapItem(mapNum, index).x = 0
    MapItem(mapNum, index).y = 0
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

Sub ClearMapNpc(ByVal index As Long, ByVal mapNum As Long)
    MapNpc(mapNum, index).num = 0
    MapNpc(mapNum, index).target = 0
    MapNpc(mapNum, index).HP = 0
    MapNpc(mapNum, index).MP = 0
    MapNpc(mapNum, index).SP = 0
    MapNpc(mapNum, index).x = 0
    MapNpc(mapNum, index).y = 0
    MapNpc(mapNum, index).Dir = 0
    MapNpc(mapNum, index).maxHP = 0
    
    ' Server use only
    MapNpc(mapNum, index).SpawnWait = 0
    MapNpc(mapNum, index).AttackTimer = 0
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

Sub ClearMap(ByVal mapNum As Long)
Dim i As Long
Dim x As Long
Dim y As Long

    map(mapNum).Name = ""
    map(mapNum).Revision = 0
    map(mapNum).Moral = 0
    map(mapNum).Up = 0
    map(mapNum).Down = 0
    map(mapNum).Left = 0
    map(mapNum).Right = 0
        
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            map(mapNum).Tile(x, y).Ground = 0
            map(mapNum).Tile(x, y).Mask = 0
            map(mapNum).Tile(x, y).Anim = 0
            map(mapNum).Tile(x, y).Fringe = 0
            map(mapNum).Tile(x, y).type = 0
            map(mapNum).Tile(x, y).Data1 = 0
            map(mapNum).Tile(x, y).Data2 = 0
            map(mapNum).Tile(x, y).Data3 = 0
            map(mapNum).Tile(x, y).Data4 = 0
            map(mapNum).Tile(x, y).Data5 = 0
            'map(MapNum).Tile(x, y).Data6 = 0
            'map(MapNum).Tile(x, y).Data7 = 0
            'map(MapNum).Tile(x, y).Data8 = 0
            'map(MapNum).Tile(x, y).Data9 = 0
            'map(MapNum).Tile(x, y).Data10 = 0
        Next x
    Next y
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(mapNum) = NO
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
    Spell(index).type = 0
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
    GetPlayerName = Trim(player(index).Char(player(index).CharNum).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    player(index).Char(player(index).CharNum).Name = Name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = player(index).Char(player(index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    player(index).Char(player(index).CharNum).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    GetPlayerSprite = player(index).Char(player(index).CharNum).sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal sprite As Long)
    player(index).Char(player(index).CharNum).sprite = sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    GetPlayerLevel = player(index).Char(player(index).CharNum).level
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal level As Long)
    player(index).Char(player(index).CharNum).level = level
End Sub

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = 500 * (((GetPlayerLevel(index) ^ 3) - GetPlayerLevel(index)) + (25))
    'GetPlayerNextLevel = (GetPlayerLevel(index) + 1) * (GetPlayerSTR(index) + GetPlayerINT(index) + GetPlayerDEX(index) + GetPlayerCON(index) + GetPlayerWIZ(index) + GetPlayerCHA(index) + GetPlayerPOINTS(index)) * 25
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = player(index).Char(player(index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)
    player(index).Char(player(index).CharNum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    GetPlayerAccess = player(index).Char(player(index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    player(index).Char(player(index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    GetPlayerPK = player(index).Char(player(index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
If PK = YES Then
    If GetPlayerAccess(index) < 1 Then
        Call SetPlayerColour(index, 12, False)
        player(index).Char(player(index).CharNum).PK = PK
    End If
    
Else
    If GetPlayerAccess(index) < 1 Then
        Call SetPlayerColour(index, 15, False)
        player(index).Char(player(index).CharNum).PK = PK
    End If
End If
    
End Sub

Function GetPlayerHP(ByVal index As Long) As Long
    GetPlayerHP = player(index).Char(player(index).CharNum).HP
End Function

Sub SetPlayerHP(ByVal index As Long, ByVal HP As Long)
    player(index).Char(player(index).CharNum).HP = HP
    
    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then
        player(index).Char(player(index).CharNum).HP = GetPlayerMaxHP(index)
    End If
    If GetPlayerHP(index) < 0 Then
        player(index).Char(player(index).CharNum).HP = 0
    End If
End Sub

Sub SetPlayerPoisonLength(ByVal index As Long, ByVal Length As Long)
    player(index).Char(player(index).CharNum).Poison_length = Length
    If GetPlayerPoisonLength(index) < 0 Then
        player(index).Char(player(index).CharNum).Poison_length = 0
        player(index).Char(player(index).CharNum).Poison = False
        player(index).Char(player(index).CharNum).Poison_vital = 0
    End If
End Sub

Function SetPlayerBoosts(ByVal index As Long, ByRef Boosts As PointBoostRec) As Boolean
Dim freeBoost As Long
freeBoost = getFreeBoost(index)
If freeBoost <> 0 Then
    player(index).Char(player(index).CharNum).Boosts(freeBoost).cha = Boosts.cha
    player(index).Char(player(index).CharNum).Boosts(freeBoost).con = Boosts.con
    player(index).Char(player(index).CharNum).Boosts(freeBoost).dex = Boosts.dex
    player(index).Char(player(index).CharNum).Boosts(freeBoost).HP = Boosts.HP
    player(index).Char(player(index).CharNum).Boosts(freeBoost).intel = Boosts.intel
    player(index).Char(player(index).CharNum).Boosts(freeBoost).MP = Boosts.MP
    player(index).Char(player(index).CharNum).Boosts(freeBoost).PP = Boosts.PP
    player(index).Char(player(index).CharNum).Boosts(freeBoost).SP = Boosts.SP
    player(index).Char(player(index).CharNum).Boosts(freeBoost).str = Boosts.str
    player(index).Char(player(index).CharNum).Boosts(freeBoost).wiz = Boosts.wiz
    SetPlayerBoosts = True
Else
    SetPlayerBoosts = False
End If
End Function

Function getFreeBoost(ByVal index As Long) As Long
Dim i As Long
    For i = 1 To MAX_BOOSTS
        If player(index).Char(player(index).CharNum).Boosts(i).isUsed = False Then
            getFreeBoost = i
            Exit Function
        End If
    Next i
    getFreeBoost = 0
End Function

Function GetPlayerPoison(index As Long)
    GetPlayerPoison = player(index).Char(player(index).CharNum).Poison
End Function

Function GetPlayerPoisonLength(ByVal index As Long) As Long
    GetPlayerPoisonLength = player(index).Char(player(index).CharNum).Poison_length
End Function

Sub setPlayerPoison(ByVal index As Long, ByVal isPoisoned As Boolean, ByVal Length As Long, ByVal vital As Long)
    player(index).Char(player(index).CharNum).Poison = isPoisoned
    player(index).Char(player(index).CharNum).Poison_length = Length
    player(index).Char(player(index).CharNum).Poison_vital = vital
End Sub

Function GetPlayerMP(ByVal index As Long) As Long
    GetPlayerMP = player(index).Char(player(index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
    player(index).Char(player(index).CharNum).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        player(index).Char(player(index).CharNum).MP = GetPlayerMaxMP(index)
    End If
    If GetPlayerMP(index) < 0 Then
        player(index).Char(player(index).CharNum).MP = 0
    End If
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
    GetPlayerSP = player(index).Char(player(index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
    player(index).Char(player(index).CharNum).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        player(index).Char(player(index).CharNum).SP = GetPlayerMaxSP(index)
    End If
    If GetPlayerSP(index) < 0 Then
        player(index).Char(player(index).CharNum).SP = 0
    End If
End Sub

Function GetPlayerPP(ByVal index As Long) As Long
    GetPlayerPP = player(index).Char(player(index).CharNum).PP
End Function

Function GetPlayerMaxPP(ByVal index As Long) As Long
    Dim CharNum As Long

    CharNum = player(index).CharNum
    'GetPlayerMaxMP = (Player(index).Char(CharNum).level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).MAGI) * 2
    'newsystem
    GetPlayerMaxPP = (player(index).Char(CharNum).wiz + player(index).Char(CharNum).intel) * 0.7 * player(index).Char(CharNum).level + 1
End Function

Sub SetPlayerPP(ByVal index As Long, ByVal PP As Long)
    player(index).Char(player(index).CharNum).PP = PP

    If GetPlayerPP(index) > GetPlayerMaxPP(index) Then
        player(index).Char(player(index).CharNum).PP = GetPlayerMaxPP(index)
    End If
End Sub

Function GetPlayerMaxHP(ByVal index As Long) As Long
Dim CharNum As Long
Dim i As Long

    CharNum = player(index).CharNum
    'GetPlayerMaxHP = (Player(index).Char(CharNum).level + Int(GetPlayerSTR(index) / 2) + Class(Player(index).Char(CharNum).Class).STR) * 2
    'new system
    GetPlayerMaxHP = (player(index).Char(CharNum).con * 2) * player(index).Char(CharNum).level
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
Dim CharNum As Long

    CharNum = player(index).CharNum
    'GetPlayerMaxMP = (Player(index).Char(CharNum).level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).MAGI) * 2
    'newsystem
    GetPlayerMaxMP = (player(index).Char(CharNum).wiz + player(index).Char(CharNum).intel) * 2.5 * player(index).Char(CharNum).level
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long
Dim CharNum As Long

    CharNum = player(index).CharNum
    GetPlayerMaxSP = (player(index).Char(CharNum).level + Int(GetPlayerDEX(index) / 2) + Class(player(index).Char(CharNum).Class).dex) * 2
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim(Class(ClassNum).Name)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (Class(ClassNum).con * 2)
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (Class(ClassNum).wiz + Class(ClassNum).intel) * 2.5
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(Class(ClassNum).dex / 2) + Class(ClassNum).dex) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = Class(ClassNum).str
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = Class(ClassNum).def
End Function



Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = Class(ClassNum).speed
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = Class(ClassNum).MAGI
End Function

Function GetClassINT(ByVal ClassNum As Long) As Long
    GetClassINT = Class(ClassNum).intel
End Function
Function GetClassDex(ByVal ClassNum As Long) As Long
    GetClassDex = Class(ClassNum).dex
End Function
Function GetClassCon(ByVal ClassNum As Long) As Long
    GetClassCon = Class(ClassNum).con
End Function
Function GetClassWiz(ByVal ClassNum As Long) As Long
    GetClassWiz = Class(ClassNum).wiz
End Function
Function GetClassCha(ByVal ClassNum As Long) As Long
    GetClassCha = Class(ClassNum).cha
End Function

Function GetPlayerSTR(ByVal index As Long) As Long
    GetPlayerSTR = player(index).Char(player(index).CharNum).str
End Function

Sub SetPlayerSTR(ByVal index As Long, ByVal str As Long)
    player(index).Char(player(index).CharNum).str = str
End Sub

Function GetPlayerINT(ByVal index As Long) As Long
    GetPlayerINT = player(index).Char(player(index).CharNum).intel
End Function

Sub SetPlayerINT(ByVal index As Long, ByVal intel As Long)
    player(index).Char(player(index).CharNum).intel = intel
End Sub

Function GetPlayerDEX(ByVal index As Long) As Long
    GetPlayerDEX = player(index).Char(player(index).CharNum).dex
End Function

Sub SetPlayerDEX(ByVal index As Long, ByVal dex As Long)
    player(index).Char(player(index).CharNum).dex = dex
End Sub

Function GetPlayerCON(ByVal index As Long) As Long
    GetPlayerCON = player(index).Char(player(index).CharNum).con
End Function

Sub SetPlayerCON(ByVal index As Long, ByVal con As Long)
    player(index).Char(player(index).CharNum).con = con
End Sub

Function GetPlayerWIZ(ByVal index As Long) As Long
    GetPlayerWIZ = player(index).Char(player(index).CharNum).wiz
End Function

Sub SetPlayerWIZ(ByVal index As Long, ByVal wiz As Long)
    player(index).Char(player(index).CharNum).con = wiz
End Sub

Function GetPlayerCHA(ByVal index As Long) As Long
    GetPlayerCHA = player(index).Char(player(index).CharNum).cha
End Function

Sub SetPlayerCHA(ByVal index As Long, ByVal cha As Long)
    player(index).Char(player(index).CharNum).con = cha
End Sub

Function GetPlayerDEF(ByVal index As Long) As Long
    GetPlayerDEF = player(index).Char(player(index).CharNum).def
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal def As Long)
    player(index).Char(player(index).CharNum).def = def
End Sub

Function GetPlayerSPEED(ByVal index As Long) As Long
    GetPlayerSPEED = player(index).Char(player(index).CharNum).speed
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal speed As Long)
    player(index).Char(player(index).CharNum).speed = speed
End Sub

Function GetPlayerMAGI(ByVal index As Long) As Long
    GetPlayerMAGI = player(index).Char(player(index).CharNum).MAGI
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal MAGI As Long)
    player(index).Char(player(index).CharNum).MAGI = MAGI
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    GetPlayerPOINTS = player(index).Char(player(index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    player(index).Char(player(index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    GetPlayerMap = player(index).Char(player(index).CharNum).map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal mapNum As Long)
    If mapNum > 0 And mapNum <= MAX_MAPS Then
        player(index).Char(player(index).CharNum).map = mapNum
    End If
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    GetPlayerX = player(index).Char(player(index).CharNum).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    player(index).Char(player(index).CharNum).x = x
End Sub

Function GetPlayerColour(ByVal index As Long, ByVal forText As Boolean) As Long
    If forText Then
        GetPlayerColour = player(index).Char(player(index).CharNum).txtColour
    Else
        GetPlayerColour = player(index).Char(player(index).CharNum).ingameColour
    End If
End Function

Function GetPlayerY(ByVal index As Long) As Long
    GetPlayerY = player(index).Char(player(index).CharNum).y
End Function

Sub SetPlayerColour(ByVal index As Long, ByVal Colour As Long, ByVal forText As Boolean)
    If forText Then
        player(index).Char(player(index).CharNum).txtColour = Colour
    Else
        player(index).Char(player(index).CharNum).ingameColour = Colour
    End If
End Sub

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    player(index).Char(player(index).CharNum).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    GetPlayerDir = player(index).Char(player(index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    player(index).Char(player(index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = player(index).Char(player(index).CharNum).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal itemnum As Long)
    player(index).Char(player(index).CharNum).Inv(InvSlot).num = itemnum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = player(index).Char(player(index).CharNum).Inv(InvSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    player(index).Char(player(index).CharNum).Inv(InvSlot).value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = 0 'Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    player(index).Char(player(index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemNum = player(index).Char(player(index).CharNum).Bank(InvSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal itemnum As Long)
    player(index).Char(player(index).CharNum).Bank(InvSlot).num = itemnum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemValue = player(index).Char(player(index).CharNum).Bank(InvSlot).value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    player(index).Char(player(index).CharNum).Bank(InvSlot).value = ItemValue
End Sub

Function GetPlayerBankItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemDur = player(index).Char(player(index).CharNum).Bank(InvSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    player(index).Char(player(index).CharNum).Bank(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = player(index).Char(player(index).CharNum).Spell(SpellSlot)
End Function

Function GetPlayerPrayer(ByVal index As Long, ByVal PrayerSlot As Long) As Long
    GetPlayerPrayer = player(index).Char(player(index).CharNum).Prayer(PrayerSlot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    player(index).Char(player(index).CharNum).Spell(SpellSlot) = SpellNum
End Sub
Sub SetPlayerPrayer(ByVal index As Long, ByVal PrayerSlot As Long, ByVal PrayerNum As Long)
    player(index).Char(player(index).CharNum).Prayer(PrayerSlot) = PrayerNum
End Sub

Function GetPlayerArmorSlot(ByVal index As Long) As Long
    GetPlayerArmorSlot = player(index).Char(player(index).CharNum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)
    player(index).Char(player(index).CharNum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal index As Long) As Long
    GetPlayerWeaponSlot = player(index).Char(player(index).CharNum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)
    player(index).Char(player(index).CharNum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long
    GetPlayerHelmetSlot = player(index).Char(player(index).CharNum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)
    player(index).Char(player(index).CharNum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long
    GetPlayerShieldSlot = player(index).Char(player(index).CharNum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)
    player(index).Char(player(index).CharNum).ShieldSlot = InvNum
End Sub

