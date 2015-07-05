Attribute VB_Name = "modTypes"
Option Explicit

' Winsock globals
Public GAME_PORT

' General constants
Public Const GAME_NAME = "After Darkness"
Public Const MAX_PLAYERS = 30
Public Const MAX_ITEMS = 1020
Public Const MAX_NPCS = 255
Public Const MAX_INV = 32
Public Const MAX_BANK = 30
Public Const MAX_MAP_ITEMS = 20
Public Const MAX_MAP_NPCS = 14
Public Const MAX_SHOPS = 10
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_SPELLS = 255
Public Const MAX_TRADES = 20
Public Const MAX_SIGNS = 255

'Editor and main form sizes
Public Const FORM_X As Long = 12090
Public Const FORM_Y As Long = 9750
Public Const FORM_EDITOR_X As Long = 15000
Public Const FORM_EDITOR_Y As Long = 9480

'tile consts
Public Const MAX_TILE_SHEETS = 20
Public Const MAX_TILE_WIDTH = 13

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

' Spell constants
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_GIVEITEM = 6

' Prayer Consts
Public Const PRAYER_TYPE_HEAL = 0
Public Const PRAYER_TYPE_CURE = 1
Public Const PRAYER_TYPE_ENHANCE = 2

Type PlayerInvRec
    num As Byte
    value As Long
    Dur As Integer
End Type

Type SignRec
    header As String
    Msg As String
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Class As Byte
    sprite As Integer
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    PP As Long ' prayer points :(
    
    ' Stats
    str As Byte
    intel As Byte
    dex As Byte
    con As Byte
    wiz As Byte
    cha As Byte
    'old stuff
    DEF As Byte
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
    
    ' Client use only
    maxHP As Long
    MaxMP As Long
    MaxSP As Long
    MaxPP As Long
    XOffset As Integer
    YOffset As Integer
    moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
    colour As Long
    
    'effects
    Poison As Boolean
    Poison_length As Long
    Poison_vital As Long
End Type
    
Type TileRec
    Ground As Integer
    mask As Integer
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

Type MapRec
    Name As String * NAME_LENGTH
    street As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    music As Byte
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
    DEF As Byte
    speed As Byte
    MAGI As Byte
    intel As Byte
    dex As Byte
    con As Byte
    wiz As Byte
    cha As Byte
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
    
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
    AttackSay As String * 100
    
    sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    
    str  As Byte
    DEF As Byte
    speed As Byte
    MAGI As Byte
    HP As Long
    
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
    num As Byte
    
    target As Byte
    
    HP As Long
    MP As Long
    SP As Long
        
    map As Integer
    x As Byte
    y As Byte
    Dir As Byte
    
    maxHP As Long

    ' Client use only
    XOffset As Integer
    YOffset As Integer
    moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Respawn As Boolean
    Attack_with_Poison As Boolean
    Poison_length As Long
    Poison_vital As Long
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
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
    Sound As Long
    ManaUse As Long
End Type

Type PrayerRec
    Name As String * NAME_LENGTH
    ClassReq As Byte
    LevelReq As Byte
    type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    Sound As Long
    ManaUse As Long
End Type

Type TempTileRec
    DoorOpen As Byte
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

'alpha belnding
Public Type rBlendProps
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
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
'
'    XOffset As Integer
'    YOffset As Integer
'
'    Dir As Integer
'    moving As Byte
'
'    AttackTimer As Long
'
'    maxHP As Long
'End Type

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

Public map As MapRec
Public TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Prayer(1 To MAX_SPELLS) As PrayerRec
Public Signs(1 To MAX_SIGNS) As SignRec
Public Quests(1 To MAX_QUESTS) As QuestRec

'Public Pets(1 To MAX_PLAYERS) As PetRec

Public blnNight As Boolean
Public showLightning As Boolean

Public canMoveNow As Boolean

'for mini map
'Public tileHDCs(0 To MAX_TILE_SHEETS) As Long
'Public gotHDCs As Boolean



Sub ClearTempTile()
Dim x As Long, y As Long

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            TempTile(x, y).DoorOpen = NO
        Next x
    Next y
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim i As Long
Dim n As Long

    Player(Index).Name = ""
    Player(Index).Class = 0
    Player(Index).Level = 0
    Player(Index).sprite = 0
    Player(Index).Exp = 0
    Player(Index).Access = 0
    Player(Index).PK = NO
        
    Player(Index).HP = 0
    Player(Index).MP = 0
    Player(Index).SP = 0
        
    Player(Index).str = 0
    Player(Index).DEF = 0
    Player(Index).speed = 0
    Player(Index).MAGI = 0
    Player(Index).intel = 0
    Player(Index).dex = 0
    Player(Index).con = 0
    Player(Index).wiz = 0
    Player(Index).cha = 0
        
    For n = 1 To MAX_INV
        Player(Index).Inv(n).num = 0
        Player(Index).Inv(n).value = 0
        Player(Index).Inv(n).Dur = 0
    Next n
    For n = 1 To MAX_BANK
        Player(Index).Bank(n).num = 0
        Player(Index).Bank(n).value = 0
        Player(Index).Bank(n).Dur = 0
    Next n
        
    Player(Index).ArmorSlot = 0
    Player(Index).WeaponSlot = 0
    Player(Index).HelmetSlot = 0
    Player(Index).ShieldSlot = 0
        
    Player(Index).map = 0
    Player(Index).x = 0
    Player(Index).y = 0
    Player(Index).Dir = 0
    
    ' Client use only
    Player(Index).maxHP = 0
    Player(Index).MaxMP = 0
    Player(Index).MaxSP = 0
    Player(Index).XOffset = 0
    Player(Index).YOffset = 0
    Player(Index).moving = 0
    Player(Index).Attacking = 0
    Player(Index).AttackTimer = 0
    Player(Index).MapGetTimer = 0
    Player(Index).CastedSpell = NO
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index).Name = ""
    
    Item(Index).type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
    
    Item(Index).BaseDamage = 0
    Item(Index).cha = 0
    Item(Index).con = 0
    Item(Index).dex = 0
    Item(Index).intel = 0
    Item(Index).str = 0
    Item(Index).wiz = 0
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearMapItem(ByVal Index As Long)
    MapItem(Index).num = 0
    MapItem(Index).value = 0
    MapItem(Index).Dur = 0
    MapItem(Index).x = 0
    MapItem(Index).y = 0
End Sub

Sub ClearMap()
Dim i As Long
Dim x As Long
Dim y As Long

    map.Name = ""
    map.Revision = 0
    map.Moral = 0
    map.Up = 0
    map.Down = 0
    map.Left = 0
    map.Right = 0
        
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            map.Tile(x, y).Ground = 0
            map.Tile(x, y).mask = 0
            map.Tile(x, y).Anim = 0
            map.Tile(x, y).Fringe = 0
            map.Tile(x, y).type = 0
            map.Tile(x, y).Data1 = 0
            map.Tile(x, y).Data2 = 0
            map.Tile(x, y).Data3 = 0
            map.Tile(x, y).Data4 = 0
            map.Tile(x, y).Data5 = 0
            'Map.Tile(X, Y).Data6 = 0
            'Map.Tile(X, Y).Data7 = 0
            'Map.Tile(X, Y).Data8 = 0
            'Map.Tile(X, Y).Data9 = 0
            'Map.Tile(X, Y).Data10 = 0
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
    MapNpc(Index).target = 0
    MapNpc(Index).HP = 0
    MapNpc(Index).MP = 0
    MapNpc(Index).SP = 0
    MapNpc(Index).map = 0
    MapNpc(Index).x = 0
    MapNpc(Index).y = 0
    MapNpc(Index).Dir = 0
    
    ' Client use only
    MapNpc(Index).XOffset = 0
    MapNpc(Index).YOffset = 0
    MapNpc(Index).moving = 0
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
    GetPlayerName = Trim(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal sprite As Long)
    Player(Index).sprite = sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Exp = Exp
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
Sub SetPlayerMaxHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).maxHP = HP
End Sub
Sub SetPlayerMaxMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).MaxMP = MP
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

Function GetPlayerPP(ByVal Index As Long) As Long
    GetPlayerPP = Player(Index).PP
End Function

Function GetPlayerMaxPP(ByVal Index As Long) As Long
    GetPlayerMaxPP = Player(Index).MaxPP
End Function

Sub SetPlayerPP(ByVal Index As Long, ByVal PP As Long)
    Player(Index).PP = PP

    If GetPlayerPP(Index) > GetPlayerMaxPP(Index) Then
        Player(Index).PP = GetPlayerMaxPP(Index)
    End If
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    GetPlayerMaxHP = Player(Index).maxHP
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    GetPlayerMaxMP = Player(Index).MaxMP
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    GetPlayerMaxSP = Player(Index).MaxSP
End Function

Function GetPlayerSTR(ByVal Index As Long) As Long
    GetPlayerSTR = Player(Index).str
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal str As Long)
    Player(Index).str = str
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

Function GetPlayerINT(ByVal Index As Long) As Long
    GetPlayerINT = Player(Index).intel
End Function

Sub SetPlayerInt(ByVal Index As Long, ByVal intel As Long)
    Player(Index).intel = intel
End Sub

Function GetPlayerDex(ByVal Index As Long) As Long
    GetPlayerDex = Player(Index).dex
End Function

Sub SetPlayerDex(ByVal Index As Long, ByVal dex As Long)
    Player(Index).dex = dex
End Sub

Function GetPlayerCon(ByVal Index As Long) As Long
    GetPlayerCon = Player(Index).con
End Function

Sub SetPlayerCon(ByVal Index As Long, ByVal con As Long)
    Player(Index).con = con
End Sub

Function GetPlayerWiz(ByVal Index As Long) As Long
    GetPlayerWiz = Player(Index).wiz
End Function

Sub SetPlayerWiz(ByVal Index As Long, ByVal wiz As Long)
    Player(Index).wiz = wiz
End Sub

Function GetPlayerCha(ByVal Index As Long) As Long
    GetPlayerCha = Player(Index).cha
End Function

Sub SetPlayerCha(ByVal Index As Long, ByVal cha As Long)
    Player(Index).cha = cha
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).map = MapNum
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

Sub SetPlayerColour(ByVal Index As Long, ByVal colour As Long)
    Player(Index).colour = colour
End Sub
Function GetPlayerColour(ByVal Index As Long) As Long
    GetPlayerColour = Player(Index).colour
End Function

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
    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(InvSlot).value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    If InvSlot > 0 And InvSlot < MAX_BANK Then
        GetPlayerBankItemNum = Player(Index).Bank(InvSlot).num
    End If
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Bank(InvSlot).num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Bank(InvSlot).value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Bank(InvSlot).value = ItemValue
End Sub

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemDur = Player(Index).Bank(InvSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Bank(InvSlot).Dur = ItemDur
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

