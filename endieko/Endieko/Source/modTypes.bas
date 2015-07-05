Attribute VB_Name = "modTypes"
Option Explicit

' Winsock globals
Public Const GAME_PORT = 4000

' General constants
Public Const GAME_NAME = "Endieko Online"
Public Const WEBSITE = "http://www.endieko.com"
Public Const MAX_MAP_ITEMS = 20
Public Const MAX_ARROWS = 100
Public Const MAX_PLAYERS = 500
Public Const MAX_SPELLS = 1000
Public Const MAX_MAPS = 1000
Public Const MAX_SHOPS = 1000
Public Const MAX_ITEMS = 1000
Public Const MAX_NPCS = 1000
Public Const MAX_EMOTICONS = 100
Public Const MAX_BANK = 20
Public Const MAX_EFFECTS = 100

' Constants
Public Const MAX_INV = 24
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 12
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10
Public Const MAX_PLAYER_ARROWS = 100

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 1

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 = "jwehiehfojcvnvnsdinaoiwheoewyriusdyrflsdjncjkxzncisdughfusyfuapsipiuahfpaijnflkjnvjnuahguiryasbdlfkjblsahgfauygewuifaunfauf"
Public Const SEC_CODE2 = "ksisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlasndowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas"
Public Const SEC_CODE3 = "saiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi"
Public Const SEC_CODE4 = "88978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672"

' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map constants
Public Const MAX_MAPX = 30
Public Const MAX_MAPY = 30
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
Public Const TILE_TYPE_BANK = 20

' Item constants
Public Const ITEM_TYPE_NONE = 0
Public Const ITEM_TYPE_WEAPON = 1
Public Const ITEM_TYPE_ARMOR = 2
Public Const ITEM_TYPE_HELMET = 3
Public Const ITEM_TYPE_SHIELD = 4
Public Const ITEM_TYPE_LEGS = 5
Public Const ITEM_TYPE_BOOTS = 6
Public Const ITEM_TYPE_POTIONADDHP = 7
Public Const ITEM_TYPE_POTIONADDMP = 8
Public Const ITEM_TYPE_POTIONADDSP = 9
Public Const ITEM_TYPE_POTIONSUBHP = 10
Public Const ITEM_TYPE_POTIONSUBMP = 11
Public Const ITEM_TYPE_POTIONSUBSP = 12
Public Const ITEM_TYPE_KEY = 13
Public Const ITEM_TYPE_CURRENCY = 14
Public Const ITEM_TYPE_ARROW = 15
Public Const ITEM_TYPE_SPELL = 16

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
Public Const SPELL_TYPE_FREEZE = 7
Public Const SPELL_TYPE_CONTINUOUS = 8
Public Const SPELL_TYPE_WARP = 9
Public Const SPELL_TYPE_TRANSFORM = 10

' Effect Constats
Public Const EFFECT_TYPE_DRAIN = 0
Public Const EFFECT_TYPE_FORTIFY = 1
Public Const EFFECT_TYPE_FREEZE = 2

'Loc of pointer
Public CurX As Integer
Public CurY As Integer

Type EffectRec
    Name As String * NAME_LENGTH
    Effect As Byte
    Time As Byte
    Data1 As Byte
    Data2 As Byte
    Data3 As Byte
End Type

Type OverHeadRec
    Color As Byte
    Msg As String
    Time As Long
    ii As Long
End Type

Type ScriptRec
    ScriptNum As Byte
    Text As String
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

Type ArrowRec
    Name As String
    Pic As Long
    Range As Byte
    HasAmmo As Byte
    Ammunition As Integer
End Type

Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
    HasAmmo As Byte
    Ammunition As Integer
End Type

Type BankInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Type ItemTradeRec
    ItemGetNum As Long
    ItemGiveNum As Long
    ItemGetVal As Long
    ItemGiveVal As Long
End Type

Type TradeRec
    ItemS(1 To MAX_TRADES) As ItemTradeRec
    Selected As Long
    SelectedItem As Long
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Class As Byte
    Sprite As Integer
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Byte
    DEF As Byte
    Speed As Byte
    MAGI As Byte
    POINTS As Byte
    
    ' Worn equipment
    ArmorSlot As Byte
    WeaponSlot As Byte
    HelmetSlot As Byte
    ShieldSlot As Byte
    LegSlot As Byte
    BootSlot As Byte
    
    ' New Stuff
    Alignment As Integer
    FishingLevel As Byte
    FishingExp As Integer
    MiningLevel As Byte
    MiningExp As Integer
    LumberLevel As Byte
    LumberExp As Integer
    
    Status As Byte
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Bank(1 To MAX_BANK) As BankInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
    Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec
       
    ' Position
    Map As Integer
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
    Emoticon As Long
    EmoticonT As Long
End Type
    
Type TileRec
    Ground As Integer
    Mask As Integer
    Anim As Integer
    Mask2 As Integer
    M2Anim As Integer
    Fringe As Integer
    FAnim As Integer
    Fringe2 As Integer
    F2Anim As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    String1 As String
    String2 As String
    String3 As String
End Type

Type MapRec
    Name As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    Music As String
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    MaleSprite As Integer
    FemaleSprite As Integer
    
    Locked As Long
    
    STR As Byte
    DEF As Byte
    Speed As Byte
    MAGI As Byte
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    desc As String
    
    Pic As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    StrReq As Integer
    DefReq As Integer
    SpeedReq As Integer
    MagiReq As Integer
    ClassReq As Integer
    AccessReq As Byte
    DropOnDeath As Byte
    CannotBeRepaired As Byte
    
    AddHP As Long
    AddMP As Long
    AddSP As Long
    AddStr As Long
    AddDef As Long
    AddMagi As Long
    AddSpeed As Long
    AddEXP As Long
End Type

Type MapItemRec
    Num As Integer
    Value As Long
    Dur As Integer
    
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
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
        
    STR  As Byte
    DEF As Byte
    Speed As Byte
    MAGI As Byte
    Big As Byte
    MaxHp As Long
    EXP As Long
    Alignment As Integer
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
End Type

Type MapNpcRec
    Num As Byte
    
    Target As Byte
    
    HP As Long
    MaxHp As Long
    MP As Long
    SP As Long
    
    Map As Integer
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

Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
    FixesItems As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Type SpellRec
    Name As String * NAME_LENGTH
    Pic As Integer
    ClassReq As Byte
    LevelReq As Integer
    Sound As Integer
    MPCost As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Type TempTileRec
    DoorOpen As Byte
End Type

Type PlayerTradeRec
    InvNum As Byte
    InvName As String
End Type

Type EmoRec
    Pic As Long
    Command As String
End Type

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte


'ReDim SaveMap.Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
'ReDim Map.Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
'ReDim TempTile() As TempTileRec

Public Map As MapRec
Public TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public Trading2(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public Emoticons(0 To MAX_EMOTICONS) As EmoRec
Public MapReport(1 To MAX_MAPS) As MapRec
Public Arrows(1 To MAX_ARROWS) As ArrowRec
Public SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public Trade(1 To 6) As TradeRec
Public Script(1 To 7) As ScriptRec
Public Overhead As OverHeadRec
Public Effect(1 To MAX_EFFECTS) As EffectRec

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
    
    For n = 1 To MAX_BANK
        Player(Index).Bank(n).Num = 0
        Player(Index).Bank(n).Value = 0
        Player(Index).Bank(n).Dur = 0
    Next n
        
    Player(Index).ArmorSlot = 0
    Player(Index).WeaponSlot = 0
    Player(Index).HelmetSlot = 0
    Player(Index).ShieldSlot = 0
    
    Player(Index).Alignment = 0
    Player(Index).FishingLevel = 0
    Player(Index).FishingExp = 0
    Player(Index).MiningLevel = 0
    Player(Index).MiningExp = 0
    Player(Index).LumberLevel = 0
    Player(Index).LumberExp = 0
        
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
    Player(Index).Emoticon = -1
    Player(Index).EmoticonT = 0
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

    Map.Name = ""
    Map.Revision = 0
    Map.Moral = 0
    Map.Up = 0
    Map.Down = 0
    Map.Left = 0
    Map.Right = 0
        
    For Y = 0 To MAX_MAPY
    For X = 0 To MAX_MAPX
    Map.Tile(X, Y).Ground = 0
    Map.Tile(X, Y).Mask = 0
    Map.Tile(X, Y).Anim = 0
    Map.Tile(X, Y).Mask2 = 0
    Map.Tile(X, Y).M2Anim = 0
    Map.Tile(X, Y).Fringe = 0
    Map.Tile(X, Y).FAnim = 0
    Map.Tile(X, Y).Fringe2 = 0
    Map.Tile(X, Y).F2Anim = 0
    Map.Tile(X, Y).Type = 0
    Map.Tile(X, Y).Data1 = 0
    Map.Tile(X, Y).Data2 = 0
    Map.Tile(X, Y).Data3 = 0
    Map.Tile(X, Y).String1 = ""
    Map.Tile(X, Y).String2 = ""
    Map.Tile(X, Y).String3 = ""
    Next X
    Next Y
End Sub

Sub ClearMapItems()
Dim X As Long

    For X = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(X)
    Next X
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
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Guild)
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
    If InvSlot > MAX_INV Then Exit Function
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

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankInvSlot As Long) As Long
    GetPlayerBankItemNum = Player(Index).Bank(BankInvSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankInvSlot As Long, ByVal BankItemNum As Long)
    Player(Index).Bank(BankInvSlot).Num = BankItemNum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankInvSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Bank(BankInvSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankInvSlot As Long, ByVal BankItemValue As Long)
    Player(Index).Bank(BankInvSlot).Value = BankItemValue
End Sub

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal BankInvSlot As Long) As Long
    GetPlayerBankItemDur = Player(Index).Bank(BankInvSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal BankInvSlot As Long, ByVal BankItemDur As Long)
    Player(Index).Bank(BankInvSlot).Dur = BankItemDur
End Sub

Function GetPlayerBootSlot(ByVal Index As Long) As Long
    GetPlayerBootSlot = Player(Index).BootSlot
End Function

Sub SetPlayerBootSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).BootSlot = InvNum
End Sub

Function GetPlayerLegSlot(ByVal Index As Long) As Long
    GetPlayerLegSlot = Player(Index).LegSlot
End Function

Sub SetPlayerLegSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).LegSlot = InvNum
End Sub
