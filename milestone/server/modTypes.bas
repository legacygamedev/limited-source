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
Public Scripting As Long

Public Const MAX_INV = 24
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 8
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 = "jwehiehfojcvnvnsdinaoiwheoewyriusdyrflsdjncjkxzncisdughfusyfuapsipiuahfpaijnflkjnvjnuahguiryasbdlfkjblsahgfauygewuifaunfauf"
Public Const SEC_CODE2 = "ksisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlasndowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas"
Public Const SEC_CODE3 = "saiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi"
Public Const SEC_CODE4 = "88978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672"

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

' Target type constants
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1

Type PlayerInvRec
    Num As Byte
    value As Long
    Dur As Integer
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Sex As Byte
    Class As Byte
    Sprite As Integer
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
    SPEED As Long
    MAGI As Long
    POINTS As Long
    
    ' Worn equipment
    ArmorSlot As Byte
    WeaponSlot As Byte
    HelmetSlot As Byte
    ShieldSlot As Byte
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
    
    ' Position
    Map As Integer
    x As Byte
    y As Byte
    Dir As Byte
End Type

Type PlayerTradeRec
    InvNum As Byte
    InvName As String
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
    Target As Byte
    CastedSpell As Byte
    PartyStarter As Byte
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
    Shop As Byte
    Indoors As Byte
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    
    AdvanceFrom As Long
    LevelReq As Long
    Type As Long
    Locked As Long
    
    MaleSprite As Integer
    FemaleSprite As Integer
    
    STR As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte
    
    Map As Long
    x As Byte
    y As Byte
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String
    
    Pic As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    StrReq As Integer
    DefReq As Integer
    SpeedReq As Integer
    ClassReq As Integer
    AccessReq As Byte
    
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
    Num As Byte
    value As Long
    Dur As Integer
    
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
    AttackSay As String * 255
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    STR  As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte
    Big As Byte
    MaxHp As Long
    Exp As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
End Type

Type MapNpcRec
    Num As Integer
    
    Target As Integer
    
    HP As Long
    MP As Long
    SP As Long
        
    x As Byte
    y As Byte
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
    LevelReq As Integer
    MPCost As Integer
    Sound As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
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

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

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
        Class(i).AdvanceFrom = 0
        Class(i).LevelReq = 0
        Class(i).Type = 1
        Class(i).STR = 0
        Class(i).DEF = 0
        Class(i).SPEED = 0
        Class(i).MAGI = 0
        Class(i).FemaleSprite = 0
        Class(i).MaleSprite = 0
        Class(i).Map = 0
        Class(i).x = 0
        Class(i).y = 0
    Next i
End Sub

Sub ClearClasses2()
Dim i As Long

    For i = 0 To Max_Classes
        Class2(i).Name = ""
        Class2(i).AdvanceFrom = 0
        Class2(i).LevelReq = 0
        Class2(i).Type = 2
        Class2(i).STR = 0
        Class2(i).DEF = 0
        Class2(i).SPEED = 0
        Class2(i).MAGI = 0
        Class2(i).FemaleSprite = 0
        Class2(i).MaleSprite = 0
        Class2(i).Map = 0
        Class2(i).x = 0
        Class2(i).y = 0
    Next i
End Sub

Sub ClearClasses3()
Dim i As Long

    For i = 0 To Max_Classes
        Class3(i).Name = ""
        Class3(i).AdvanceFrom = 0
        Class3(i).LevelReq = 0
        Class3(i).Type = 3
        Class3(i).STR = 0
        Class3(i).DEF = 0
        Class3(i).SPEED = 0
        Class3(i).MAGI = 0
        Class3(i).FemaleSprite = 0
        Class3(i).MaleSprite = 0
        Class3(i).Map = 0
        Class3(i).x = 0
        Class3(i).y = 0
    Next i
End Sub

Sub ClearPlayer(ByVal index As Long)
Dim i As Long
Dim n As Long

    Player(index).Login = ""
    Player(index).Password = ""
    
    For i = 1 To MAX_CHARS
        Player(index).Char(i).Name = ""
        Player(index).Char(i).Class = 0
        Player(index).Char(i).Level = 0
        Player(index).Char(i).Sprite = 0
        Player(index).Char(i).Exp = 0
        Player(index).Char(i).Access = 0
        Player(index).Char(i).PK = NO
        Player(index).Char(i).POINTS = 0
        Player(index).Char(i).Guild = ""
        
        Player(index).Char(i).HP = 0
        Player(index).Char(i).MP = 0
        Player(index).Char(i).SP = 0
        
        Player(index).Char(i).STR = 0
        Player(index).Char(i).DEF = 0
        Player(index).Char(i).SPEED = 0
        Player(index).Char(i).MAGI = 0
        
        For n = 1 To MAX_INV
            Player(index).Char(i).Inv(n).Num = 0
            Player(index).Char(i).Inv(n).value = 0
            Player(index).Char(i).Inv(n).Dur = 0
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            Player(index).Char(i).Spell(n) = 0
        Next n
        
        Player(index).Char(i).ArmorSlot = 0
        Player(index).Char(i).WeaponSlot = 0
        Player(index).Char(i).HelmetSlot = 0
        Player(index).Char(i).ShieldSlot = 0
        
        Player(index).Char(i).Map = 0
        Player(index).Char(i).x = 0
        Player(index).Char(i).y = 0
        Player(index).Char(i).Dir = 0
        
        ' Temporary vars
        Player(index).Buffer = ""
        Player(index).IncBuffer = ""
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
            Player(index).Trading(n).InvName = ""
            Player(index).Trading(n).InvNum = 0
        Next n
        Player(index).InTrade = 0
        Player(index).ChatPlayer = 0
    Next i
End Sub

Sub ClearChar(ByVal index As Long, ByVal CharNum As Long)
Dim n As Long
    
    Player(index).Char(CharNum).Name = ""
    Player(index).Char(CharNum).Class = 0
    Player(index).Char(CharNum).Sprite = 0
    Player(index).Char(CharNum).Level = 0
    Player(index).Char(CharNum).Exp = 0
    Player(index).Char(CharNum).Access = 0
    Player(index).Char(CharNum).PK = NO
    Player(index).Char(CharNum).POINTS = 0
    Player(index).Char(CharNum).Guild = ""
    
    Player(index).Char(CharNum).HP = 0
    Player(index).Char(CharNum).MP = 0
    Player(index).Char(CharNum).SP = 0
    
    Player(index).Char(CharNum).STR = 0
    Player(index).Char(CharNum).DEF = 0
    Player(index).Char(CharNum).SPEED = 0
    Player(index).Char(CharNum).MAGI = 0
    
    For n = 1 To MAX_INV
        Player(index).Char(CharNum).Inv(n).Num = 0
        Player(index).Char(CharNum).Inv(n).value = 0
        Player(index).Char(CharNum).Inv(n).Dur = 0
    Next n
    
    For n = 1 To MAX_PLAYER_SPELLS
        Player(index).Char(CharNum).Spell(n) = 0
    Next n
    
    Player(index).Char(CharNum).ArmorSlot = 0
    Player(index).Char(CharNum).WeaponSlot = 0
    Player(index).Char(CharNum).HelmetSlot = 0
    Player(index).Char(CharNum).ShieldSlot = 0
    
    Player(index).Char(CharNum).Map = 0
    Player(index).Char(CharNum).x = 0
    Player(index).Char(CharNum).y = 0
    Player(index).Char(CharNum).Dir = 0
End Sub
    
Sub ClearItem(ByVal index As Long)
    Item(index).Name = ""
    Item(index).Desc = ""
    
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
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearNpc(ByVal index As Long)
Dim i As Long
    Npc(index).Name = ""
    Npc(index).AttackSay = ""
    Npc(index).Sprite = 0
    Npc(index).SpawnSecs = 0
    Npc(index).Behavior = 0
    Npc(index).Range = 0
    Npc(index).STR = 0
    Npc(index).DEF = 0
    Npc(index).SPEED = 0
    Npc(index).MAGI = 0
    Npc(index).Big = 0
    Npc(index).MaxHp = 0
    Npc(index).Exp = 0
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

Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Long)
    MapItem(MapNum, index).Num = 0
    MapItem(MapNum, index).value = 0
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
    MapNpc(MapNum, index).Num = 0
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

Sub ClearMap(ByVal MapNum As Long)
Dim i As Long
Dim x As Long
Dim y As Long

    Map(MapNum).Name = ""
    Map(MapNum).Revision = 0
    Map(MapNum).Moral = 0
    Map(MapNum).Up = 0
    Map(MapNum).Down = 0
    Map(MapNum).Left = 0
    Map(MapNum).Right = 0
        
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
        Next x
    Next y
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
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
    Spell(index).Type = 0
    Spell(index).Data1 = 0
    Spell(index).Data2 = 0
    Spell(index).Data3 = 0
    Spell(index).MPCost = 0
    Spell(index).Sound = 0
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
    GetPlayerLogin = Trim(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    GetPlayerPassword = Trim(Player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    Player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String
    GetPlayerName = Trim(Player(index).Char(Player(index).CharNum).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Char(Player(index).CharNum).Name = Name
End Sub

Function GetPlayerGuild(ByVal index As Long) As String
    GetPlayerGuild = Trim(Player(index).Char(Player(index).CharNum).Guild)
End Function

Sub SetPlayerGuild(ByVal index As Long, ByVal Guild As String)
    Player(index).Char(Player(index).CharNum).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal index As Long) As Long
    GetPlayerGuildAccess = Player(index).Char(Player(index).CharNum).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal Guildaccess As Long)
    Player(index).Char(Player(index).CharNum).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Char(Player(index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Char(Player(index).CharNum).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    GetPlayerSprite = Player(index).Char(Player(index).CharNum).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Char(Player(index).CharNum).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    GetPlayerLevel = Player(index).Char(Player(index).CharNum).Level
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
    Player(index).Char(Player(index).CharNum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = Experience(GetPlayerLevel(index))
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).Char(Player(index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)
    Player(index).Char(Player(index).CharNum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    GetPlayerAccess = Player(index).Char(Player(index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Char(Player(index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    GetPlayerPK = Player(index).Char(Player(index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).Char(Player(index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal index As Long) As Long
    GetPlayerHP = Player(index).Char(Player(index).CharNum).HP
End Function

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

Function GetPlayerMP(ByVal index As Long) As Long
    GetPlayerMP = Player(index).Char(Player(index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
    Player(index).Char(Player(index).CharNum).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        Player(index).Char(Player(index).CharNum).MP = GetPlayerMaxMP(index)
    End If
    If GetPlayerMP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).MP = 0
    End If
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
    GetPlayerSP = Player(index).Char(Player(index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
    Player(index).Char(Player(index).CharNum).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        Player(index).Char(Player(index).CharNum).SP = GetPlayerMaxSP(index)
    End If
    If GetPlayerSP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).SP = 0
    End If
End Sub

Function GetPlayerMaxHP(ByVal index As Long) As Long
Dim CharNum As Long
Dim i As Long
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

    CharNum = Player(index).CharNum
    GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSTR(index) / 2) + Class(Player(index).Char(CharNum).Class).STR) * 2) + add
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

    CharNum = Player(index).CharNum
    GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).MAGI) * 2) + add
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

    CharNum = Player(index).CharNum
    GetPlayerMaxSP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSPEED(index) / 2) + Class(Player(index).Char(CharNum).Class).SPEED) * 2) + add
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim(Class(ClassNum).Name)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(Class(ClassNum).STR / 2) + Class(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(Class(ClassNum).MAGI / 2) + Class(ClassNum).MAGI) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(Class(ClassNum).SPEED / 2) + Class(ClassNum).SPEED) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = Class(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = Class(ClassNum).DEF
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = Class(ClassNum).SPEED
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = Class(ClassNum).MAGI
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
    GetPlayerSTR = Player(index).Char(Player(index).CharNum).STR + add
End Function

Sub SetPlayerSTR(ByVal index As Long, ByVal STR As Long)
    Player(index).Char(Player(index).CharNum).STR = STR
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
    GetPlayerDEF = Player(index).Char(Player(index).CharNum).DEF + add
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal DEF As Long)
    Player(index).Char(Player(index).CharNum).DEF = DEF
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
    GetPlayerSPEED = Player(index).Char(Player(index).CharNum).SPEED + add
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal SPEED As Long)
    Player(index).Char(Player(index).CharNum).SPEED = SPEED
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
    GetPlayerMAGI = Player(index).Char(Player(index).CharNum).MAGI + add
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal MAGI As Long)
    Player(index).Char(Player(index).CharNum).MAGI = MAGI
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    GetPlayerPOINTS = Player(index).Char(Player(index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    Player(index).Char(Player(index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    GetPlayerMap = Player(index).Char(Player(index).CharNum).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(index).Char(Player(index).CharNum).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    GetPlayerX = Player(index).Char(Player(index).CharNum).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    Player(index).Char(Player(index).CharNum).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    GetPlayerY = Player(index).Char(Player(index).CharNum).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    Player(index).Char(Player(index).CharNum).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    GetPlayerDir = Player(index).Char(Player(index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Char(Player(index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(index).Char(Player(index).CharNum).Inv(InvSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(index).Char(Player(index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(index).Char(Player(index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerArmorSlot(ByVal index As Long) As Long
    GetPlayerArmorSlot = Player(index).Char(Player(index).CharNum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal index As Long) As Long
    GetPlayerWeaponSlot = Player(index).Char(Player(index).CharNum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long
    GetPlayerHelmetSlot = Player(index).Char(Player(index).CharNum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long
    GetPlayerShieldSlot = Player(index).Char(Player(index).CharNum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).ShieldSlot = InvNum
End Sub

