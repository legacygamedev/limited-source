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
Public Const MAX_BANK = 50
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 8
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10
Public Const MAX_PARTY_MEMS = 3 ' This is the number of players that can be in a party *NOT*

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 = "d32d223d23d323243r3453534fdsfds"
Public Const SEC_CODE2 = "fsdffsdf43fwdsdfsdfs34r34434fddfds"

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
Public Const MAP_MORAL_TRAINING = 3

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
Public Const TILE_TYPE_SEX_BLOCK = 20
Public Const TILE_TYPE_LEVEL_BLOCK = 21
Public Const TILE_TYPE_BANK = 22

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
Public Const ITEM_TYPE_SCROLL = 14
Public Const ITEM_TYPE_ORB = 15
Public Const ITEM_TYPE_BOOTS = 16
Public Const ITEM_TYPE_GLOVES = 17
Public Const ITEM_TYPE_RING = 18
Public Const ITEM_TYPE_AMULET = 19
Public Const ITEM_TYPE_BORB = 20
Public Const ITEM_TYPE_GGORB = 21

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
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_GIVEITEM = 6

' Target type constants
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1

Type BankRec
    Num As Long
    Value As Long
    Dur As Long
End Type

'Party info type
Type PartyInfo
    InParty As Byte
    Started As Byte
    PlayerNums(1 To MAX_PARTY_MEMS) As Long
End Type

'Binding orb position type
Type BindingPosition
     Map As Integer
     x As Byte
     y As Byte
End Type

Type PlayerInvRec
    Num As Byte
    Value As Long
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
    VIT As Long
    POINTS As Long
    
    ' Worn equipment
    ArmorSlot As Byte
    WeaponSlot As Byte
    HelmetSlot As Byte
    ShieldSlot As Byte
    BootsSlot As Byte
    GlovesSlot As Byte
    RingSlot As Byte
    AmuletSlot As Byte
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
    Bank(1 To MAX_BANK) As BankRec
    
    ' Position
    Map As Integer
    x As Byte
    y As Byte
    Dir As Byte
    
    ' Binding orb
    Binding  As BindingPosition
    
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
    TargetType As Byte
    Target As Byte
    CastedSpell As Byte
    GettingMap As Byte
    Emoticon As Long
    
    HardDrive As Long
    
    Reply As Long
    
    GuildInvitation As Boolean
    GuildTemp As String
    GuildInviter As Long
    
    Mute As Boolean

    InTrade As Byte
    TradePlayer As Long
    TradeOk As Byte
    TradeItemMax As Byte
    TradeItemMax2 As Byte
    Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
    
    InChat As Byte
    ChatPlayer As Long
    
    ' Party information
    Party As PartyInfo
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
    
    STR As Long
    DEF As Long
    SPEED As Long
    MAGI As Long
    VIT As Long
    
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
    Value As Long
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
On Error GoTo ErrorHandler
Dim i As Long, y As Long, x As Long

    For i = 1 To MAX_MAPS
        TempTile(i).DoorTimer = 0
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                TempTile(i).DoorOpen(x, y) = NO
            Next x
        Next y
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearTempTile", Err.Number, Err.Description
End Sub

Sub ClearClasses()
On Error GoTo ErrorHandler
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
        Class(i).VIT = 0
        Class(i).FemaleSprite = 0
        Class(i).MaleSprite = 0
        Class(i).Map = 0
        Class(i).x = 0
        Class(i).y = 0
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearClasses", Err.Number, Err.Description
End Sub

Sub ClearPlayer(ByVal index As Long)
On Error GoTo ErrorHandler
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
        Player(index).Char(i).VIT = 0
        
        For n = 1 To MAX_INV
            Player(index).Char(i).Inv(n).Num = 0
            Player(index).Char(i).Inv(n).Value = 0
            Player(index).Char(i).Inv(n).Dur = 0
        Next n
        
        For n = 1 To MAX_BANK
            Player(index).Char(i).Bank(n).Num = 0
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
        Player(index).Char(i).BootsSlot = 0
        Player(index).Char(i).GlovesSlot = 0
        Player(index).Char(i).RingSlot = 0
        Player(index).Char(i).AmuletSlot = 0
        
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
        Player(index).Target = 0
        Player(index).TargetType = 0
        Player(index).CastedSpell = NO
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
        Player(index).Mute = False
        Player(index).GuildInvitation = False
        Player(index).GuildTemp = ""
        Player(index).GuildInviter = 0
        For n = 1 To MAX_PARTY_MEMS
            Player(index).Party.PlayerNums(n) = 0
        Next n
        Player(index).Party.InParty = 0
        Player(index).Party.Started = NO
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearPlayer", Err.Number, Err.Description
End Sub

Sub ClearChar(ByVal index As Long, ByVal CharNum As Long)
On Error GoTo ErrorHandler
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
    Player(index).Char(CharNum).VIT = 0
    
    For n = 1 To MAX_INV
        Player(index).Char(CharNum).Inv(n).Num = 0
        Player(index).Char(CharNum).Inv(n).Value = 0
        Player(index).Char(CharNum).Inv(n).Dur = 0
    Next n
    
    For n = 1 To MAX_BANK
        Player(index).Char(CharNum).Bank(n).Num = 0
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
    Player(index).Char(CharNum).BootsSlot = 0
    Player(index).Char(CharNum).GlovesSlot = 0
    Player(index).Char(CharNum).RingSlot = 0
    Player(index).Char(CharNum).AmuletSlot = 0
    
    Player(index).Char(CharNum).Map = 0
    Player(index).Char(CharNum).x = 0
    Player(index).Char(CharNum).y = 0
    Player(index).Char(CharNum).Dir = 0
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearChar", Err.Number, Err.Description
End Sub

Sub ClearItem(ByVal index As Long)
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearItem", Err.Number, Err.Description
End Sub

Sub ClearItems()
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearItems", Err.Number, Err.Description
End Sub

Sub ClearNpc(ByVal index As Long)
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearNpc", Err.Number, Err.Description
End Sub

Sub ClearNpcs()
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearNpcs", Err.Number, Err.Description
End Sub

Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Long)
On Error GoTo ErrorHandler
    MapItem(MapNum, index).Num = 0
    MapItem(MapNum, index).Value = 0
    MapItem(MapNum, index).Dur = 0
    MapItem(MapNum, index).x = 0
    MapItem(MapNum, index).y = 0
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearMapItem", Err.Number, Err.Description
End Sub

Sub ClearMapItems()
On Error GoTo ErrorHandler
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next x
    Next y
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearMapItems", Err.Number, Err.Description
End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal MapNum As Long)
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearMapNpc", Err.Number, Err.Description
End Sub

Sub ClearMapNpcs()
On Error GoTo ErrorHandler
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next x
    Next y
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearMapNpcs", Err.Number, Err.Description
End Sub

Sub ClearMap(ByVal MapNum As Long)
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearMap", Err.Number, Err.Description
End Sub

Sub ClearMaps()
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearMaps", Err.Number, Err.Description
End Sub

Sub ClearShop(ByVal index As Long)
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearShop", Err.Number, Err.Description
End Sub

Sub ClearShops()
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearShops", Err.Number, Err.Description
End Sub

Sub ClearSpell(ByVal index As Long)
On Error GoTo ErrorHandler
    Spell(index).Name = ""
    Spell(index).ClassReq = 0
    Spell(index).LevelReq = 0
    Spell(index).Type = 0
    Spell(index).Data1 = 0
    Spell(index).Data2 = 0
    Spell(index).Data3 = 0
    Spell(index).MPCost = 0
    Spell(index).Sound = 0
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearSpell", Err.Number, Err.Description
End Sub

Sub ClearSpells()
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "ClearSpells", Err.Number, Err.Description
End Sub




' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Function GetPlayerLogin(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerLogin = Trim(Player(index).Login)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerLogin", Err.Number, Err.Description
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
On Error GoTo ErrorHandler
    Player(index).Login = Login
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerLogin", Err.Number, Err.Description
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerPassword = Trim(Player(index).Password)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerPassword", Err.Number, Err.Description
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
On Error GoTo ErrorHandler
    Player(index).Password = Password
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerPassword", Err.Number, Err.Description
End Sub

Function GetPlayerName(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerName = Trim(Player(index).Char(Player(index).CharNum).Name)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerName", Err.Number, Err.Description
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Name = Name
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerName", Err.Number, Err.Description
End Sub

Function GetPlayerGuild(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerGuild = Trim(Player(index).Char(Player(index).CharNum).Guild)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerGuild", Err.Number, Err.Description
End Function

Sub SetPlayerGuild(ByVal index As Long, ByVal Guild As String)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Guild = Guild
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerGuild", Err.Number, Err.Description
End Sub

Function GetPlayerGuildAccess(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerGuildAccess = Player(index).Char(Player(index).CharNum).Guildaccess
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerGuildAccess", Err.Number, Err.Description
End Function

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal Guildaccess As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Guildaccess = Guildaccess
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerGuildAccess", Err.Number, Err.Description
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerClass = Player(index).Char(Player(index).CharNum).Class
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerClass", Err.Number, Err.Description
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Class = ClassNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerClass", Err.Number, Err.Description
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerSprite = Player(index).Char(Player(index).CharNum).Sprite
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerSprite", Err.Number, Err.Description
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Sprite = Sprite
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerSprite", Err.Number, Err.Description
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerLevel = Player(index).Char(Player(index).CharNum).Level
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerLevel", Err.Number, Err.Description
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Level = Level
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerLevel", Err.Number, Err.Description
End Sub

Function GetPlayerNextLevel(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerNextLevel = Experience(GetPlayerLevel(index))
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerNextLevel", Err.Number, Err.Description
End Function

Function GetPlayerExp(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerExp = Player(index).Char(Player(index).CharNum).Exp
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerExp", Err.Number, Err.Description
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Exp = Exp
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerExp", Err.Number, Err.Description
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerAccess = Player(index).Char(Player(index).CharNum).Access
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerAccess", Err.Number, Err.Description
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Access = Access
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerAccess", Err.Number, Err.Description
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerPK = Player(index).Char(Player(index).CharNum).PK
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerPK", Err.Number, Err.Description
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).PK = PK
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerPK", Err.Number, Err.Description
End Sub

Function GetPlayerHP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerHP = Player(index).Char(Player(index).CharNum).HP
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerHP", Err.Number, Err.Description
End Function

Sub SetPlayerHP(ByVal index As Long, ByVal HP As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).HP = HP
    
    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then
        Player(index).Char(Player(index).CharNum).HP = GetPlayerMaxHP(index)
    End If
    If GetPlayerHP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).HP = 0
    End If
    Call SendStats(index)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerHP", Err.Number, Err.Description
End Sub

Function GetPlayerMP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerMP = Player(index).Char(Player(index).CharNum).MP
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMP", Err.Number, Err.Description
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        Player(index).Char(Player(index).CharNum).MP = GetPlayerMaxMP(index)
    End If
    If GetPlayerMP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).MP = 0
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerMP", Err.Number, Err.Description
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerSP = Player(index).Char(Player(index).CharNum).SP
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerSP", Err.Number, Err.Description
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        Player(index).Char(Player(index).CharNum).SP = GetPlayerMaxSP(index)
    End If
    If GetPlayerSP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).SP = 0
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerSP", Err.Number, Err.Description
End Sub

Function GetPlayerMaxHP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
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
    If GetPlayerBootsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerBootsSlot(index))).AddHP
    End If
    If GetPlayerGlovesSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerGlovesSlot(index))).AddHP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddHP
    End If
    If GetPlayerAmuletSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerAmuletSlot(index))).AddHP
    End If

    CharNum = Player(index).CharNum
    GetPlayerMaxHP = GetPlayerVIT(index) + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMaxHP", Err.Number, Err.Description
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
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
    If GetPlayerBootsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerBootsSlot(index))).AddMP
    End If
    If GetPlayerGlovesSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerGlovesSlot(index))).AddMP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddMP
    End If
    If GetPlayerAmuletSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerAmuletSlot(index))).AddMP
    End If

    CharNum = Player(index).CharNum
    GetPlayerMaxMP = GetPlayerMAGI(index) + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMaxMP", Err.Number, Err.Description
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
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
    If GetPlayerBootsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerBootsSlot(index))).AddSP
    End If
    If GetPlayerGlovesSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerGlovesSlot(index))).AddSP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddSP
    End If
    If GetPlayerAmuletSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerAmuletSlot(index))).AddSP
    End If

    CharNum = Player(index).CharNum
    GetPlayerMaxSP = GetPlayerSPEED(index) + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMaxSP", Err.Number, Err.Description
End Function

Function GetClassName(ByVal ClassNum As Long) As String
On Error GoTo ErrorHandler
    GetClassName = Trim(Class(ClassNum).Name)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassName", Err.Number, Err.Description
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassMaxHP = (1 + Int(Class(ClassNum).STR / 2) + Class(ClassNum).STR) * 2
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassMaxHP", Err.Number, Err.Description
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassMaxMP = (1 + Int(Class(ClassNum).MAGI / 2) + Class(ClassNum).MAGI) * 2
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassMaxMP", Err.Number, Err.Description
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassMaxSP = (1 + Int(Class(ClassNum).SPEED / 2) + Class(ClassNum).SPEED) * 2
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassMaxSP", Err.Number, Err.Description
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassSTR = Class(ClassNum).STR
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassSTR", Err.Number, Err.Description
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassDEF = Class(ClassNum).DEF
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassDEF", Err.Number, Err.Description
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassSPEED = Class(ClassNum).SPEED
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassSPEED", Err.Number, Err.Description
End Function

Function GetClassVIT(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassVIT = Class(ClassNum).VIT
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassVIT", Err.Number, Err.Description
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassMAGI = Class(ClassNum).MAGI
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassMAGI", Err.Number, Err.Description
End Function

Function GetPlayerBaseSTR(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBaseSTR = Player(index).Char(Player(index).CharNum).STR
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBaseSTR", Err.Number, Err.Description
End Function

Function GetPlayerBaseDEF(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBaseDEF = Player(index).Char(Player(index).CharNum).DEF
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBaseDEF", Err.Number, Err.Description
End Function

Function GetPlayerBaseMAGI(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBaseMAGI = Player(index).Char(Player(index).CharNum).MAGI
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBaseMAGI", Err.Number, Err.Description
End Function

Function GetPlayerBaseSPEED(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBaseSPEED = Player(index).Char(Player(index).CharNum).SPEED
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBaseSPEED", Err.Number, Err.Description
End Function

Function GetPlayerBaseVIT(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBaseVIT = Player(index).Char(Player(index).CharNum).VIT
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBaseVIT", Err.Number, Err.Description
End Function

Function GetPlayerSTR(ByVal index As Long) As Long
On Error GoTo ErrorHandler
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
    If GetPlayerBootsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerBootsSlot(index))).AddStr
    End If
    If GetPlayerGlovesSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerGlovesSlot(index))).AddStr
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddStr
    End If
    If GetPlayerAmuletSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerAmuletSlot(index))).AddStr
    End If
    GetPlayerSTR = Player(index).Char(Player(index).CharNum).STR + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerSTR", Err.Number, Err.Description
End Function

Sub SetPlayerSTR(ByVal index As Long, ByVal STR As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).STR = STR
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerSTR", Err.Number, Err.Description
End Sub

Function GetPlayerDEF(ByVal index As Long) As Long
On Error GoTo ErrorHandler
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
    If GetPlayerBootsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerBootsSlot(index))).AddDef
    End If
    If GetPlayerGlovesSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerGlovesSlot(index))).AddDef
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddDef
    End If
    If GetPlayerAmuletSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerAmuletSlot(index))).AddDef
    End If
    GetPlayerDEF = Player(index).Char(Player(index).CharNum).DEF + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerDEF", Err.Number, Err.Description
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal DEF As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).DEF = DEF
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerDEF", Err.Number, Err.Description
End Sub

Function GetPlayerVIT(ByVal index As Long) As Long
On Error GoTo ErrorHandler
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).Data3
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).Data3
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).Data3
    End If
    If GetPlayerBootsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerBootsSlot(index))).Data3
    End If
    If GetPlayerGlovesSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerGlovesSlot(index))).Data3
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).Data3
    End If
    If GetPlayerAmuletSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerAmuletSlot(index))).Data3
    End If
    GetPlayerVIT = Player(index).Char(Player(index).CharNum).VIT + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerVIT", Err.Number, Err.Description
End Function

Sub SetPlayerVIT(ByVal index As Long, ByVal VIT As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).VIT = VIT
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerVIT", Err.Number, Err.Description
End Sub

Function GetPlayerSPEED(ByVal index As Long) As Long
On Error GoTo ErrorHandler
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
    If GetPlayerBootsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerBootsSlot(index))).AddSpeed
    End If
    If GetPlayerGlovesSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerGlovesSlot(index))).AddSpeed
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddSpeed
    End If
    If GetPlayerAmuletSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerAmuletSlot(index))).AddSpeed
    End If
    
    GetPlayerSPEED = Player(index).Char(Player(index).CharNum).SPEED + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerSPEED", Err.Number, Err.Description
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal SPEED As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).SPEED = SPEED
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerSPEED", Err.Number, Err.Description
End Sub

Function GetPlayerMAGI(ByVal index As Long) As Long
On Error GoTo ErrorHandler
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
     If GetPlayerBootsSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerBootsSlot(index))).AddMagi
    End If
    If GetPlayerGlovesSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerGlovesSlot(index))).AddMagi
    End If
    If GetPlayerRingSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddMagi
    End If
    If GetPlayerAmuletSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerAmuletSlot(index))).AddMagi
    End If
    GetPlayerMAGI = Player(index).Char(Player(index).CharNum).MAGI + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMAGI", Err.Number, Err.Description
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal MAGI As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).MAGI = MAGI
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerMAGI", Err.Number, Err.Description
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerPOINTS = Player(index).Char(Player(index).CharNum).POINTS
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerPOINTS", Err.Number, Err.Description
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).POINTS = POINTS
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerPOINTS", Err.Number, Err.Description
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerMap = Player(index).Char(Player(index).CharNum).Map
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMap", Err.Number, Err.Description
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
On Error GoTo ErrorHandler
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(index).Char(Player(index).CharNum).Map = MapNum
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerMap", Err.Number, Err.Description
End Sub

Function GetPlayerX(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerX = Player(index).Char(Player(index).CharNum).x
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerX", Err.Number, Err.Description
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).x = x
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerX", Err.Number, Err.Description
End Sub

Function GetPlayerY(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerY = Player(index).Char(Player(index).CharNum).y
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerY", Err.Number, Err.Description
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).y = y
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerY", Err.Number, Err.Description
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerDir = Player(index).Char(Player(index).CharNum).Dir
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerDir", Err.Number, Err.Description
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Dir = Dir
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerDir", Err.Number, Err.Description
End Sub

Function GetPlayerIP(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerIP", Err.Number, Err.Description
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerInvItemNum = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Num
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerInvItemNum", Err.Number, Err.Description
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Num = ItemNum
    Call SendInventoryUpdate(index, InvSlot)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerInvItemNum", Err.Number, Err.Description
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerInvItemValue = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerInvItemValue", Err.Number, Err.Description
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value = ItemValue
    Call SendInventoryUpdate(index, InvSlot)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerInvItemValue", Err.Number, Err.Description
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerInvItemDur = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerInvItemDur", Err.Number, Err.Description
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur = ItemDur
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerInvItemDur", Err.Number, Err.Description
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerSpell = Player(index).Char(Player(index).CharNum).Spell(SpellSlot)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerSpell", Err.Number, Err.Description
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Spell(SpellSlot) = SpellNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerSpell", Err.Number, Err.Description
End Sub

Function GetPlayerArmorSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerArmorSlot = Player(index).Char(Player(index).CharNum).ArmorSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerArmorSlot", Err.Number, Err.Description
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).ArmorSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerArmorSlot", Err.Number, Err.Description
End Sub

Function GetPlayerWeaponSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerWeaponSlot = Player(index).Char(Player(index).CharNum).WeaponSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerWeaponSlot", Err.Number, Err.Description
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).WeaponSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerWeaponSlot", Err.Number, Err.Description
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerHelmetSlot = Player(index).Char(Player(index).CharNum).HelmetSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerHelmetSlot", Err.Number, Err.Description
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).HelmetSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerHelmetSlot", Err.Number, Err.Description
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerShieldSlot = Player(index).Char(Player(index).CharNum).ShieldSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerShieldSlot", Err.Number, Err.Description
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).ShieldSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerShieldSlot", Err.Number, Err.Description
End Sub
Function GetPlayerBootsSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBootsSlot = Player(index).Char(Player(index).CharNum).BootsSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBootsSlot", Err.Number, Err.Description
End Function

Sub SetPlayerBootsSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).BootsSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerBootsSlot", Err.Number, Err.Description
End Sub
Function GetPlayerGlovesSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerGlovesSlot = Player(index).Char(Player(index).CharNum).GlovesSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerGlovesSlot", Err.Number, Err.Description
End Function

Sub SetPlayerGlovesSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).GlovesSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerGlovesSlot", Err.Number, Err.Description
End Sub
Function GetPlayerRingSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerRingSlot = Player(index).Char(Player(index).CharNum).RingSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerRingSlot", Err.Number, Err.Description
End Function

Sub SetPlayerRingSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).RingSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerRingSlot", Err.Number, Err.Description
End Sub
Function GetPlayerAmuletSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerAmuletSlot = Player(index).Char(Player(index).CharNum).AmuletSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerAmuletSlot", Err.Number, Err.Description
End Function

Sub SetPlayerAmuletSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).AmuletSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerAmuletSlot", Err.Number, Err.Description
End Sub

Sub BattleMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long, ByVal Who As Long)
On Error GoTo ErrorHandler
    Call SendDataTo(index, "damagedisplay" & SEP_CHAR & Who & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "BattleMsg", Err.Number, Err.Description
End Sub

Function GetPlayerClassName(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerClassName = Class(GetPlayerClass(index)).Name
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerClassName", Err.Number, Err.Description
End Function

Sub PlaySound(ByVal index As Long, ByVal Sound As String)
On Error GoTo ErrorHandler
    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "PlaySound", Err.Number, Err.Description
End Sub

Function Rand(ByVal High As Long, ByVal Low As Long)
On Error GoTo ErrorHandler
Rand = -99999999
Randomize
High = High + 1
Do Until Rand >= Low
    Rand = Int(Rnd * High)
Loop
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "Rand", Err.Number, Err.Description
End Function

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBankItemNum = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Num
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBankItemNum", Err.Number, Err.Description
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Num = ItemNum
    Call SendBankUpdate(index, BankSlot)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerBankItemNum", Err.Number, Err.Description
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBankItemValue = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Value
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBankItemValue", Err.Number, Err.Description
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Value = ItemValue
    Call SendBankUpdate(index, BankSlot)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerBankItemValue", Err.Number, Err.Description
End Sub

Function GetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBankItemDur = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Dur
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBankItemDur", Err.Number, Err.Description
End Function

Sub SetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Dur = ItemDur
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerBankItemDur", Err.Number, Err.Description
End Sub
