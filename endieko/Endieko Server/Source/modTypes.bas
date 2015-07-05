Attribute VB_Name = "modTypes"
Option Explicit
Global PlayerI As Integer

' Winsock globals
Public Const GAME_PORT = 4000
Public Const GAME_IP = "0.0.0.0" 'You can leave this, or use your IP.

' Packet Buffer Globals
Public Const MAX_PACKETLEN As Long = 32768 'About 32K

' General constants
Public Const GAME_NAME = "Endieko Online"
Public Const MAX_PLAYERS = 500 'For starters... More optimization needed for a lot more.
Public Const MAX_SPELLS = 1000
Public Const MAX_SHOPS = 1000
Public Const MAX_ITEMS = 1000
Public Const MAX_NPCS = 1000
Public Const MAX_MAP_ITEMS = 20
Public Const MAX_GUILDS = 20
Public Const MAX_GUILD_MEMBERS = 15
Public Const MAX_PARTY_MEMBERS = 8
Public Const MAX_EMOTICONS = 100
Public Const MAX_LEVEL = 100
Public Const MAX_ARROWS = 100
Public Const MAX_BANK = 20
Public Const MAX_INV = 24
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 12
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10
Public Const MAX_BANS = 50
Public Const MAX_MAPS = 1000
Public Const Scripting = 1
Public Const MAX_EFFECTS = 100

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 1

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 = "88978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672auygewuifaunfauf"
Public Const SEC_CODE2 = "ksisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlaEND_CHARowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas"
Public Const SEC_CODE3 = "saiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi"
Public Const SEC_CODE4 = "88978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672"
' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map constants
Public Const MAX_MAPX = 30
Public Const MAX_MAPY = 30
Public Const JAIL_MAP = 255
Public Const JAIL_X = 1
Public Const JAIL_Y = 30
'Public MAX_MAPX As Byte
'Public MAX_MAPY As Byte
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

' Chat Constants
Public GlobalChatDisabled As Byte
Public BroadcastChatDisabled As Byte
Public MapChatDisabled As Byte
Public PrivateChatDisabled As Byte
Public EmoteChatDisabled As Byte
Public GuildChatDisabled As Byte
Public PartyChatDisabled As Byte
Public AdminChatDisabled As Byte

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

' Target type constants
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1

' Effect Constats
Public Const EFFECT_TYPE_DRAIN = 0
Public Const EFFECT_TYPE_FORTIFY = 1
Public Const EFFECT_TYPE_FREEZE = 2

Type EffectRec
    Name As String * NAME_LENGTH
    Effect As Byte
    Time As Byte
    Data1 As Byte
    Data2 As Byte
    Data3 As Byte
End Type

Type PartyRec
    InParty As Byte
    Started As Byte
    PlayerNums(1 To MAX_PARTY_MEMBERS) As Long
End Type

Type BanRec
    BannedIP As String
    BannedChar As String
    BannedBy As String
    BannedHD As String
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
End Type

Type BankInvRec
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
    POINTS As Long
    
    ' Worn equipment
    ArmorSlot As Byte
    WeaponSlot As Byte
    HelmetSlot As Byte
    ShieldSlot As Byte
    LegSlot As Byte
    BootSlot As Byte
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Bank(1 To MAX_BANK) As BankInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
    
    ' Position
    Map As Integer
    x As Byte
    y As Byte
    Dir As Byte
    
    ' Muted
    BroadcastMute As Byte
    GlobalMute As Byte
    AdminMute As Byte
    MapMute As Byte
    EmotMute As Byte
    PrivMute As Byte
    GuildMute As Byte
    PartyMute As Byte
    Jailed As Byte
    
    ' New
    Alignment As Integer
    FishingLevel As Byte
    FishingExp As Integer
    MiningLevel As Byte
    MiningExp As Integer
    LumberLevel As Byte
    LumberExp As Integer
    
    ' Status
    Status As Byte
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
    HDSerial As String
    MuteTimer As Long
    JailTimer As Long
    GuildInvitation As Boolean
    GuildTemp As String
    GuildInviter As Long
    TempSprite As Integer
    Invisible As Byte
    StatusTimer As Long

    InTrade As Byte
    TradePlayer As Long
    TradeOk As Byte
    TradeItemMax As Byte
    TradeItemMax2 As Byte
    Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
    
    InChat As Byte
    ChatPlayer As Long
    
    Party As PartyRec
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
    Alignment As Integer
    
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
    Pic As Integer
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
    PlayerName(1 To MAX_GUILD_MEMBERS) As String
    Rank(1 To MAX_GUILD_MEMBERS) As Byte
End Type

Type EmoRec
    Pic As Long
    Command As String
End Type

Type ConDataQueue
  Lock As Boolean
  Lines As String
End Type

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

Public Map() As MapRec
Public TempTile() As TempTileRec
Public PlayersOnMap() As Long
Public Player(1 To MAX_PLAYERS) As AccountRec
Public Class() As ClassRec
Public Class2() As ClassRec
Public Class3() As ClassRec
Public Item(0 To MAX_ITEMS) As ItemRec
Public Npc(0 To MAX_NPCS) As NpcRec
Public MapItem() As MapItemRec
Public MapNpc() As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Guild(1 To MAX_GUILDS) As GuildRec
Public Emoticons(0 To MAX_EMOTICONS) As EmoRec
Public Experience(1 To MAX_LEVEL) As Long
Public Arrows(1 To MAX_ARROWS) As ArrowRec
Public Ban(0 To MAX_BANS) As BanRec
Public Effect(1 To MAX_EFFECTS) As EffectRec
Public ConQueues(MAX_PLAYERS) As ConDataQueue
Public QueueDisconnect(MAX_PLAYERS) As Boolean

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
        Class(i).Name = vbNullString
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
        Class2(i).Name = vbNullString
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
        Class3(i).Name = vbNullString
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

Sub ClearPlayer(ByVal Index As Long)
Dim i As Long
Dim n As Long

    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    
    For i = 1 To MAX_CHARS
        Player(Index).Char(i).Name = vbNullString
        Player(Index).Char(i).Class = 0
        Player(Index).Char(i).Level = 0
        Player(Index).Char(i).Sprite = 0
        Player(Index).Char(i).Exp = 0
        Player(Index).Char(i).Access = 0
        Player(Index).Char(i).PK = NO
        Player(Index).Char(i).POINTS = 0
        Player(Index).Char(i).Guild = vbNullString
        Player(Index).Char(i).Guildaccess = 0
        
        Player(Index).Char(i).HP = 0
        Player(Index).Char(i).MP = 0
        Player(Index).Char(i).SP = 0
        
        Player(Index).Char(i).STR = 0
        Player(Index).Char(i).DEF = 0
        Player(Index).Char(i).SPEED = 0
        Player(Index).Char(i).MAGI = 0
        
        Player(Index).Char(i).BroadcastMute = 0
        Player(Index).Char(i).GlobalMute = 0
        Player(Index).Char(i).AdminMute = 0
        Player(Index).Char(i).MapMute = 0
        Player(Index).Char(i).EmotMute = 0
        Player(Index).Char(i).PrivMute = 0
        Player(Index).Char(i).GuildMute = 0
        Player(Index).Char(i).PartyMute = 0
        Player(Index).Char(i).Jailed = 0
        
        Player(Index).Char(i).Alignment = 0
        Player(Index).Char(i).FishingLevel = 0
        Player(Index).Char(i).FishingExp = 0
        Player(Index).Char(i).MiningLevel = 0
        Player(Index).Char(i).MiningExp = 0
        Player(Index).Char(i).LumberLevel = 0
        Player(Index).Char(i).LumberExp = 0
        
        Player(Index).Char(i).Status = 0
        
        For n = 1 To MAX_INV
            Player(Index).Char(i).Inv(n).Num = 0
            Player(Index).Char(i).Inv(n).Value = 0
            Player(Index).Char(i).Inv(n).Dur = 0
        Next n
        
        For n = 1 To MAX_BANK
            Player(Index).Char(i).Bank(n).Num = 0
            Player(Index).Char(i).Bank(n).Value = 0
            Player(Index).Char(i).Bank(n).Dur = 0
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(i).Spell(n) = 0
        Next n
        
        Player(Index).Char(i).ArmorSlot = 0
        Player(Index).Char(i).WeaponSlot = 0
        Player(Index).Char(i).HelmetSlot = 0
        Player(Index).Char(i).ShieldSlot = 0
        Player(Index).Char(i).LegSlot = 0
        Player(Index).Char(i).BootSlot = 0
        
        Player(Index).Char(i).Map = 0
        Player(Index).Char(i).x = 0
        Player(Index).Char(i).y = 0
        Player(Index).Char(i).Dir = 0
        
        ' Temporary vars
        Player(Index).Buffer = vbNullString
        Player(Index).IncBuffer = vbNullString
        Player(Index).CharNum = 0
        Player(Index).InGame = False
        Player(Index).AttackTimer = 0
        Player(Index).DataTimer = 0
        Player(Index).DataBytes = 0
        Player(Index).DataPackets = 0
        Player(Index).PartyPlayer = 0
        Player(Index).InParty = 0
        Player(Index).Target = 0
        Player(Index).TargetType = 0
        Player(Index).CastedSpell = NO
        Player(Index).PartyStarter = NO
        Player(Index).GettingMap = NO
        Player(Index).Emoticon = -1
        Player(Index).InTrade = 0
        Player(Index).TradePlayer = 0
        Player(Index).TradeOk = 0
        Player(Index).TradeItemMax = 0
        Player(Index).TradeItemMax2 = 0
        For n = 1 To MAX_PLAYER_TRADES
            Player(Index).Trading(n).InvName = vbNullString
            Player(Index).Trading(n).InvNum = 0
        Next n
        Player(Index).InTrade = 0
        Player(Index).ChatPlayer = 0
        Player(Index).JailTimer = 0
        Player(Index).MuteTimer = 0
        Player(Index).GuildInvitation = False
        Player(Index).GuildInviter = 0
        Player(Index).GuildTemp = vbNullString
        For n = 1 To MAX_PARTY_MEMBERS
            Player(Index).Party.PlayerNums(n) = 0
        Next n
        Player(Index).Party.InParty = 0
        Player(Index).Party.Started = NO
        Player(Index).TempSprite = 0
        Player(Index).Invisible = 0
        Player(Index).StatusTimer = 0
    Next i
End Sub

Sub ClearChar(ByVal Index As Long, ByVal CharNum As Long)
Dim n As Long
    
    Player(Index).Char(CharNum).Name = vbNullString
    Player(Index).Char(CharNum).Class = 0
    Player(Index).Char(CharNum).Sprite = 0
    Player(Index).Char(CharNum).Level = 0
    Player(Index).Char(CharNum).Exp = 0
    Player(Index).Char(CharNum).Access = 0
    Player(Index).Char(CharNum).PK = NO
    Player(Index).Char(CharNum).POINTS = 0
    Player(Index).Char(CharNum).Guild = vbNullString
    Player(Index).Char(CharNum).Guildaccess = 0
    
    Player(Index).Char(CharNum).HP = 0
    Player(Index).Char(CharNum).MP = 0
    Player(Index).Char(CharNum).SP = 0
    
    Player(Index).Char(CharNum).STR = 0
    Player(Index).Char(CharNum).DEF = 0
    Player(Index).Char(CharNum).SPEED = 0
    Player(Index).Char(CharNum).MAGI = 0
    
    Player(Index).Char(CharNum).BroadcastMute = 0
    Player(Index).Char(CharNum).GlobalMute = 0
    Player(Index).Char(CharNum).AdminMute = 0
    Player(Index).Char(CharNum).MapMute = 0
    Player(Index).Char(CharNum).EmotMute = 0
    Player(Index).Char(CharNum).PrivMute = 0
    Player(Index).Char(CharNum).GuildMute = 0
    Player(Index).Char(CharNum).PartyMute = 0
    Player(Index).Char(CharNum).Jailed = 0
    
    Player(Index).Char(CharNum).Alignment = 0
    Player(Index).Char(CharNum).FishingLevel = 0
    Player(Index).Char(CharNum).FishingExp = 0
    Player(Index).Char(CharNum).MiningLevel = 0
    Player(Index).Char(CharNum).MiningExp = 0
    Player(Index).Char(CharNum).LumberLevel = 0
    Player(Index).Char(CharNum).LumberExp = 0
    
    Player(Index).Char(CharNum).Status = 0
    
    For n = 1 To MAX_INV
        Player(Index).Char(CharNum).Inv(n).Num = 0
        Player(Index).Char(CharNum).Inv(n).Value = 0
        Player(Index).Char(CharNum).Inv(n).Dur = 0
    Next n
    
    For n = 1 To MAX_BANK
        Player(Index).Char(CharNum).Bank(n).Num = 0
        Player(Index).Char(CharNum).Bank(n).Value = 0
        Player(Index).Char(CharNum).Bank(n).Dur = 0
    Next n
    
    For n = 1 To MAX_PLAYER_SPELLS
        Player(Index).Char(CharNum).Spell(n) = 0
    Next n
    
    Player(Index).Char(CharNum).ArmorSlot = 0
    Player(Index).Char(CharNum).WeaponSlot = 0
    Player(Index).Char(CharNum).HelmetSlot = 0
    Player(Index).Char(CharNum).ShieldSlot = 0
    Player(Index).Char(CharNum).LegSlot = 0
    Player(Index).Char(CharNum).BootSlot = 0
    
    Player(Index).Char(CharNum).Map = 0
    Player(Index).Char(CharNum).x = 0
    Player(Index).Char(CharNum).y = 0
    Player(Index).Char(CharNum).Dir = 0
End Sub
    
Sub ClearItem(ByVal Index As Long)
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    
    Item(Index).Type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
    Item(Index).StrReq = 0
    Item(Index).DefReq = 0
    Item(Index).SpeedReq = 0
    Item(Index).MagiReq = 0
    Item(Index).ClassReq = -1
    Item(Index).AccessReq = 0
    Item(Index).CannotBeRepaired = 0
    Item(Index).DropOnDeath = 0
    
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

Sub ClearEffect(ByVal Index As Long)
    Effect(Index).Name = vbNullString
    Effect(Index).Effect = 0
    Effect(Index).Time = 0
    Effect(Index).Data1 = 0
    Effect(Index).Data2 = 0
    Effect(Index).Data3 = 0
End Sub

Sub ClearEffects()
Dim i As Long
    
    For i = 1 To MAX_EFFECTS
        Call ClearEffect(i)
    Next i
End Sub

Sub ClearNpc(ByVal Index As Long)
Dim i As Long
    Npc(Index).Name = vbNullString
    Npc(Index).AttackSay = vbNullString
    Npc(Index).Sprite = 0
    Npc(Index).SpawnSecs = 0
    Npc(Index).Behavior = 0
    Npc(Index).Range = 0
    Npc(Index).STR = 0
    Npc(Index).DEF = 0
    Npc(Index).SPEED = 0
    Npc(Index).MAGI = 0
    Npc(Index).Big = 0
    Npc(Index).MaxHp = 0
    Npc(Index).Exp = 0
    For i = 1 To MAX_NPC_DROPS
        Npc(Index).ItemNPC(i).Chance = 0
        Npc(Index).ItemNPC(i).ItemNum = 0
        Npc(Index).ItemNPC(i).ItemValue = 0
    Next i
End Sub

Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next i
End Sub

Sub ClearArrow(ByVal Index As Long)
    Arrows(Index).Name = vbNullString
    Arrows(Index).Pic = 0
    Arrows(Index).Range = 0
End Sub

Sub ClearArrows()
Dim i As Long

    For i = 1 To MAX_ARROWS
        Call ClearArrow(i)
    Next i
End Sub

Sub ClearEmos()
Dim i As Long

    For i = 0 To MAX_EMOTICONS
        Emoticons(i).Pic = 0
        Emoticons(i).Command = vbNullString
    Next i
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    MapItem(MapNum, Index).Num = 0
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
        Next x
    Next y
End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    MapNpc(MapNum, Index).Num = 0
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

    Map(MapNum).Name = vbNullString
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
            Map(MapNum).Tile(x, y).String1 = vbNullString
            Map(MapNum).Tile(x, y).String2 = vbNullString
            Map(MapNum).Tile(x, y).String3 = vbNullString
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

Sub ClearShop(ByVal Index As Long)
Dim i As Long

    Shop(Index).Name = vbNullString
    Shop(Index).JoinSay = vbNullString
    Shop(Index).LeaveSay = vbNullString
    
    For i = 1 To MAX_TRADES
        Shop(Index).TradeItem(i).GiveItem = 0
        Shop(Index).TradeItem(i).GiveValue = 0
        Shop(Index).TradeItem(i).GetItem = 0
        Shop(Index).TradeItem(i).GetValue = 0
    Next i
End Sub

Sub ClearShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next i
End Sub

Sub ClearSpell(ByVal Index As Long)
    Spell(Index).Name = vbNullString
    Spell(Index).ClassReq = 0
    Spell(Index).LevelReq = 0
    Spell(Index).Type = 0
    Spell(Index).Data1 = 0
    Spell(Index).Data2 = 0
    Spell(Index).Data3 = 0
    Spell(Index).MPCost = 0
    Spell(Index).Sound = 0
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

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Char(Player(Index).CharNum).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Char(Player(Index).CharNum).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Guildaccess As Long)
    Player(Index).Char(Player(Index).CharNum).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Char(Player(Index).CharNum).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Char(Player(Index).CharNum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Char(Player(Index).CharNum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(Player(Index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(Player(Index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).Char(Player(Index).CharNum).HP = HP
    
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    End If
    If GetPlayerHP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).HP = 0
    End If
    Call SendStats(Index)
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    End If
    If GetPlayerMP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).MP = 0
    End If
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    End If
    If GetPlayerSP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).SP = 0
    End If
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
Dim CharNum As Long
Dim i As Long
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

    CharNum = Player(Index).CharNum
    GetPlayerMaxHP = ((Player(Index).Char(CharNum).Level + Int(GetPlayerSTR(Index) / 2) + Class(Player(Index).Char(CharNum).Class).STR) * 2) + Add
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

    CharNum = Player(Index).CharNum
    GetPlayerMaxMP = ((Player(Index).Char(CharNum).Level + Int(GetPlayerMAGI(Index) / 2) + Class(Player(Index).Char(CharNum).Class).MAGI) * 2) + Add
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

    CharNum = Player(Index).CharNum
    GetPlayerMaxSP = ((Player(Index).Char(CharNum).Level + Int(GetPlayerSPEED(Index) / 2) + Class(Player(Index).Char(CharNum).Class).SPEED) * 2) + Add
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
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

Function GetPlayerSTR(ByVal Index As Long) As Long
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
    GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).STR + Add
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

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
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF + Add
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub

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
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).SPEED + Add
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal SPEED As Long)
    Player(Index).Char(Player(Index).CharNum).SPEED = SPEED
End Sub

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
    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).MAGI + Add
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal MAGI As Long)
    Player(Index).Char(Player(Index).CharNum).MAGI = MAGI
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(Player(Index).CharNum).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).Char(Player(Index).CharNum).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).Char(Player(Index).CharNum).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
    'GetPlayerIP = frmServer.Socket(index).RemoteHostIP --Pre IOCP
    GetPlayerIP = GameServer.Sockets(Index).RemoteAddress
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).Char(Player(Index).CharNum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).Char(Player(Index).CharNum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).Char(Player(Index).CharNum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).Char(Player(Index).CharNum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ShieldSlot = InvNum
End Sub

Function GetPlayerLegSlot(ByVal Index As Long)
    GetPlayerLegSlot = Player(Index).Char(Player(Index).CharNum).LegSlot
End Function

Sub SetPlayerLegSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).LegSlot = InvNum
End Sub

Function GetPlayerBootSlot(ByVal Index As Long)
    GetPlayerBootSlot = Player(Index).Char(Player(Index).CharNum).BootSlot
End Function

Sub SetPLayerBootSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).LegSlot = InvNum
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemNum = Player(Index).Char(Player(Index).CharNum).Bank(InvSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(InvSlot).Num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Char(Player(Index).CharNum).Bank(InvSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(InvSlot).Value = ItemValue
End Sub

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemDur = Player(Index).Char(Player(Index).CharNum).Bank(InvSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerHD(ByVal Index As Long) As String
    GetPlayerHD = Player(Index).HDSerial
End Function

Sub SetPlayerMute(ByVal Index As Long, YesNo As Long)
    Player(Index).Char(Player(Index).CharNum).Muted = YesNo
End Sub

Sub SetPlayerJail(ByVal Index As Long, YesNo As Long)
    Player(Index).Char(Player(Index).CharNum).Jailed = YesNo
End Sub

Function GetPlayerAlignment(ByVal Index As Long) As Long
    GetPlayerAlignment = Player(Index).Char(Player(Index).CharNum).Alignment
End Function

Sub SetPlayerAlignment(ByVal Index As Long, ByVal Alignment As Long)
    Player(Index).Char(Player(Index).CharNum).Alignment = Alignment
End Sub

Sub SetPlayerInvisible(ByVal Index As Long, ByVal YesNo As Boolean)
    If YesNo = True Then
        Player(Index).Invisible = 1
        Player(Index).TempSprite = GetPlayerSprite(Index)
        Call SetPlayerSprite(Index, 1000)
    Else
        Player(Index).Invisible = 0
        Call SetPlayerSprite(Index, Player(Index).TempSprite)
    End If
End Sub
