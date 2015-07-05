Attribute VB_Name = "modTypes"
Option Explicit

Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Type BankItemRec
    Num As Long
    Value As Long
    Dur As Long
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Sex As Byte
    Class As Byte
    Sprite As Integer
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    Guild As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Long
    DEF As Long
    SPEED As Long
    MAGI As Long
    POINTS As Byte
    
    ' Worn equipment
    ArmorSlot As Byte
    WeaponSlot As Byte
    HelmetSlot As Byte
    ShieldSlot As Byte
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    BankInv(1 To MAX_BANK_ITEMS) As BankItemRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
    
    ' Position
    Map As Integer
    x As Byte
    y As Byte
    Dir As Byte
    
    'Player's text
    Text As String
    
    'Friends
    Friends(1 To MAX_FRIENDS) As String
    
    'Non-char data
    Trackers(1 To MAX_TRACKERS) As String
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
End Type

Type TileRec
    Ground As Integer
    Mask As Integer
    Mask2 As Integer
    Anim As Integer
    Anim2 As Integer
    Fringe As Integer
    FringeAnim As Integer
    Fringe2 As Integer
    Walkable As Byte
    Blocked As Byte
    Warp As Byte
    WarpMap As Integer
    WarpX As Byte
    WarpY As Byte
    Item As Byte
    ItemNum As Integer
    ItemValue As Integer
    NpcAvoid As Byte
    Key As Byte
    KeyNum As Long
    KeyTake As Byte
    KeyOpen As Byte
    KeyOpenX As Byte
    KeyOpenY As Byte
    North As Byte
    West As Byte
    East As Byte
    South As Byte
    Shop As Byte
    ShopNum As Integer
    Bank As Byte
    Heal As Byte
    HealValue As Integer
    Damage As Byte
    DamageValue As Integer
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
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    
    Sprite As Integer
    
    HP As Long
    MP As Long
    SP As Long
    
    STR As Long
    DEF As Long
    SPEED As Long
    MAGI As Long
    
    'Server-side only -smchronos
    Map As Long
    x As Long
    y As Long
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    Description As String * 50
    
    Pic As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    Data4 As Integer
    Data5 As Integer
    
    'A string to hold the name of the .wav file
    Sound As String
End Type

Type MapItemRec
    Num As Byte
    Value As Long
    Dur As Integer
    
    x As Byte
    y As Byte
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 500
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    
    HP As Long
    STR As Long
    DEF As Long
    SPEED As Long
    MAGI As Long
    EXP As Long
    
    'To see if the npc shall flee when its HP is low
    '-smchronos
    Fear As Boolean
    
    'To see if the npc can cast spells
    '-smchronos
    MagicUser As Boolean
    
    'The total number of magic spells
    NpcSpell(1 To MAX_NPC_SPELLS) As Integer
End Type

Type MapNpcRec
    Num As Integer
    
    Target As Integer
    
    'The map npc's behavior setting
    Behavior As Byte
    
    HP As Long
    MaxHP As Long
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
    Stock As Long
    MaxStock As Long
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 255
    LeaveSay As String * 255
    FixesItems As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
    Restock As Long
End Type
    
Type SpellRec
    Name As String * NAME_LENGTH
    ClassReq As Byte
    LevelReq As Byte
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Type TempTileRec
    DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY)  As Byte
    DoorTimer As Long
End Type

Type GuildRec
    Name As String * NAME_LENGTH
    Founder As String * NAME_LENGTH
    Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
End Type

'SADScript Type
Public Type define
    sVari As String
    sValue As String
End Type

'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
