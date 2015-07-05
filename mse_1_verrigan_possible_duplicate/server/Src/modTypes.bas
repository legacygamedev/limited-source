Attribute VB_Name = "modTypes"
Option Explicit

Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Sex As Byte
    Class As Byte
    Sprite As Integer
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    Guild As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Byte
    DEF As Byte
    SPEED As Byte
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
    
    ' Position
    Map As Integer
    x As Byte
    y As Byte
    Dir As Byte
End Type
    
Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
       
    ' Characters (we use 0 to prevent a crash that still needs to be figured out)
    Char(0 To MAX_CHARS) As PlayerRec
    
    ' None saved local vars
    Buffer() As Byte
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
    Anim As Integer
    Fringe As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
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
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
End Type

Type MapRec
    MapName(NAME_LENGTH - 1) As Byte
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
    Indoors As Byte
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
End Type

Type ClassRec
    ClassName(NAME_LENGTH - 1) As Byte
    
    Sprite As Integer
    
    STR As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte
End Type

Type ItemRec
    ItemName(NAME_LENGTH - 1) As Byte
    
    Pic As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
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
    AttackSay As String * 255
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    
    STR  As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte
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

