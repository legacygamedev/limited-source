Attribute VB_Name = "modTypes"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

' Public data structures
Public Map() As MapRec
Public MapSpawn() As MapSpawnRec
Public MapCache() As String
Public TempTile() As TempTileRec
Public PlayersOnMap() As Boolean
Public Class() As ClassRec
Public MapItem() As MapItemRec
Public MapNpc() As MapNpcHoldRec
Public WalkFrame() As Byte
Public AttackFrame() As Byte
Public Direction_Anim(0 To 3) As Byte
Public Player() As AccountRec
Public TempPlayer() As TempPlayerRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Sign() As SignRec
Public Guild() As GuildRec
Public Animation() As AnimationRec

Type AnimationRec
    Name As String * 10
    Width As Long
    Height As Long
    Delay As Integer
    Pic As Long
End Type

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
    
    Muted As Boolean
    MuteTime As Currency
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Long
    POINTS As Long
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Position
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte
    
    Guild As Long
    GuildRank As Long
End Type

Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
       
    ' Characters (we use 0 to prevent a crash that still needs to be figured out)
    Char(0 To MAX_CHARS) As PlayerRec
End Type

Type TempPlayerRec
    ' Non saved local vars
    Buffer As String
    CharNum As Byte
    InGame As Boolean
    AttackTimer As Currency
    CastTimer(1 To MAX_PLAYER_SPELLS) As Currency
    DataTimer As Currency
    DataBytes As Long
    DataPackets As Long
    PartyPlayer As Long
    InParty As Byte
    TargetType As Byte
    Target As Byte
    CastedSpell As Byte
    PartyStarter As Byte
    GettingMap As Byte
    GInviteWaiting As Boolean
    DOT_Tile As Currency
End Type

Type TileRec
    Layer(0 To 3) As Integer
    LayerSet(0 To 3) As Byte
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Type NpcSpawnRec
    Num As Long
    X As Long
    Y As Long
End Type

Type MapSpawnRec
    Npc() As NpcSpawnRec
End Type

Type MapRec
    Name As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    Music As String * 50
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
End Type

Type ClassStartLocRec
    MapNum As Long
    X As Long
    Y As Long
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    Sprite As Integer
    Stat(1 To Stats.Stat_Count - 1) As Byte
    StartLoc As ClassStartLocRec
    PointsPerLevel As Byte
End Type

Type ItemRec
    Name As String
    
    Pic As Integer
    Type As Byte
    Durability As Integer
    Anim As Long
    
    CostItem As Long
    CostAmount As Long
    
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    
    BuffStats(1 To Stats.Stat_Count - 1) As Integer
    BuffVitals(1 To Vitals.Vital_Count - 1) As Integer
    Required(0 To Item_Requires.Count - 1) As Integer
End Type

Type MapItemRec
    Num As Byte
    Value As Long
    Dur As Integer
    Anim As Long
    
    X As Byte
    Y As Byte
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    HP As Long
    Experience As Long
    Sound(0 To 2) As String * 50
    Reflection(0 To 1) As Integer
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Long
    
    GivesGuild As Byte
End Type

Type MapNpcRec
    Num As Integer
    
    Target As Integer
    
    Vital(1 To Vitals.Vital_Count - 1) As Long
        
    X As Byte
    Y As Byte
    Dir As Integer
    
    ' For server use only
    SpawnWait As Currency
    AttackTimer As Currency
End Type

Type MapNpcHoldRec
    MapNpc() As MapNpcRec
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
    CastSound As String * 50
    Timer As Long
    MPReq As Long
    Type As Byte
    Anim As Byte
    Range As Byte
    AOE As Byte
    Icon As Integer
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Type TempTileRec
    DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY)  As Byte
    DoorTimer As Currency
End Type

Type SignRec
    Name As String * 10
    Section() As String
End Type

Type GuildRec
    Name As String
    TotalMembers As Long
    Member_Account() As String * NAME_LENGTH
    Member_CharNum() As Long
End Type
