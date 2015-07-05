Attribute VB_Name = "modTypes"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

' Public data structures
Public Map As MapRec
Public MapSpawn As MapSpawnRec
Public TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
Public Class() As ClassRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc() As MapNpcRec
Public Config As ConfigRec
Public GameConfig As GameConfigRec
Public Direction_Anim(0 To 3) As Byte
Public Sprite_Size() As Sprite_SizeRec
Public Animations(1 To 100) As AnimationStorageRec
Public ScrollText As ScrollTextRec
Public ShopTrade As ShopTradeRec

Public Player() As PlayerRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public Sign() As SignRec
Public Animation() As AnimationRec

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

Type ShopTradeRec
    BuyItem As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Type AnimationRec
    Name As String * 10
    Width As Long
    Height As Long
    Delay As Long
    Pic As Long
End Type

Type AnimationStorageRec
    Active As Boolean
    Timer As Currency
    DelayTime As Long
    X As Long
    Y As Long
    Picture As Long
    Frame As Long
    Key As Long
End Type

Type Sprite_SizeRec
    SizeX As Byte
    SizeY As Byte
End Type

Type ConfigRec
    Password As String * 10
    IP As String * 15
    Port As Integer
End Type

Type GameConfigRec
    Game_Name As String * 30
    Website As String * 30
    Sprite_Offset As Byte
    WalkFrame() As Byte
    AttackFrame() As Byte
    Total_WalkFrames As Byte
    Total_AttackFrames As Byte
    WalkAnim_Speed As Integer
    StandFrame As Byte
End Type

Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Class As Byte
    Sprite As Integer
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    
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
    
    ' Client use only
    MaxHP As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Currency
    CastTimer(1 To MAX_PLAYER_SPELLS) As Currency
    MapGetTimer As Currency
    CastedSpell As Byte
    WalkTimer As Currency
    WalkAnim As Long
    GuildName As String
    GuildRank As Long
End Type

Type TileRec
    Layer(0 To 3) As Integer
    LayerSet(0 To 3) As Integer
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

Type ClassRec
    Name As String * NAME_LENGTH
    Sprite As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Byte
    
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
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
    
    AnimFrame As Long
    AnimTimer As Currency
    AnimItem As Long
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
    Num As Byte
    
    Target As Byte
    
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte
    
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Currency
    WalkTimer As Currency
    WalkAnim As Byte
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
    DoorOpen As Byte
End Type

Type SignRec
    Name As String * 10
    Section() As String
End Type

Type ScrollTextRec
    CurKey As Long
    KeyValue As Long
    Running As Boolean
    Text As String
    CurLetter As Long
End Type
