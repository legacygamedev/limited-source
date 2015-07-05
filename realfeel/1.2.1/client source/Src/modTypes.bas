Attribute VB_Name = "modTypes"
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Trim$med module.
'****************************************************************

Option Explicit

  Type BITMAPINFOHEADER
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
  End Type
  
  Type BITMAPFILEHEADER
    bfType            As Integer
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOffBits         As Long
  End Type
  
  Type BITMAPINFO
    Width             As Long
    Height            As Long
  End Type

Type PlayerInvRec
    num As Byte
    Value As Long
    Dur As Integer
End Type

Type BankItemRec
    num As Long
    Value As Long
    Dur As Long
End Type

Type PlayerDyingRec
    DeathMap As Long 'Keep map if packet arrives early
    DeathX As Byte 'Keep X if packet arrives early
    DeathY As Byte 'Keep Y if packet arrives early
    AnimCount As Byte 'Number of times the text has been drawn
    Display As Byte 'Uses true and false
End Type

Type PlayerDamageRec
    Value As Long 'Damage
    TextX As Long 'Text X Position
    TextY As Long 'Text Y Position
    AnimCount As Byte 'Number of times the text has been drawn
    Display As Byte 'Uses true and false
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
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
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    BankInv(1 To MAX_BANK_ITEMS) As BankItemRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
       
    ' Position
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte
    
    ' Friends
    Friends(1 To MAX_FRIENDS) As String
    
    ' Tracker string
    Tracker As String
    
    ' Tint Variables, not sure how to save them as a hex.
    TintR As Byte
    TintG As Byte
    TintB As Byte
    
    ' Client use only
    MaxHP As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
    
    ' Client use only
    Death As PlayerDyingRec
    Damage(1 To 5) As PlayerDamageRec
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
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    Sprite As Integer
    
    STR As Byte
    DEF As Byte
    Speed As Byte
    MAGI As Byte
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
    
    Map As Long
    X As Long
    Y As Long
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
    num As Byte
    Value As Long
    Dur As Integer
    
    X As Byte
    Y As Byte
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
    STR  As Long
    DEF As Long
    Speed As Long
    MAGI As Long
    EXP As Long
    
    ' Tint Variables, not sure how to save them as a hex.
    TintR As Byte
    TintG As Byte
    TintB As Byte
    
    'To see if the npc shall flee when its HP is low
    '-smchronos
    Fear As Boolean
    
    'To see if the npc can cast spells
    '-smchronos
    MagicUser As Boolean
    
    'The total number of magic spells
    NpcSpell(1 To MAX_NPC_SPELLS) As Integer
End Type

Type MapNpcDyingRec
    DeathX As Byte 'Keep X if packet arrives early
    DeathY As Byte 'Keep Y if packet arrives early
    DeathDir As Byte 'Keep Dir if packet arrives early
    AnimCount As Byte 'Number of times the text has been drawn
    Display As Byte 'Uses true and false
End Type

Type MapNpcDamageRec
    Value As Long 'Damage
    TextX As Long 'Text X Position
    TextY As Long 'Text Y Position
    AnimCount As Byte 'Number of times the text has been drawn
    Display As Byte 'Uses true and false
End Type

Type MapNpcRec
    num As Byte
    
    Target As Byte
    
    HP As Long
    MaxHP As Long
    MP As Long
    SP As Long
        
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte

    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    
    ' Client use only
    Death As MapNpcDyingRec
    Damage(1 To 5) As MapNpcDamageRec
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
    JoinSay As String * 100
    LeaveSay As String * 100
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
    DoorOpen As Byte
End Type
