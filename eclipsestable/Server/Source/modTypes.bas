Attribute VB_Name = "modTypes"
Option Explicit

' ---NOTE to future developers!----------
' When loading, types ARE order-sensitive!
' This means do not change the order of variables in between
' versions, and add new variables to the end. This way, we can
' just load the old files! I learned that the hard way :D
' -Pickle

Type PlayerInvRec
    num As Integer
    Value As Long
    Dur As Integer
End Type

Type BankRec
    num As Integer
    Value As Long
    Dur As Integer
End Type

Type ElementRec
    Name As String * NAME_LENGTH
    Strong As Integer
    Weak As Integer
End Type

Public Type V000PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    GuildAccess As Byte
    Sex As Byte
    Class As Integer
    Sprite As Long
    LEVEL As Integer
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
    Speed As Long
    Magi As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Integer
    WeaponSlot As Integer
    HelmetSlot As Integer
    ShieldSlot As Integer
    LegsSlot As Integer
    RingSlot As Integer
    NecklaceSlot As Integer

    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Integer
    Bank(1 To MAX_BANK) As BankRec

    ' Position and movement
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte

    TargetNPC As Integer

    Head As Integer
    Body As Integer
    Leg As Integer

    PAPERDOLL As Byte

    MAXHP As Long
    MAXMP As Long
    MAXSP As Long
End Type

Public Type PlayerRec
    ' General
'090829 Scorpious2k
    Vflag As Byte       ' version flag - always > 127
    Ver As Byte
    SubVer As Byte
    Rel As Byte
'090829 End
    Name As String * NAME_LENGTH
    Guild As String
    GuildAccess As Byte
    Sex As Byte
    Class As Integer
    Sprite As Long
    LEVEL As Integer
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
    Speed As Long
    Magi As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Integer
    WeaponSlot As Integer
    HelmetSlot As Integer
    ShieldSlot As Integer
    LegsSlot As Integer
    RingSlot As Integer
    NecklaceSlot As Integer


    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Integer
    Bank(1 To MAX_BANK) As BankRec

    ' Position and movement
    Map As Integer
'090829    X As Byte
'090829    Y As Byte
    X As Integer
    Y As Integer
    Dir As Byte

    TargetNPC As Integer

    Head As Integer
    Body As Integer
    Leg As Integer

    PAPERDOLL As Byte

    MAXHP As Long
    MAXMP As Long
    MAXSP As Long
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
    InvVal As Long
End Type

Type PartyRec
    Leader As Byte
    Member() As Byte
    ShareExp As Boolean
End Type

Public Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    Email As String

    ' Some error here that needs to be fixed. [Mellowz]
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
    InParty As Boolean
    TargetType As Byte
    Target As Byte
    CastedSpell As Byte

    SpellTime As Long
    SpellVar As Long
    SpellDone As Long
    SpellNum As Long

    PartyStarter As Byte
    GettingMap As Byte
    Party As PartyRec
    InvitedBy As Byte

    Emoticon As Long

    InTrade As Boolean
    TradePlayer As Long
    TradeOk As Byte
    TradeItemMax As Byte
    TradeItemMax2 As Byte
    Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

    InChat As Byte
    ChatPlayer As Long

    Mute As Boolean
    Locked As Boolean
    LockedSpells As Boolean
    LockedItems As Boolean
    LockedAttack As Boolean
    TargetNPC As Long

    Pet As Long
    HookShotX As Byte
    HookShotY As Byte

    ' MENUS
    CustomMsg As String
    CustomTitle As String
End Type

Type TileRec
    Ground As Long
    Mask As Long
    Anim As Long
    Mask2 As Long
    M2Anim As Long
    Fringe As Long
    FAnim As Long
    Fringe2 As Long
    F2Anim As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    String2 As String
    String3 As String
    Light As Long
    GroundSet As Byte
    MaskSet As Byte
    AnimSet As Byte
    Mask2Set As Byte
    M2AnimSet As Byte
    FringeSet As Byte
    FAnimSet As Byte
    Fringe2Set As Byte
    F2AnimSet As Byte
End Type

Type MapRec
    Name As String * 20
    Revision As Integer
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    music As String
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Shop As Integer
    Indoors As Byte
    Tile() As TileRec
    NPC(1 To 15) As Integer
    SpawnX(1 To 15) As Byte
    SpawnY(1 To 15) As Byte
    Owner As String
    Scrolling As Byte
    Weather As Integer
End Type

Type ClassRec
    Name As String * NAME_LENGTH

    AdvanceFrom As Long
    LevelReq As Long
    Type As Long
    Locked As Long

    MaleSprite As Long
    FemaleSprite As Long

    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long

    Map As Long
    X As Byte
    Y As Byte

    ' Description
    Desc As String
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 150

    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    StrReq As Long
    DefReq As Long
    SpeedReq As Long
    MagicReq As Long
    ClassReq As Long
    AccessReq As Byte

    addHP As Long
    addMP As Long
    addSP As Long
    AddStr As Long
    AddDef As Long
    AddMagi As Long
    AddSpeed As Long
    AddEXP As Long
    AttackSpeed As Long
    Price As Long
    Stackable As Byte
    Bound As Byte

    ' Moved back to bottom... I suck :P -Pickle
    TwoHanded As Long
End Type

Type MapItemRec
    num As Long
    Value As Long
    Dur As Long

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

    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte

    STR  As Long
    DEF As Long
    Speed As Long
    Magi As Long
    Big As Long
    MAXHP As Long
    Exp As Long
    SpawnTime As Long

    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec

    Element As Long

    SPRITESIZE As Byte
End Type

Type MapNpcRec
    num As Long

    Target As Long

    HP As Long
    MP As Long
    SP As Long

    X As Byte
    Y As Byte
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

Type ShopItemRec
    ItemNum As Long
    Price As Long
    Amount As Long
End Type


Type ShopRec
    Name As String * NAME_LENGTH
    FixesItems As Byte
    BuysItems As Byte
    ShowInfo As Byte
    ShopItem(1 To MAX_SHOP_ITEMS) As ShopItemRec
    CurrencyItem As Integer
End Type

Type SpellRec
    Name As String * NAME_LENGTH
    ClassReq As Long
    LevelReq As Long
    MPCost As Long
    Sound As Long
    Type As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Range As Byte

    SpellAnim As Long
    SpellTime As Long
    SpellDone As Long

    AE As Long
    Big As Long

    Element As Long
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

Type ArrowRec
    Name As String
    Pic As Long
    Range As Byte
    Amount As Integer
End Type

Type StatRec
    LEVEL As Long
    STR As Long
    DEF As Long
    Magi As Long
    Speed As Long
End Type
                                
