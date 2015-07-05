Attribute VB_Name = "modTypes"
Option Explicit

Type ChatBubble
    Text As String
    Created As Long
End Type

Type ScriptBubble
    Text As String
    Created As Long
    Map As Long
    X As Long
    Y As Long
    Colour As Long
End Type

Type BankRec
    Num As Long
    value As Long
    Dur As Long
End Type

Type PlayerInvRec
    Num As Long
    value As Long
    Dur As Long
End Type

Type ElementRec
    name As String * NAME_LENGTH
    Strong As Long
    Weak As Long
End Type

Type SpellAnimRec
    CastedSpell As Byte

    SpellTime As Long
    SpellVar As Long
    SpellDone As Long

    Target As Long
    TargetType As Long
End Type

Type ScriptSpellAnimRec
    CastedSpell As Byte

    SpellTime As Long
    SpellVar As Long
    SpellDone As Long

    SpellNum As Long
    X As Long
    Y As Long
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
    ArrowAmount As Long
End Type

Type PartyRec
    Leader As Byte
    Member() As Byte
    ShareExp As Boolean
End Type

Type PlayerRec
    ' General
    name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Class As Long
    Sprite As Long
    Level As Long
    Exp As Long
    Access As Byte
    PK As Byte
    input As Byte
    iso As Byte
    Party As PartyRec
    Step As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long

    ' Stats
    STR As Long
    DEF As Long
    speed As Long
    MAGI As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    LegsSlot As Long
    RingSlot As Long
    NecklaceSlot As Long

    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    Bank(1 To MAX_BANK) As BankRec

    ' Position
    Map As Long
    X As Integer
    Y As Integer
    Dir As Byte

    ' Client use only
    MaxHp As Long
    MaxMP As Long
    MaxSP As Long
    xOffset As Integer
    yOffset As Integer
    MovingH As Integer
    MovingV As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte

    SpellNum As Long
    SpellAnim() As SpellAnimRec

    EmoticonNum As Long
    EmoticonTime As Long
    EmoticonVar As Long

    LevelUp As Long
    LevelUpT As Long

    Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec

    SkilLvl() As Long
    SkilExp() As Long

    Armor As Long
    Helmet As Long
    Shield As Long
    Weapon As Long
    legs As Long
    Ring As Long
    Necklace As Long
    color As Long
    pet As Long

    head As Long
    body As Long
    leg As Long

    HookShotX As Long
    HookShotY As Long
    HookShotSucces As Long
    HookShotAnim As Long
    HookShotTime As Long
    HookShotToX As Long
    HookShotToY As Long
    HookShotDir As Long

    paperdoll As Byte
End Type

Type TileRec
    Ground As Long
    mask As Long
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
    light As Long
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
    name As String * 20
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
    Npc(1 To 15) As Integer
    SpawnX(1 To 15) As Byte
    SpawnY(1 To 15) As Byte
    owner As String
    scrolling As Byte
    Weather As Integer
End Type

Type ClassRec
    name As String * NAME_LENGTH
    MaleSprite As Long
    FemaleSprite As Long
    
    Locked As Long
    
    STR As Long
    DEF As Long
    speed As Long
    MAGI As Long
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
    
    ' Description
    desc As String
End Type

Type ItemRec
    name As String * NAME_LENGTH
    desc As String * 150
    
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
    
    AddHP As Long
    AddMP As Long
    AddSP As Long
    AddSTR As Long
    AddDEF As Long
    AddMAGI As Long
    AddSpeed As Long
    AddEXP As Long
    AttackSpeed As Long
    Price As Long
    
    Stackable As Long
    Bound As Long
End Type
    
Type MapItemRec
    Num As Long
    value As Long
    Dur As Long
    
    X As Byte
    Y As Byte
End Type

Type NPCEditorRec
    ItemNum As Long
    ItemValue As Long
    chance As Long
End Type

Type NpcRec
    name As String * NAME_LENGTH
    AttackSay As String * 100
    
    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    SpriteSize As Long
    
    STR  As Long
    DEF As Long
    speed As Long
    MAGI As Long
    Big As Long
    MaxHp As Long
    Exp As Long
    SpawnTime As Long
    Spell As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
    
    Element As Long
End Type

Type MapNpcRec
    Num As Long
    
    Target As Long
    
    HP As Long
    MaxHp As Long
    MP As Long
    SP As Long
    
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    Big As Byte
    
    ' Client use only
    xOffset As Integer
    yOffset As Integer
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

Type ShopItemRec
    ItemNum As Long
    Price As Long
    Amount As Long
End Type

Type ShopRec
    name As String * NAME_LENGTH
    FixesItems As Byte
    BuysItems As Byte
    ShowInfo As Byte
    ShopItem(1 To MAX_SHOP_ITEMS) As ShopItemRec
    currencyItem As Integer
End Type

Type SpellRec
    name As String * NAME_LENGTH
    ClassReq As Long
    LevelReq As Long
    Sound As Long
    MPCost As Long
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
    reload As Long
End Type

Type TempTileRec
    DoorOpen As Byte
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
    InvVal As Long
End Type

Type EmoRec
    Pic As Long
    Command As String
End Type

Type DropRainRec
    X As Long
    Y As Long
    Randomized As Boolean
    speed As Byte
End Type

Type ItemTradeRec
    ItemGetNum As Long
    ItemGiveNum As Long
    ItemGetVal As Long
    ItemGiveVal As Long
End Type

Type TradeRec
    Items(1 To MAX_TRADES) As ItemTradeRec
    Selected As Long
    SelectedItem As Long
End Type

Type ArrowRec
    name As String
    Pic As Long
    Range As Byte
    Amount As Long
End Type

Type BattleMsgRec
    Msg As String
    index As Byte
    color As Byte
    time As Long
    Done As Byte
    Y As Long
End Type

Type ItemDurRec
    Item As Long
    Dur As Long
    Done As Byte
End Type
