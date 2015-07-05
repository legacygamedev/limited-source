Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec
Public TempMap As MapRec
Public Bank As BankRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Item() As ItemRec
Public NPC() As NPCRec
Public MapItem(MAX_MAP_ITEMS) As MapItemRec
Public MapNPC(MAX_MAP_NPCS) As MapNPCRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Resource() As ResourceRec
Public Animation() As AnimationRec
Public events(1 To MAX_EVENTS) As EventWrapperRec
Public Ban() As BanRec
Public title() As TitleRec
Public Moral() As MoralRec
Public Class() As ClassRec
Public Emoticon() As EmoticonRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public WeatherParticle(1 To MAX_WEATHER_PARTICLES) As WeatherParticleRec
Public Autotile() As AutotileRec
Public MapSounds() As SoundsRec
Public MapSoundCount As Long
Public Sounds(1 To MAX_SOUNDS) As Long
Public SoundsXY(1 To MAX_SOUNDS) As XYRec

' Battle Music
Public CacheNPCTargets() As Byte
Public ActiveNPCTarget As Byte
Public InitBattleMusic As Boolean

' Logs
Public Log As LogRec

' Options
Public Options As OptionsRec

' Client-side stuff
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public MenuButton(1 To MAX_MENUBUTTONS) As ButtonRec
Public MainButton(1 To MAX_MAINBUTTONS) As ButtonRec
Public Party As PartyRec

Public Type SoundsRec
    X As Long
    Y As Long
    handle As Long
    InUse As Boolean
    channel As Long
End Type

Private Type XYRec
    X As Double
    Y As Double
End Type

' Type recs
Private Type OptionsRec
    SaveUsername As Byte
    UserName As String * NAME_LENGTH
    IP As String
    Port As Long
    MenuMusic As String
    Music As Byte
    Sound As Byte
    WASD As Byte
    Levels As Byte
    Guilds As Byte
    PlayerVitals As Byte
    NPCVitals As Byte
    Titles As Byte
    BattleMusic As Byte
    Mouse As Byte
    Debug As Byte
    SwearFilter As Byte
    Weather As Byte
    Autotile As Byte
    Blood As Byte
    MusicVolume As Double
    SoundVolume As Double
    ResolutionWidth As Long
    ResolutionHeight As Long
End Type

Public Type PartyRec
    num As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerItemRec
    num As Byte
    Value As Long
    Durability As Integer
    Bind As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerItemRec
End Type

Private Type SpellAnim
    SpellNum As Integer
    Timer As Long
    FramePointer As Long
End Type

Public Type BuffRec
    ID As Long
    Behavior As Long
    Vital As Long
    Timer As Long
End Type

Type FriendsRec
    Name As String * NAME_LENGTH
End Type

Type FoesRec
    Name As String * NAME_LENGTH
End Type

Public Type SkillRec
    Level As Byte
    exp As Long
End Type

Private Type QuestAmountRec
    ID() As Integer
End Type

Public Type PlayerRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    
    ' Face
    Face As Integer
    
    ' General
    Name As String * NAME_LENGTH
    Gender As Byte
    Class As Byte
    Sprite As Integer
    Level As Byte
    exp As Long
    Access As Byte
    PK As Byte
    Status As String * NAME_LENGTH
    
    ' Position
    Map As Integer
    X As Long
    Y As Long
    Dir As Byte
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Integer
    Points As Integer
    
    ' Amount of titles
    AmountOfTitles As Byte
    
    ' Titles
    title() As Byte
    
    ' Current title
    CurTitle As Byte
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As PlayerItemRec
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerItemRec
    
    ' Buffs
    Buff(1 To MAX_BUFFS) As BuffRec
    
    ' Guild
    Guild As String * NAME_LENGTH
    GuildAcc As Byte
    
    ' Friends
    Friends(1 To MAX_PEOPLE) As FriendsRec
    
    ' Foes
    Foes(1 To MAX_PEOPLE) As FoesRec
    
    ' Skill
    Skills(1 To Skill_Count - 1) As SkillRec
    
    ' Questing
    QuestCLI() As Long
    QuestTask() As Long
    QuestAmount() As QuestAmountRec
    QuestCompleted() As Boolean
End Type

' Character editor
Public Type PlayerEditableRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
  
    ' General
    Name As String * NAME_LENGTH
    Gender As Byte
    Class As Byte
    Sprite As Integer
    Level As Byte
    exp As Long
    Access As Byte

    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Integer
    Points As Integer
End Type

Private Type TempPlayerRec
    xOffset As Double
    yOffset As Double
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    AnimTimer As Long
    Anim As Long
    EmoticonNum As Long
    EmoticonTimer As Long
    EventTimer As Long
    PvPTimer As Long
End Type

Private Type TileDataRec
    X As Byte
    Y As Byte
    Tileset As Byte
End Type

Public Type ConditionalBranchRec
    Condition As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    CommandList As Long
    ElseCommandList As Long
End Type

Public Type MoveRouteRec
    Index As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
End Type

Public Type EventCommandRec
    Index As Long
    Text1 As String
    Text2 As String
    Text3 As String
    Text4 As String
    Text5 As String
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
    ConditionalBranch As ConditionalBranchRec
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
End Type

Public Type CommandListRec
    CommandCount As Long
    ParentList As Long
    Commands() As EventCommandRec
End Type

Public Type EventPageRec
    ' These are condition variables that decide if the event even appears to the player
    chkVariable As Long
    VariableIndex As Long
    VariableCondition As Long
    VariableCompare As Long
    
    chkSwitch As Long
    SwitchIndex As Long
    SwitchCompare As Long
    
    chkHasItem As Long
    HasItemIndex As Long
    
    chkSelfSwitch As Long
    SelfSwitchIndex As Long
    SelfSwitchCompare As Long
    
    ' Handles the event sprite
    GraphicType As Byte
    Graphic As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    
    ' Handles movement
    MoveType As Byte
    MoveSpeed As Byte
    MoveFreq As Byte
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    IgnoreMoveRoute As Long
    RepeatMoveRoute As Long
    
    ' Guidelines for the event
    WalkAnim As Byte
    DirFix As Byte
    WalkThrough As Byte
    ShowName As Byte
    
    ' Trigger for the event
    Trigger As Byte
    
    ' Commands for the event
    CommandListCount As Long
    CommandList() As CommandListRec
    
    Position As Byte
    
    ' Client needed only
    X As Long
    Y As Long
End Type

Public Type EventRec
    Name As String * NAME_LENGTH
    Global As Long
    PageCount As Long
    Pages() As EventPageRec
    X As Long
    Y As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As String
    DirBlock As Byte
End Type

Private Type MapEventRec
    Name As String
    Dir As Long
    X As Long
    Y As Long
    GraphicType As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    GraphicNum As Long
    Moving As Long
    MovementSpeed As Long
    Position As Long
    xOffset As Long
    yOffset As Long
    Step As Long
    Visible As Long
    WalkAnim As Long
    DirFix As Long
    ShowDir As Long
    WalkThrough As Long
    ShowName As Long
    Trigger As Byte
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * FILE_LENGTH
    BGS As String * FILE_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    Weather As Long
    WeatherIntensity As Long
    
    Fog As Long
    FogSpeed As Long
    FogOpacity As Long
    
    Panorama As Long
    
    Red As Long
    Green As Long
    Blue As Long
    Alpha As Long
    
    MaxX As Byte
    MaxY As Byte
    
    NPC_HighIndex As Byte
    
    Tile() As TileRec
    NPC(MAX_MAP_NPCS) As Long
    NPCSpawnType(MAX_MAP_NPCS) As Long
    EventCount As Long
    events() As EventRec
    
    ' Client side only
    CurrentEvents As Long
    MapEvents() As MapEventRec
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Integer
    MaleSprite As Integer
    FemaleSprite As Integer
    
    StartItem(1 To MAX_INV) As Long
    StartItemValue(1 To MAX_INV) As Long
    StartSpell(1 To MAX_PLAYER_SPELLS) As Long
    
    Locked As Byte
    
    ' Faces
    MaleFace As Integer
    FemaleFace As Integer
    
    ' Color
    Color As Long
    
    ' Start position
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte
    
    ' Combat tree
    CombatTree As Byte
    
    Animated As Byte
End Type

Public Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 256
    Sound As String * FILE_LENGTH
    
    Pic As Integer
    Type As Byte
    
    EquipSlot As Byte
    
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    
    ClassReq As Byte
    AccessReq As Byte
    LevelReq As Byte
    GenderReq As Byte
    ProficiencyReq As Byte
    
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Integer
    Rarity As Byte
    WeaponSpeed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Integer
    Animation As Long
    Paperdoll As Long
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    InstaCast As Byte
    ChanceModifier As Byte
    IsReusable As Boolean
    Tool As Integer
    HoT As Byte
    TwoHanded As Byte
    stackable As Byte
    Indestructable As Byte
    SkillReq As Byte
    ToolRequired As Integer
    Skill As Byte
    SkillExp As Integer
    SkillLevelReq As Byte
End Type

Private Type MapItemRec
    PlayerName As String * NAME_LENGTH
    num As Long
    Value As Long
    Durability As Integer
    Frame As Byte
    X As Long
    Y As Long
End Type

Private Type NPCRec
    Name As String * NAME_LENGTH
    title As String * NAME_LENGTH
    Music As String * FILE_LENGTH
    Sound As String * FILE_LENGTH
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Byte
    DropValue(1 To MAX_NPC_DROPS) As Integer
    Damage As Long
    Stat(1 To Stats.Stat_Count - 1) As Integer
    HP As Long
    MP As Long
    exp As Long
    Animation As Long
    Level As Byte
    Spell(1 To MAX_NPC_SPELLS) As Integer
    Faction As Byte
    AttackSay As String * 100
    FactionThreat As Boolean
    SwitchNum As Long
    VariableNum As Long
    SwitchVal As Byte
    VariableVal As Long
    AddToVariable As Byte
    ShowQuestCompleteIcon As Long
    DropRandom(1 To MAX_NPC_DROPS) As Byte
    Animated As Byte
End Type

Private Type MapNPCRec
    num As Byte
    Target As Byte
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    X As Byte
    Y As Byte
    Dir As Byte
    
    ' Client use only
    xOffset As Long
    yOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    AnimTimer As Long
    Anim As Long
    SpellBufferTimer As Long
    SpellBuffer As Long
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
    CostItem2 As Long
    CostValue2 As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Integer
    SellRate As Integer
    TradeItem(1 To MAX_TRADES) As TradeItemRec
    CanFix As Byte
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 256
    Sound As String * FILE_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Byte
    AccessReq As Byte
    ClassReq As Byte
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    AoE As Long
    IsAoe As Boolean
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    Sprite As Integer
    WeaponDamage As Boolean
    CastRequired As Integer
    NewSpell As Integer
End Type

Public Type MapResourceRec
    X As Long
    Y As Long
    ResourceState As Byte
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    FailMessage As String * NAME_LENGTH
    Sound As String * FILE_LENGTH

    Skill As Byte
    exp As Integer
    ResourceImage As Byte
    ExhaustedImage As Byte
    ItemReward As Long
    ToolRequired As Long
    Reward_Min As Byte
    Reward_Max As Byte
    RespawnTime As Long
    Animation As Long
    LowChance As Byte
    HighChance As Byte
    LevelReq As Byte
End Type

Private Type ActionMsgRec
    Message As String
    Timer As Long
    Type As Long
    Color As Long
    Scroll As Long
    X As Long
    Y As Long
    Alpha As Byte
    WaitTimer As Long
End Type

Private Type BloodRec
    Sprite As Long
    Timer As Long
    X As Long
    Y As Long
    Alpha As Byte
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * FILE_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    X As Byte
    Y As Byte
    
    ' Used for locking to players/npcs
    LockIndex As Long
    LockType As Byte
    
    ' Timing
    Timer(0 To 1) As Long
    
    ' Rendering check
    Used(0 To 1) As Boolean
    
    ' Counting the loop
    LoopIndex(0 To 1) As Long
    frameIndex(0 To 1) As Long
End Type

Public Type CameraRec
    Top As Long
    Bottom As Long
    Left As Long
    Right As Long
End Type

Private Type BanRec
    Date As String * NAME_LENGTH
    time As String * NAME_LENGTH
    PlayerName As String * NAME_LENGTH
    PlayerLogin As String * NAME_LENGTH
    IP As String * NAME_LENGTH
    HDSerial As String * NAME_LENGTH
    Reason As String * 100
    By As String * NAME_LENGTH
End Type

Public Type TitleRec
    Name As String * NAME_LENGTH
    Color As Byte
    LevelReq As Byte
    PKReq As Integer
    Desc As String * 100
End Type

Public Type MoralRec
    Name As String * NAME_LENGTH
    Color As Byte
    CanPK As Byte
    CanCast As Byte
    CanUseItem As Byte
    LoseExp As Byte
    DropItems As Byte
    CanPickupItem As Byte
    CanDropItem As Byte
    PlayerBlocked As Byte
End Type

Public Type HotbarRec
    Slot As Byte
    sType As Byte
End Type

Type DropRec
    X As Long
    Y As Long
    YSpeed As Long
    XSpeed As Long
    Init As Boolean
End Type

Public Type ButtonRec
    FileName As String
    State As Byte
End Type

Private Type EmoticonRec
    Command As String * NAME_LENGTH
    Pic As Long
End Type

Public Type LogRec
    Msg As String * 512
    file As String * NAME_LENGTH
End Type

Public Type EventListRec
    CommandList As Long
    CommandNum As Long
End Type

Public Type WeatherParticleRec
    Type As Long
    X As Long
    Y As Long
    Velocity As Long
    InUse As Long
End Type

' Auto tiles
Public Type PointRec
    X As Long
    Y As Long
End Type

Public Type QuarterTileRec
    QuarterTile(1 To 4) As PointRec
    RenderState As Byte
    srcX(1 To 4) As Long
    srcY(1 To 4) As Long
End Type

Public Type AutotileRec
    Layer(1 To MapLayer.Layer_Count - 1) As QuarterTileRec
End Type

Public Type ChatBubbleRec
    Msg As String
    Color As Long
    Target As Long
    TargetType As Byte
    Timer As Long
    active As Boolean
    Alpha As Byte
End Type

Public Type SubEventRec
    Type As EventType
    HasText As Boolean
    text() As String
    HasData As Boolean
    data() As Long
End Type

Public Type EventWrapperRec
    Name As String
    chkSwitch As Byte
    chkVariable As Byte
    chkHasItem As Byte
    
    SwitchIndex As Long
    SwitchCompare As Byte
    VariableIndex As Long
    VariableCompare As Byte
    VariableCondition As Long
    HasItemIndex As Long
    
    HasSubEvents As Boolean
    SubEvents() As SubEventRec
    
    Trigger As Byte
    WalkThrought As Byte
    Animated As Byte
    Graphic(0 To 2) As Long
End Type
