Attribute VB_Name = "modTypes"
Option Explicit

Type AnimationRec
    Name            As String * NAME_LENGTH
    Animation       As Long
    AnimationFrames As Long
    AnimationSpeed  As Long
    AnimationSize   As Long
    AnimationLayer  As Long
End Type

Type PositionRec
    Map As Long
    X   As Byte
    Y   As Byte
End Type

Type StatusRec
    SpellNum    As Long
    TickCount   As Long
    TickUpdate  As Long
    Caster      As String * NAME_LENGTH ' Will hold the name of the caster - used for if the spell kills the player
End Type
    
Type PlayerInvRec
    Num     As Byte
    Value   As Long
    Bound As Boolean
End Type

Type PlayerSpellRec
    SpellNum    As Long
    Cooldown    As Long
End Type

Type QuestProgressUDT
    QuestNum As Long
    Progress(1 To MAX_QUEST_NEEDS) As Long
End Type

Type PlayerRec
    ' General
    Name    As String * NAME_LENGTH
    Sex     As Byte
    Class   As Byte
    Sprite  As Integer
    Level   As Byte
    Exp     As Long
    Access  As Byte
    PK      As Byte
    Guild   As Byte
    GuildRank As Byte
    GuildName As String * NAME_LENGTH
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count) As Long
    Points As Long
    
    ' Worn equipment
    Equipment(1 To Slots.Slot_Count) As Long
    
    ' Gold amount
    Gold As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As PlayerSpellRec
    
    ' Current Direction
    Dir As Byte
    
    ' Position
    Position As PositionRec
    Bound As PositionRec
    
    ' Used for buffs/debuffs/dots/hots
    Status(1 To MAX_STATUS) As StatusRec
    
    ' For Death
    IsDead   As Boolean
    IsDeadTimer As Long
    
    ' For Quests
    ActiveQuestCount As Long
    CompletedQuests(1 To MAX_QUESTS) As Long
    QuestProgress(1 To MAX_PLAYER_QUESTS) As QuestProgressUDT
End Type

Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
       
    ' The actual player data
    Char As PlayerRec
    
    ' None saved local vars
    Buffer As clsBuffer
    
    CharNum As Byte
    InGame As Boolean
    AttackTimer As Long
    
    ' Network Data
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    PacketInIndex As Byte   ' Holds the index of what packetkey for incoming packets
    PacketOutIndex As Byte  ' Holds the index of what packetkey for outgoing packets
    
    ' Target Data
    Target As Byte
    TargetType As Byte
    
    ' Spell Casting Data
    CastingSpell As Long    ' Holds what spell they are casting
    CastTime As Long        ' Holds the timer for when a spell is ready to be cast
    CastTarget As Long      ' Holds the target for the spell being cast
    CastTargetType As Long  ' Holds the targettype for the spell being cast
    
    ' Party Data
    InParty As Boolean
    PartyStarter As Byte
    PartyIndex As Long      ' Holds what Party ID you were either invited to or in
    PartyInvitedBy As String    ' Holds the name of the person who invited you
    
    GettingMap As Byte

    GuildInvite As Byte
    GuildInviter As Byte
    
    ModStat(1 To Stats.Stat_Count) As Long      ' Holds our modstat
    ModVital(1 To Vitals.Vital_Count) As Long   ' Hold our modvital
    
    Revivable As Long       ' (Spellnum) This isn't saved because there's no need to persist a revive spell through logout/login
    
    LastUpdateVitals As Long
    LastUpdateSave As Long
End Type

Type MobsRec
    NpcCount As Long    ' Count of npcs - used in total map_npc_count
    Npc() As Long       ' List of all npcs for this grouping
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
End Type

Type MapRec
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
    TileSet As Byte
    MaxX As Byte
    MaxY As Byte
    Tile() As TileRec
    Mobs(1 To MAX_MOBS) As MobsRec
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    
    MaleSprite As String
    FemaleSprite As String
    
    Vital(1 To Vitals.Vital_Count) As Long
    Stat(1 To Stats.Stat_Count) As Long
    BaseDodge As Long
    BaseCrit As Long
    BaseBlock As Long
    Threat As Byte
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    Pic As Integer
    
    LevelReq As Long
    ClassReq As Integer ' Flags for each class
    StatReq(1 To Stats.Stat_Count) As Byte
    
    Type As Byte
    Rarity As Byte
    Bound As Long
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    
    ModVital(1 To Vitals.Vital_Count) As Long
    ModStat(1 To Stats.Stat_Count) As Long
    
    Stack As Byte
    StackMax As Long
End Type

Type NpcDropItemRec
    Item As Long
    ItemValue As Long
    Chance As Byte
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 255
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    MovementSpeed As Byte
    MovementFrequency As Byte
    
    Drop(1 To 4) As NpcDropItemRec
    
    Stat(1 To Stats.Stat_Count) As Long
    MaxHP As Long
    MaxEXP As Long
    Level As Long
End Type

Type MapNpcRec
    Num As Long
    Target As Long
    Vital(1 To Vitals.Vital_Count) As Long
        
    X As Byte
    Y As Byte
    Dir As Byte
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    
    Status(1 To MAX_STATUS) As StatusRec
    
    ' Damage dictionary - Used for exp % and threat
    Damage As Dictionary
    
    ' Used for running away
    StepsTaken As Long
    IsRunning As Boolean
    
    ' This will be used for determining
    AmountHealed As Long
    
    ' This will be used to restore the NPC back to full
    LastDamageTaken As Long
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 500
    Type As Long                ' What type
    BindPoint As PositionRec                    ' Inns - Used for home point setting
    TradeItem(1 To MAX_TRADES) As TradeItemRec  ' Shops
End Type

Type SpellRec
    Name As String * NAME_LENGTH
    
    Type As Byte
    Range As Byte
    
    LevelReq As Byte
    ClassReq As Long ' Flags for each class
    VitalReq(1 To Vitals.Vital_Count) As Long
    
    TargetFlags As Long ' Refer to the "Targets" enum for flag bits
    
    CastTime As Long
    Cooldown As Long
    
    ModVital(1 To Vitals.Vital_Count) As Long
    ModStat(1 To Stats.Stat_Count) As Long
    
    TickCount As Long   ' overtime spells-how many ticks         buffs-always set to 1
    TickUpdate As Long  ' overtime spells-time in between ticks  buffs-total length of buff
    
    Animation As Long       ' What animation to play
End Type

Type GuildRec
    Guild As Byte
    GuildName As String * NAME_LENGTH
    GuildAbbreviation As String * NAME_LENGTH
    GMOTD As String * 255
    Owner As String * NAME_LENGTH
    Rank(1 To MAX_GUILD_RANKS) As String * NAME_LENGTH
End Type

Type EmoRec
    Pic As Long
    Command As String * NAME_LENGTH
End Type

Type Cache
    Data() As Byte
End Type

Type NpcGroupRec
    Npc As Long
    MobGroup As Long
End Type

Type TempTileRec
    DoorOpen() As Boolean
    DoorTimer() As Long
End Type

Type MapItemRec
    Num As Long
    Value As Long
    X As Byte
    Y As Byte
    DropTime As Long
End Type

Type MapDataRec
    MapPlayersCount As Long
    MapPlayers() As Long
    NpcCount As Long
    Npc() As NpcGroupRec
    MapNpc() As MapNpcRec
    MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
    TempTile As TempTileRec
End Type

' Used to send the following all at once instead of a ton of packets
Public MapCache(1 To MAX_MAPS) As Cache
Public ItemsCache() As Byte
Public NpcsCache() As Byte
Public EmoticonsCache() As Byte
Public ShopsCache() As Byte
Public SpellsCache() As Byte
Public AnimationsCache() As Byte
Public QuestsCache() As Byte

Public Map(1 To MAX_MAPS) As MapRec
Public MapData(1 To MAX_MAPS) As MapDataRec
Public Player(1 To MAX_PLAYERS) As AccountRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Guild(1 To MAX_GUILDS) As GuildRec
Public Emoticons(1 To MAX_EMOTICONS) As EmoRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
