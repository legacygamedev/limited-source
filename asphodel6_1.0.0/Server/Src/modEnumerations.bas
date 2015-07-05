Attribute VB_Name = "modEnumerations"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

' Window state constants
Public Enum Window_State
    Main_Menu = 0
    Login
    New_Account
    New_Char
    Chars
    Credits
    Main_Game
    Count
End Enum

' Guild rank constants
Public Enum Guild_Rank
    Rank1 = 1
    Rank2
    Rank3
    Rank4
End Enum

' Item requirement constants
Public Enum Item_Requires
    Strength_ = 0
    Defense_
    Speed_
    Magic_
    Class_
    Level_
    Access_
    Count
End Enum

' Player stat windows constants
Public Enum Player_StatWindow
    Points_ = 0
End Enum

' Constants for admin panel actions
Public Enum ACP_Action
    LevelSelf = 0
    LevelTarget
    SetTargetSprite
    CheckAccount
    CheckInventory
    GiveSelfPK
    GiveTargetPK
    MutePlayer
End Enum

' Constants for tile layers
Public Enum Tile_Layer
    Ground = 0
    Mask
    Anim
    Fringe
End Enum

' NPC reflection constants
Public Enum NPC_Reflection
    Magic_ = 0
    Melee_
End Enum

' NPC behavior constants
Public Enum NPC_Behavior
    AttackOnSight = 0
    AttackWhenAttacked
    Friendly
    ShopKeeper
    Guard
End Enum

' Game editor types
Public Enum GameEditor
    Item_ = 1
    NPC_
    Spell_
    Shop_
    Sign_
    Anim_
End Enum

' Movement type constants
Public Enum MovementType
    Walking = 1
    Running
End Enum

' Spell type constants
Public Enum Spell_Type
    AddHP_ = 0
    AddMP_
    AddSP_
    SubHP_
    SubMP_
    SubSP_
    GiveItem_
End Enum

' Tile type constants
Public Enum Tile_Type
    None_ = 0
    Blocked_
    Warp_
    Item_
    NpcAvoid_
    Key_
    KeyOpen_
    Shop_
    Sign_
    Guild_
    Heal_
    Damage_
End Enum

' Staff level constants
Public Enum StaffType
    Monitor = 1
    Mapper
    Developer
    Creator
End Enum

' Gender constants
Public Enum GenderType
    Male_ = 0
    Female_
End Enum

' NPC editor sound constants
Public Enum NpcSound
    Attack_ = 0
    Spawn_
    Death_
End Enum

' Target type constants
Public Enum E_Target
    None = 0
    Player_
    NPC_
End Enum

' Chat type constants
Public Enum E_ChatType
    MapMsg_ = 0
    EmoteMsg_
    BroadcastMsg_
    GlobalMsg_
    AdminMsg_
    PrivateMsg_
End Enum

' Color codes
Public Enum Color
    Black = 0
    Blue
    Green
    Cyan
    Red
    Magenta
    Brown
    Grey
    DarkGrey
    BrightBlue
    BrightGreen
    BrightCyan
    BrightRed
    Pink
    Yellow
    White
End Enum

' Item type constants
Public Enum ItemType
    None = 0
    Weapon_
    Armor_
    Helmet_
    Shield_
    Potion
    Key
    Currency_
    Spell_
End Enum

' Direction constants
Public Enum E_Direction
    Up_ = 0
    Down_
    Left_
    Right_
End Enum

' Packets sent by the client
Public Enum ClientPackets
    CGetClasses = 1
    CNewAccount
    CLogin
    CAddChar
    CDelChar
    CUseChar
    CMessage
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CPlayerInfoRequest
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CGetStats
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CEditItem
    CSaveItem
    CRequestEditNpc
    CEditNpc
    CSaveNpc
    CRequestEditShop
    CEditShop
    CSaveShop
    CRequestEditSpell
    CEditSpell
    CSaveSpell
    CDelete
    CSetAccess
    CWhosOnline
    CSetMotd
    CTradeRequest
    CFixItem
    CSearch
    CParty
    CJoinParty
    CLeaveParty
    CSpells
    CCast
    CQuit
    CConfigPass
    CACPAction
    CRCWarp
    CRequestEditSign
    CEditSign
    CSaveSign
    CPressReturn
    CGuildCreation
    CGuildDisband
    CGuildInvite
    CInviteResponse
    CGuildPromoteDemote
    CRequestEditAnim
    CSaveAnim
    CEditAnim
    CPing
    CLogout
    CSellItem
End Enum

' Packets recieved by the client
Public Enum ServerPackets
    SAlertMsg = 1
    SAllChars
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    SPlayerSp
    SPlayerStats
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SMessage
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SEditItem
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SEditNpc
    SMapKey
    SEditMap
    SShopEditor
    SUpdateShop
    SEditShop
    SSpellEditor
    SUpdateSpell
    SEditSpell
    STrade
    SSpells
    SLeft
    SConfigPass
    SGameOptions
    SAnimation
    SSoundPlay
    SPlayerPoints
    SPlayerLevel
    SClassName
    SPlayerStatBuffs
    SSignEditor
    SEditSign
    SUpdateSign
    SScrollingText
    SGuildCreation
    SPlayerGuild
    SGuildInvite
    SAnimEditor
    SEditAnim
    SUpdateAnim
    SPing
    SNpcHP
    SNormalMsg
    SCastSuccess
    SExpUpdate
End Enum

' ****************
' ** Statistics **
' ****************

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Defense
    SPEED
    Magic
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    SP
    ' Mak sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    Helmet
    Shield
    ' Mak sure Equipment_Count is below everything else
    Equipment_Count
End Enum
