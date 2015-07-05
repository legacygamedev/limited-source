Attribute VB_Name = "modEnumerations"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
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
    AddHP = 0
    AddMP
    AddSP
    SubHP
    SubMP
    SubSP
    GiveItem
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

' Menu state constants
Public Enum Menu_State
    NewAccount_ = 0
    Login_
    NewChar_
    AddChar_
    DelChar_
    UseChar_
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
Public Enum ChatType
    MapMsg = 0
    EmoteMsg
    BroadcastMsg
    GlobalMsg
    AdminMsg
    PrivateMsg
End Enum

' Color codes
Public Enum Color
    Black = 0
    blue
    green
    Cyan
    red
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
    Count
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
    Speed
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
