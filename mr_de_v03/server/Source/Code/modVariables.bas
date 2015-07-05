Attribute VB_Name = "modVariables"
Option Explicit
'//////////////
'// Declares //
'//////////////

' General
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Sleep API - used to "sleep" the process and free the CPU usage
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

'///////////////
'// Constants //
'///////////////

' Winsock globals
Public Const GAME_PORT As Long = 4000

' General constants
Public Const GAME_NAME As String = "Mirage Realms Developer Editions"
Public Const GAME_WEBSITE As String = "www.mirage-realms.com"
Public Const MAX_PLAYERS As Byte = 100
Public Const MAX_ITEMS As Integer = 1000
Public Const MAX_NPCS As Integer = 1000
Public Const MAX_SHOPS As Integer = 1000
Public Const MAX_SPELLS As Integer = 1000
Public Const MAX_GUILDS As Byte = 255
Public Const MAX_GUILD_MEMBERS As Byte = 255
Public Const MAX_EMOTICONS As Byte = 255
Public Const MAX_ANIMATIONS As Byte = 255
Public Const MAX_MAP_ITEMS As Integer = 20
Public Const MAX_TRADES As Byte = 12
Public Const MAX_MOBS As Byte = 20

Public Const MAX_LEVEL As Byte = 100

Public Const MAX_INV As Byte = 16
Public Const MAX_PLAYER_SPELLS As Byte = 20

Public Const MAX_STATUS As Byte = 24

Public Const STATS_PER_LEVEL As Byte = 3

Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public Const MAX_MAPS As Integer = 1000
Public Const MAX_MAPX As Byte = 15
Public Const MAX_MAPY As Byte = 11
Public Const MAP_MORAL_SAFE As Byte = 0
Public Const MAP_MORAL_NONE As Byte = 1

' Image constants
Public Const PIC_X As Integer = 32
Public Const PIC_Y As Integer = 32

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_MOBSPAWN As Byte = 7

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_EQUIPMENT As Byte = 1
Public Const ITEM_TYPE_POTION As Byte = 2
Public Const ITEM_TYPE_KEY As Byte = 3
Public Const ITEM_TYPE_SPELL As Byte = 4

Public Enum ItemRarity
    Common = 0
    Uncommon
    Rare
    Epic
End Enum

Public Enum ItemBind
    None = 0
    BindOnEquip
    BindOnPickup
End Enum

' Shop Constants
Public Const SHOP_TYPE_SHOP As Byte = 0
Public Const SHOP_TYPE_INN As Byte = 1

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
'Public Const MOVING_RUNNING = 2

' Admin constants
Public Const ADMIN_MONITER As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOR_QUEST As Byte = 4

' Spell constants
Public Const SPELL_TYPE_VITAL As Byte = 0
Public Const SPELL_TYPE_OVERTIME As Byte = 1
Public Const SPELL_TYPE_BUFF As Byte = 2
Public Const SPELL_TYPE_REVIVE As Byte = 3
Public Const SPELL_TYPE_WARP As Byte = 4

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

'Spawn Map
Public Const START_MAP As Byte = 1
Public Const START_X As Byte = 11
Public Const START_Y As Byte = 8

Public Const ADMIN_LOG As String = "admin.txt"
Public Const PLAYER_LOG As String = "player.txt"

' Text Variables

Public Const MAX_LINES As Long = 500

Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15

Public Const SayColor = White
Public Const GlobalColor = Brown
Public Const TellColor = BrightGreen
Public Const AdminColor = BrightCyan
Public Const JoinLeftColor = Cyan
Public Const NpcColor = Brown
Public Const AlertColor = BrightCyan
Public Const NewMapColor = Pink
Public Const ActionColor = Yellow

Public Const MAX_GUILD_RANKS As Byte = 5

' Consts for Action Msgs
Public Const ACTIONMSG_STATIC As Byte = 0
Public Const ACTIONMSG_SCROLL As Byte = 1
Public Const ACTIONMSG_SCREEN As Byte = 2

'//////////////////////
'// Public Variables //
'//////////////////////

' For our online player counts
Public OnlinePlayers() As Long
Public OnlinePlayersCount As Long

' Maximum classes
Public MAX_CLASSES As Byte

' Used for logging
Public ServerLog As Boolean

' Used for shutting down server
Public ShutOn As Boolean

' Used for server loop
Public ServerOnline As Byte

Public EncryptPackets As Byte           ' Flag for encrypting packets - 1 for Yes - 0 for No
Public PacketKeys() As String           ' Holds our array of packet encryption keys

' Consts for data pathcs
Public AccountPath As String
Public EmoticonPath As String
Public GuildPath As String
Public ItemPath As String
Public LogPath As String
Public MapPath As String
Public NpcPath As String
Public ShopPath As String
Public SpellPath As String
Public AnimationPath As String
Public QuestPath As String

Public GameMOTD As String
Public ExpMod As Long               ' Used to control the modifier on global exp gain

' Below are used to track the size of the specified UDT
' Set in server init
Public AnimationSize As Long    ' Animation UDT
Public EmoticonSize As Long     ' Emoticon UDT
Public ItemSize As Long         ' Item UDT
Public NpcSize As Long          ' Npc UDT
Public ShopSize As Long         ' Shop UDT
Public SpellSize As Long        ' Spell UDT
Public QuestSize As Long        ' Quest UDT

' Below are used for Enum Names
Public StatName() As String
Public StatAbbreviation() As String
Public VitalName() As String
Public EquipmentName() As String

'//////////////////
'// Public Enums //
'//////////////////

' Menu states           ' Used for client messages
Public Enum MenuStates
    MainMenu = 0
    NewAccount
    Login
    GetChars
    NewChar
    AddChar
    DelChar
    UserChar
    Chars
    Shutdown
End Enum

Public Enum Stats
    Strength = 1
    Dexterity
    Vitality
    Intelligence
    Wisdom
    ' Make sure Stat_Count is below everything else
    Stat_Count = Wisdom
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    SP
    ' Mak sure Vital_Count is below everything else
    Vital_Count = SP
End Enum

Public Enum Slots
    Weapon = 1
    Armor
    Helmet
    Shield
    ' Make sure Slot_Count is below everything else
    Slot_Count = Shield
End Enum

Public Enum Targets
    Target_None = 1
    Target_SelfOnly = 2         ' Will only be cast on you - no matter what other flags are set
    Target_PlayerHostile = 4    ' Can be cast on other players - on pvp map
    Target_PlayerBeneficial = 8 ' Can be cast on other players - on any map
    Target_Npc = 16             ' Can be cast on all npcs
    Target_PlayerParty = 32     ' Can only be cast on party members - Will override other flags except selfonly
    ' Make sure Target_Count is below everything else
    Target_Count = 6
End Enum

Public Enum DamageTypes
    Physical = 1
    Magical
End Enum

Public Enum ElementTypes
    Fire = 1
    Water
    Earth
    Wind
    Holy
    Shadow
End Enum

'*********************
'   For Packet Data
'*********************
Public Enum CMsgTypes                   ' Server -> Client // Make sure it's the same for the server
    CMsgAlertMsg = 1
    CMsgClientMsg
    CMsgAllChars
    CMsgLoginOk
    CMsgNewCharClasses
    CMsgClassesData
    CMsgInGame
    CMsgPlayerInv
    CMsgPlayerInvUpdate
    CMsgPlayerWornEq
    CMsgPlayerVital
    CMsgPlayerStats
    CMsgPlayerData
    CMsgPlayerMove
    CMsgNpcMove
    CMsgPlayerDir
    CMsgNpcDir
    CMsgPlayerXY
    CMsgAttack
    CMsgNpcAttack
    CMsgCheckForMap
    CMsgMapData
    CMsgMapItemData
    CMsgMapNpcData
    CMsgMapDone
    CMsgChatMsg
    CMsgSpawnItem
    CMsgItemEditor
    CMsgUpdateItem
    CMsgUpdateItems
    CMsgEditItem
    CMsgEditEmoticon
    CMsgUpdateEmoticon
    CMsgUpdateEmoticons
    CMsgEmoticonEditor
    CMsgCheckEmoticon
    CMsgNewTarget
    CMsgSpawnNpc
    CMsgNpcDead
    CMsgNpcEditor
    CMsgUpdateNpc
    CMsgUpdateNpcs
    CMsgEditNpc
    CMsgMapKey
    CMsgEditMap
    CMsgShopEditor
    CMsgUpdateShop
    CMsgUpdateShops
    CMsgEditShop
    CMsgSpellEditor
    CMsgUpdateSpell
    CMsgUpdateSpells
    CMsgEditSpell
    CMsgAnimationEditor
    CMsgUpdateAnimation
    CMsgUpdateAnimations
    CMsgEditAnimation
    CMsgTrade
    CMsgSpells
    CMsgActionMsg
    CMsgAnimation
    CMsgPlayerGuild
    CMsgPlayerExp
    CMsgCancelSpell
    CMsgSpellReady
    CMsgSpellCooldown
    CMsgLeftGame
    CMsgPlayerDead
    CMsgPlayerGold
    CMsgPlayerRevival
    ' Quest
    CMsgQuestEditor
    CMsgUpdateQuest
    CMsgUpdateQuests
    CMsgEditQuest
    CMsgAvailableQuests
    CMsgPlayerQuests
    CMsgPlayerQuest
    'The following enum member automatically stores the number of messages,
    'since it is last. Any new messages must be placed above this entry.
    CMSG_COUNT
End Enum

Public Enum SMsgTypes                   ' Client -> Server // Make sure it's the same for the server
    SMsgGetClasses = 1
    SMsgNewAccount
    SMsgLogin
    SMsgRequestEditEmoticon
    SMsgEditEmoticon
    SMsgSaveEmoticon
    SMsgCheckEmoticon
    SMsgAddChar
    SMsgDelChar
    SMsgUseChar
    SMsgSayMsg
    SMsgEmoteMsg
    SMsgGlobalMsg
    SMsgAdminMsg
    SMsgPartyMsg
    SMsgPlayerMsg
    SMsgPlayerMove
    SMsgPlayerDir
    SMsgUseItem
    SMsgUnequipSlot
    SMsgAttack
    SMsgUseStatPoint
    SMsgPlayerInfoRequest
    SMsgWarpMeTo
    SMsgWarpToMe
    SMsgWarpTo
    SMsgSetSprite
    SMsgGetStats
    SMsgClickWarp
    SMsgRequestNewMap
    SMsgMapData
    SMsgNeedMap
    SMsgMapGetItem
    SMsgMapDropItem
    SMsgMapRespawn
    SMsgMapReport
    SMsgKickPlayer
    SMsgListBans
    SMsgBanDestroy
    SMsgBanPlayer
    SMsgRequestEditMap
    SMsgRequestEditItem
    SMsgEditItem
    SMsgSaveItem
    SMsgRequestEditNpc
    SMsgEditNpc
    SMsgSaveNpc
    SMsgRequestEditShop
    SMsgEditShop
    SMsgSaveShop
    SMsgRequestEditSpell
    SMsgEditSpell
    SMsgSaveSpell
    SMsgRequestEditAnimation
    SMsgEditAnimation
    SMsgSaveAnimation
    SMsgSetAccess
    SMsgWhosOnline
    SMsgSetMOTD
    SMsgTradeRequest
    SMsgSearch
    SMsgParty
    SMsgJoinParty
    SMsgLeaveParty
    SMsgCast
    SMsgRequestLocation
    SMsgFix
    SMsgChangeInvSlots
    SMsgClearTarget
    SMsgGCreate
    SMsgSetGMOTD
    SMsgGQuit
    SMsgGDelete
    SMsgGPromote
    SMsgGDemote
    SMsgGKick
    SMsgGInvite
    SMsgGJoin
    SMsgGDecline
    SMsgGuildMsg
    SMsgKill
    SMsgSetBound
    SMsgCancelSpell
    SMsgRelease
    SMsgRevive
    SMsgRequestEditQuest
    SMsgEditQuest
    SMsgSaveQuest
    SMsgAcceptQuest
    SMsgCompleteQuest
    SMsgDropQuest
    'The following enum member automatically stores the number of messages,
    'since it is last. Any new messages must be placed above this entry.
    SMSG_COUNT
End Enum

' Has to be below the enums
Public HandleDataSub(SMSG_COUNT) As Long
