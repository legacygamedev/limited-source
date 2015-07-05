Attribute VB_Name = "modConstants"
Option Explicit

' API
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' Version constants
Public Const MAX_LINES As Integer = 500 ' Used for frmServer.txtText

' ********************************************************
' * The values below must match with the client's values *
' ********************************************************
' General constants
Public Const MAX_PLAYERS As Byte = 70
Public Const MAX_MAP_ITEMS As Byte = 50
Public Const MAX_MAP_NPCS As Byte = 50
Public Const MAX_TRADES As Byte = 30
Public Const MAX_BANK As Byte = 88
Public Const MAX_NPC_DROPS As Byte = 25
Public Const MAX_BUFFS As Byte = 30
Public Const MAX_PARTYS As Byte = 35
Public Const MAX_PARTY_MEMBERS As Byte = 4
Public Const MAX_HOTBAR As Byte = 12
Public Const MAX_GUILDS As Byte = 100
Public Const MAX_SWITCHES As Long = 1000
Public Const MAX_VARIABLES As Long = 1000
Public Const MAX_COMMON_EVENTS As Long = 100
Public Const MAX_NPC_SPELLS As Byte = 5
Public Const MAX_CHARS As Byte = 1

' Game editor constants
Public Const EDITOR_ANIMATION As Byte = 0
Public Const EDITOR_BAN As Byte = 1
Public Const EDITOR_CLASS As Byte = 2
Public Const EDITOR_EMOTICON As Byte = 3
Public Const EDITOR_ITEM As Byte = 4
Public Const EDITOR_MAP As Byte = 5
Public Const EDITOR_MORAL As Byte = 6
Public Const EDITOR_NPC As Byte = 7
Public Const EDITOR_RESOURCE As Byte = 8
Public Const EDITOR_SHOP As Byte = 9
Public Const EDITOR_SPELL As Byte = 10
Public Const EDITOR_TITLE As Byte = 11
Public Const EDITOR_EVENTS As Byte = 12
Public Const EDITOR_QUESTS As Byte = 13

' Gender constants
Public Const GENDER_MALE As Byte = 0
Public Const GENDER_FEMALE As Byte = 1

' Map constants
Public Const MIN_MAPX As Byte = 25
Public Const MIN_MAPY As Byte = 20

' Player constants
Public Const MAX_PEOPLE As Byte = 100
Public Const MAX_GUILD_MEMBERS As Byte = 100
Public Const MAX_PLAYER_SPELLS As Byte = 35
Public Const MAX_INV As Byte = 35
Public Const MAX_TREES As Byte = 3
Public Const MAX_GUILDACCESS As Byte = 4

' Text color constants
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
Public Const DarkBrown As Byte = 16
Public Const Orange As Byte = 17
Public Const SayColor As Byte = White
Public Const PrivateColor As Byte = Pink
Public Const EmoteColor As Byte = Cyan
Public Const WhoColor As Byte = BrightBlue
Public Const NPCColor As Byte = Brown
Public Const NewMapColor As Byte = BrightBlue

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' PK constants
Public Const PLAYER_KILLER As Byte = 1
Public Const PLAYER_DEFENDER As Byte = 2

' Length constants
Public Const NAME_LENGTH As Byte = 21
Public Const FILE_LENGTH As Byte = 50

' Speed moving vars
Public Const WALK_SPEED As Byte = 4
Public Const RUN_SPEED As Byte = 8

' Tile constants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_RESOURCE As Byte = 5
Public Const TILE_TYPE_NPCSPAWN As Byte = 6
Public Const TILE_TYPE_SHOP As Byte = 7
Public Const TILE_TYPE_BANK As Byte = 8
Public Const TILE_TYPE_HEAL As Byte = 9
Public Const TILE_TYPE_TRAP As Byte = 10
Public Const TILE_TYPE_SLIDE As Byte = 11
Public Const TILE_TYPE_CHECKPOINT As Byte = 12
Public Const TILE_TYPE_SOUND As Byte = 13

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_EQUIPMENT As Byte = 1
Public Const ITEM_TYPE_CONSUME As Byte = 2
Public Const ITEM_TYPE_TITLE As Byte = 3
Public Const ITEM_TYPE_SPELL As Byte = 4
Public Const ITEM_TYPE_TELEPORT As Byte = 5
Public Const ITEM_TYPE_RESETSTATS As Byte = 6
Public Const ITEM_TYPE_AUTOLIFE As Byte = 7
Public Const ITEM_TYPE_SPRITE As Byte = 8
Public Const ITEM_TYPE_RECIPE As Byte = 9

' Equip constants
Public Const BIND_ON_PICKUP As Byte = 1
Public Const BIND_ON_EQUIP As Byte = 2

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3
Public Const DIR_UPLEFT As Byte = 4
Public Const DIR_UPRIGHT As Byte = 5
Public Const DIR_DOWNLEFT As Byte = 6
Public Const DIR_DOWNRIGHT As Byte = 7

' Constants for player movement
Public Const MOVING_MOVING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Admin constants
Public Const STAFF_MODERATOR As Byte = 1
Public Const STAFF_MAPPER As Byte = 2
Public Const STAFF_DEVELOPER As Byte = 3
Public Const STAFF_ADMIN As Byte = 4
Public Const STAFF_OWNER As Byte = 5

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOR_GUARD As Byte = 2
Public Const NPC_BEHAVIOR_QUEST As Byte = 3

' Spell constants
Public Const SPELL_TYPE_DAMAGEHP As Byte = 0
Public Const SPELL_TYPE_DAMAGEMP As Byte = 1
Public Const SPELL_TYPE_HEALHP As Byte = 2
Public Const SPELL_TYPE_HEALMP As Byte = 3
Public Const SPELL_TYPE_WARP As Byte = 4
Public Const SPELL_TYPE_RECALL As Byte = 5
Public Const SPELL_TYPE_WARPTOTARGET As Byte = 6

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

Public Const TARGET_TYPE_EVENT As Byte = 3

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Byte = 0
Public Const ACTIONMSG_SCROLL As Byte = 1
Public Const ACTIONMSG_SCREEN As Byte = 2

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)

' ********************************************
Public Const ITEM_SPAWN_TIME As Long = 30000 ' 30 seconds
Public Const ITEM_DESPAWN_TIME As Long = 90000 ' 90 seconds
Public Const MAX_DOTS As Byte = 30

' How long do we have to wait before we can use another potion
Public Const PotionWaitTimer As Integer = 3000

'- Pathfinding Constant -
'1 is the old method, faster but not smart at all
'2 is the new method, smart but can slow the server down if maps are huge and alot of npcs have targets.
Public Const PathfindingType As Long = 2

Public Const EFFECT_TYPE_FADEIN As Long = 1
Public Const EFFECT_TYPE_FADEOUT As Long = 2
Public Const EFFECT_TYPE_FLASH As Long = 3
Public Const EFFECT_TYPE_FOG As Long = 4
Public Const EFFECT_TYPE_WEATHER As Long = 5
Public Const EFFECT_TYPE_TINT As Long = 6
