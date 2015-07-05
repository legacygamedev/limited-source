Attribute VB_Name = "modConstants"
Option Explicit

' ********************************************
' **               rootSource               **
' ********************************************

' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' path constants
Public Const ADMIN_LOG As String = "admin.log"
Public Const PLAYER_LOG As String = "player.log"

' Version constants
Public Const CLIENT_MAJOR As Byte = 1
Public Const CLIENT_MINOR As Byte = 0
Public Const CLIENT_REVISION As Byte = 0

' **********************************************************
' * The values below must match with the client's values *
' **********************************************************

' General Game-Useage constants
Public GAME_NAME As String
Public GAME_PORT As Integer
Public MAX_PLAYERS As Integer
Public MAX_ITEMS As Integer
Public MAX_NPCS As Integer
Public MAX_SHOPS As Integer
Public MAX_SPELLS As Integer
Public Const MAX_PLAYER_SPELLS As Byte = 20
Public Const MAX_INV As Byte = 50
Public Const MAX_MAP_ITEMS As Byte = 20
Public Const MAX_MAP_NPCS As Byte = 5
Public Const MAX_TRADES As Byte = 8
Public Const MAX_LEVELS As Byte = 255

' Website
Public GAME_WEBSITE As String

' text color constants
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

Public Const SayColor As Byte = Grey
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = Pink
Public Const TellColor As Byte = BrightGreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = Pink
Public Const WhoColor As Byte = Pink
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = Pink

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' Account constants
Public Const NAME_LENGTH As Byte = 20
Public Const MAX_CHARS As Byte = 3

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public MAX_MAPS As Long
Public Const MAX_MAPX As Byte = 15
Public Const MAX_MAPY As Byte = 11
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_INN As Byte = 2
Public Const MAP_MORAL_ARENA As Byte = 3

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
'New Tile Constants
Public Const TILE_TYPE_HEAL As Byte = 7
Public Const TILE_TYPE_KILL As Byte = 8
Public Const TILE_TYPE_DOOR As Byte = 9
Public Const TILE_TYPE_SIGN As Byte = 10
Public Const TILE_TYPE_MSG As Byte = 11
Public Const TILE_TYPE_SPRITE As Byte = 12
Public Const TILE_TYPE_NPCSPAWN As Byte = 13
Public Const TILE_TYPE_NUDGE As Byte = 14

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_POTIONADDHP As Byte = 5
Public Const ITEM_TYPE_POTIONADDMP As Byte = 6
Public Const ITEM_TYPE_POTIONADDSP As Byte = 7
Public Const ITEM_TYPE_POTIONSUBHP As Byte = 8
Public Const ITEM_TYPE_POTIONSUBMP As Byte = 9
Public Const ITEM_TYPE_POTIONSUBSP As Byte = 10
Public Const ITEM_TYPE_KEY As Byte = 11
Public Const ITEM_TYPE_CURRENCY As Byte = 12
Public Const ITEM_TYPE_SPELL As Byte = 13

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' Npc constants
Public Const Npc_BEHAVIOR_ATTACKONSIGHT As Byte = 0
Public Const Npc_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1
Public Const Npc_BEHAVIOR_FRIENDLY As Byte = 2
Public Const Npc_BEHAVIOR_SHOPKEEPER As Byte = 3
Public Const Npc_BEHAVIOR_GUARD As Byte = 4

' Spell constants
Public Const SPELL_TYPE_ADDHP As Byte = 0
Public Const SPELL_TYPE_ADDMP As Byte = 1
Public Const SPELL_TYPE_ADDSP As Byte = 2
Public Const SPELL_TYPE_SUBHP As Byte = 3
Public Const SPELL_TYPE_SUBMP As Byte = 4
Public Const SPELL_TYPE_SUBSP As Byte = 5
Public Const SPELL_TYPE_GIVEITEM As Byte = 6

' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_Npc As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_Npc As Byte = 2

' **********************************************
' Default starting location [Server Only]
Public START_MAP As Integer
Public START_X As Integer
Public START_Y As Integer

