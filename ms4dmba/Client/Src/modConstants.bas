Attribute VB_Name = "modConstants"
Option Explicit

' API Declares
Public Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' Winsock globals
Public Const GAME_IP As String = "127.0.0.1" ' "localhost"

' path constants
Public Const SOUND_PATH As String = "\sound\"
Public Const MUSIC_PATH As String = "\music\"

' Font variables
Public Const FONT_NAME As String = "fixedsys"
Public Const FONT_SIZE As Byte = 18

' Log Path and variables
Public Const LOG_DEBUG As String = "debug.txt"
Public Const LOG_PATH As String = "\logs\"

' Map Path and variables
Public Const MAP_PATH As String = "\maps\"
Public Const MAP_EXT As String = ".map"

' Gfx Path and variables
Public Const GFX_PATH As String = "\gfx\"
Public Const GFX_EXT As String = ".bmp"

' Key constants
Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

' Menu states
Public Const MENU_STATE_NEWACCOUNT As Byte = 0
Public Const MENU_STATE_DELACCOUNT As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_NEWCHAR As Byte = 4
Public Const MENU_STATE_ADDCHAR As Byte = 5
Public Const MENU_STATE_DELCHAR As Byte = 6
Public Const MENU_STATE_USECHAR As Byte = 7
Public Const MENU_STATE_INIT As Byte = 8

' Number of tiles in width in tilesets
Public Const TILESHEET_WIDTH As Integer = 7 ' * PIC_X pixels

' Speed moving vars
Public Const WALK_SPEED As Byte = 2
Public Const RUN_SPEED As Byte = 4

' Tile size constants
Public Const PIC_X As Integer = 32
Public Const PIC_Y As Integer = 32

' Sprite, item, spell size constants
Public Const SIZE_X As Integer = 32
Public Const SIZE_Y As Integer = 32

Public Const MAX_SPELLANIM As Integer = 10

' ********************************************************
' * The values below must match with the server's values *
' ********************************************************

' General constants
Public Const GAME_NAME As String = "Mirage Source 4"
Public Const GAME_PORT As Integer = 7000
Public Const MAX_PLAYERS As Byte = 70
Public Const MAX_ITEMS As Byte = 255
Public Const MAX_NPCS As Byte = 255
Public Const MAX_INV As Byte = 50
Public Const MAX_MAP_ITEMS As Byte = 20
Public Const MAX_MAP_NPCS As Byte = 5
Public Const MAX_SHOPS As Byte = 50
Public Const MAX_PLAYER_SPELLS As Byte = 20
Public Const MAX_SPELLS As Byte = 255
Public Const MAX_TRADES As Byte = 8
Public Const MAX_LEVELS As Byte = 100

' Website
Public Const GAME_WEBSITE As String = "http://www.miragesource.com"

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
Public Const MAX_MAPS As Long = 100
Public Const MAX_MAPX As Byte = 15
Public Const MAX_MAPY As Byte = 11
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6

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

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOR_GUARD As Byte = 4

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
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

Public Const HalfX As Integer = ((MAX_MAPX + 1) \ 2) * PIC_X
Public Const HalfY As Integer = ((MAX_MAPY + 1) \ 2) * PIC_Y
Public Const ScreenX As Integer = (MAX_MAPX + 1) * PIC_X
Public Const ScreenY As Integer = (MAX_MAPY + 1) * PIC_Y
