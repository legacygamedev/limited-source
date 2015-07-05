Attribute VB_Name = "modConstants"
Option Explicit

' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' Sounds
Public Const MAX_SOUNDS As Byte = 30

' Animated buttons
Public Const MAX_MENUBUTTONS As Byte = 4
Public Const MAX_MAINBUTTONS As Byte = 16
Public Const MENUBUTTON_PATH As String = "\data files\graphics\gui\menu\buttons\"
Public Const MAINBUTTON_PATH As String = "\data files\graphics\gui\main\buttons\"

' PK constants
Public Const PLAYER_KILLER As Byte = 1
Public Const PLAYER_DEFENDER As Byte = 2

' Hotbar
Public Const HotbarTop As Byte = 2
Public Const HotbarLeft As Byte = 2
Public Const HotbarOffsetX As Byte = 8

' Inventory constants
Public Const InvTop As Long = 16
Public Const InvLeft As Long = 12
Public Const InvOffsetY As Long = 3
Public Const InvOffsetX As Long = 3
Public Const InvColumns As Long = 5

' Bank constants
Public Const BankTop As Long = 38
Public Const BankLeft As Long = 42
Public Const BankOffsetX As Long = 4
Public Const BankOffsetY As Long = 4
Public Const BankColumns As Long = 11

' Spells constants
Public Const SpellTop As Long = 16
Public Const SpellLeft As Long = 12
Public Const SpellOffsetY As Long = 3
Public Const SpellOffsetX As Long = 3
Public Const SpellColumns As Long = 5

' Shop constants
Public Const ShopTop As Byte = 6
Public Const ShopLeft As Byte = 8
Public Const ShopOffsetX As Byte = 4
Public Const ShopOffsetY As Byte = 2
Public Const ShopColumns As Byte = 5

' Character consts
Public Const EqTop As Byte = 202
Public Const EqLeft As Byte = 18
Public Const EqOffsetX As Byte = 10
Public Const EqColumns As Byte = 4

' Value constants
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' Hardcoded sound effects
Public Const Sound_ButtonHover As String = "Cursor1.wav"
Public Const Sound_Thunder As String = "Thunder1.ogg"

' Battle Music
Public BattleMusicActive As Boolean

' Path constants
Public Const SOUND_PATH As String = "\data files\sound\"
Public Const MUSIC_PATH As String = "\data files\music\"

' Button sound constants
Public Const ButtonHover As String = "Cursor1.ogg"
Public Const ButtonClick As String = "Decision3.ogg"
Public Const ButtonBuzzer As String = "Buzzer3.ogg"

' Ping constant
Public PingToDraw As String

' Gfx Path and variables
Public Const GFX_PATH As String = "\data files\graphics\"
Public Const GFX_EXT As String = ".png"

' Font path
Public Const FONT_PATH As String = "\data files\graphics\fonts\"

' Key constants
Public Const VK_A As Long = &H41
Public Const VK_D As Long = &H44
Public Const VK_S As Long = &H53
Public Const VK_W As Long = &H57
Public Const VK_SHIFT As Long = &H10
Public Const VK_CONTROL As Long = &H11
Public Const VK_TAB As Long = &H9
Public Const VK_LEFT As Long = &H25
Public Const VK_UP As Long = &H26
Public Const VK_RIGHT As Long = &H27
Public Const VK_DOWN As Long = &H28

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

' Speed moving variables
Public Const MOVEMENT_SPEED As Byte = 4

' Tile size constants
Public Const PIC_X As Integer = 32
Public Const PIC_Y As Integer = 32

' Random tiles
Public RandomTile(0 To 3) As Integer
Public RandomTileSheet(0 To 3) As Byte
Public RandomTileSelected As Long

' Map Editor constants
Public CurrentLayer As Byte

' Sprite, item, and spell size constants
Public Const SIZE_X As Integer = 32
Public Const SIZE_Y As Integer = 32

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
Public Const EDITOR_QUEST As Byte = 13

' Dialogue box constants
Public Const DIALOGUE_TYPE_NONE As Byte = 0
Public Const DIALOGUE_TYPE_TRADE As Byte = 1
Public Const DIALOGUE_TYPE_FORGET As Byte = 2
Public Const DIALOGUE_TYPE_PARTY As Byte = 3
Public Const DIALOGUE_TYPE_RESETSTATS As Byte = 4
Public Const DIALOGUE_TYPE_ADDFRIEND As Byte = 5
Public Const DIALOGUE_TYPE_REMOVEFRIEND As Byte = 6
Public Const DIALOGUE_TYPE_ADDFOE As Byte = 7
Public Const DIALOGUE_TYPE_REMOVEFOE As Byte = 8
Public Const DIALOGUE_TYPE_GUILD As Byte = 9
Public Const DIALOGUE_TYPE_DESTROYITEM As Byte = 10
Public Const DIALOGUE_TYPE_CHANGEGUILDACCESS As Byte = 11
Public Const DIALOGUE_TYPE_PARTYINVITE As Byte = 12
Public Const DIALOGUE_TYPE_GUILDINVITE As Byte = 13
Public Const DIALOGUE_TYPE_GUILDREMOVE As Byte = 14
Public Const DIALOGUE_TYPE_GUILDDISBAND As Byte = 15

' X/Y Constants
Public HalfX As Integer
Public HalfY As Integer
Public ScreenX As Integer
Public ScreenY As Integer
Public StartXValue As Integer
Public StartYValue As Integer
Public EndXValue As Integer
Public EndYValue As Integer
Public CameraEndXValue As Integer
Public CameraEndYValue As Integer

' Autotiles
Public Const AUTO_INNER As Byte = 1
Public Const AUTO_OUTER As Byte = 2
Public Const AUTO_HORIZONTAL As Byte = 3
Public Const AUTO_VERTICAL As Byte = 4
Public Const AUTO_FILL As Byte = 5

' Autotile types
Public Const AUTOTILE_NONE As Byte = 0
Public Const AUTOTILE_NORMAL As Byte = 1
Public Const AUTOTILE_FAKE As Byte = 2
Public Const AUTOTILE_ANIM As Byte = 3
Public Const AUTOTILE_CLIFF As Byte = 4
Public Const AUTOTILE_WATERFALL As Byte = 5

' Rendering
Public Const RENDER_STATE_NONE As Long = 0
Public Const RENDER_STATE_NORMAL As Long = 1
Public Const RENDER_STATE_AUTOTILE As Long = 2

' Chat Bubble
Public Const ChatBubbleWidth As Long = 200

Public Const EFFECT_TYPE_FADEIN As Long = 1
Public Const EFFECT_TYPE_FADEOUT As Long = 2
Public Const EFFECT_TYPE_FLASH As Long = 3
Public Const EFFECT_TYPE_FOG As Long = 4
Public Const EFFECT_TYPE_WEATHER As Long = 5
Public Const EFFECT_TYPE_TINT As Long = 6

' ********************************************************
' * The values below must match with the server's values *
' ********************************************************

' General constants
Public Const MAX_PLAYERS As Byte = 70
Public Const MAX_MAP_ITEMS As Byte = 50
Public Const MAX_MAP_NPCS As Byte = 50
Public Const MAX_EVENTS As Long = 255
Public Const MAX_TRADES As Byte = 30
Public Const MAX_BANK As Byte = 88
Public Const MAX_NPC_DROPS As Byte = 25
Public Const MAX_PARTYS As Byte = 35
Public Const MAX_PARTY_MEMBERS As Byte = 4
Public Const MAX_BUFFS As Byte = 30
Public Const MAX_HOTBAR As Byte = 12
Public Const MAX_GUILDS As Byte = 50
Public Const MAX_NPC_SPELLS As Byte = 5
Public Const MAX_SWITCHES As Long = 1000
Public Const MAX_VARIABLES As Long = 1000
Public Const MAX_WEATHER_PARTICLES As Long = 250
Public Const MAX_COMMON_EVENTS As Long = 100

' Player constants
Public Const MAX_PEOPLE As Byte = 100
Public Const MAX_PLAYER_SPELLS As Byte = 35
Public Const MAX_INV As Byte = 35
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
Public Const HelpColor As Byte = BrightBlue

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' Length constants
Public Const NAME_LENGTH As Byte = 21
Public Const FILE_LENGTH As Byte = 50

' Gender constants
Public Const Gender_MALE As Byte = 0
Public Const Gender_FEMALE As Byte = 1

' Map constants
Public MIN_MAPX As Byte
Public MIN_MAPY As Byte

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

' Weather Type Constants
Public Const WEATHER_TYPE_NONE As Byte = 0
Public Const WEATHER_TYPE_RAIN As Byte = 1
Public Const WEATHER_TYPE_SNOW As Byte = 2
Public Const WEATHER_TYPE_HAIL As Byte = 3
Public Const WEATHER_TYPE_SANDSTORM As Byte = 4
Public Const WEATHER_TYPE_STORM As Byte = 5

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
Public Const MOVING_WALKING As Byte = 1
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
Public Const SPELL_TYPE_WARP As Byte = 4
Public Const SPELL_TYPE_RECALL As Byte = 5
Public Const SPELL_TYPE_WARPTOTARGET As Byte = 6

' Target Type Constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2
Public Const TARGET_TYPE_EVENT As Byte = 3

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Byte = 0
Public Const ACTIONMSG_SCROLL As Byte = 1
Public Const ACTIONMSG_SCREEN As Byte = 2

Public Const WM_MOUSEMOVE      As Long = &H200
Public Const WM_LBUTTONDOWN    As Long = &H201
Public Const WM_LBUTTONUP      As Long = &H202
Public Const WM_CAPTURECHANGED As Long = &H215
Public Const WM_GETMINMAXINFO  As Long = &H24
Public Const WM_ACTIVATEAPP    As Long = &H1C
Public Const WM_SETFOCUS       As Long = &H7
Public Const WM_MOUSEWHEEL     As Long = &H20A
Public Const WM_NCACTIVATE     As Long = &H86
Public Const WM_MOVE           As Long = &H3
Public Const WM_DESTROY        As Long = &H2
