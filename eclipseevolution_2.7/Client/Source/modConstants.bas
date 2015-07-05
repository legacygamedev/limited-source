Attribute VB_Name = "modConstants"
Option Explicit

'Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
'Public Const SRCPAINT = &HEE0086

Public Const TilesInSheets = 14 'Number of tiles on a tilesheet (width)
Public Const ExtraSheets = 10

Public Const VK_UP = &H26
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11

' Menu states
Public Const MENU_STATE_NEWACCOUNT = 0
Public Const MENU_STATE_DELACCOUNT = 1
Public Const MENU_STATE_LOGIN = 2
Public Const MENU_STATE_GETCHARS = 3
Public Const MENU_STATE_NEWCHAR = 4
Public Const MENU_STATE_ADDCHAR = 5
Public Const MENU_STATE_DELCHAR = 6
Public Const MENU_STATE_USECHAR = 7
Public Const MENU_STATE_INIT = 8
Public Const MENU_STATE_AUTO_LOGIN = 9

' Speed moving vars
Public Const WALK_SPEED = 4
Public Const RUN_SPEED = 8
Public Const GM_WALK_SPEED = 4
Public Const GM_RUN_SPEED = 8
Public SS_WALK_SPEED
Public SS_RUN_SPEED
' Set the variable to your desire,
' 32 is a safe and recommended setting

' Used for AlwaysOnTop
Public Const FLAGS As Long = 3
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST As Long = -2

Public SetTop As Boolean
Public Declare Function SetWindowPos Lib "user32" (ByVal H As Long, ByVal hb As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal f As Long) As Long

Public Const MAX_ARROWS = 100
Public Const MAX_PLAYER_ARROWS = 100
Public Const MAX_BUBBLES = 20
Public Const MAX_BANK = 50
Public Const MAX_INV = 24
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 66
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10
Public Const MAX_SHOP_ITEMS = 25

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE = "270"

' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map morals
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_NO_PENALTY = 2
Public Const MAP_MORAL_HOUSE = 3

' Image constants
Public Const PIC_X = 32
Public Const PIC_Y = 32

' Tile consants
Public Const TILE_TYPE_WALKABLE = 0
Public Const TILE_TYPE_BLOCKED = 1
Public Const TILE_TYPE_WARP = 2
Public Const TILE_TYPE_ITEM = 3
Public Const TILE_TYPE_NPCAVOID = 4
Public Const TILE_TYPE_KEY = 5
Public Const TILE_TYPE_KEYOPEN = 6
Public Const TILE_TYPE_HEAL = 7
Public Const TILE_TYPE_KILL = 8
Public Const TILE_TYPE_SHOP = 9
Public Const TILE_TYPE_CBLOCK = 10
Public Const TILE_TYPE_ARENA = 11
Public Const TILE_TYPE_SOUND = 12
Public Const TILE_TYPE_SPRITE_CHANGE = 13
Public Const TILE_TYPE_SIGN = 14
Public Const TILE_TYPE_DOOR = 15
Public Const TILE_TYPE_NOTICE = 16
Public Const TILE_TYPE_CHEST = 17
Public Const TILE_TYPE_CLASS_CHANGE = 18
Public Const TILE_TYPE_SCRIPTED = 19
'Public Const TILE_TYPE_NPC_SPAWN = 20
Public Const TILE_TYPE_HOUSE = 21
'Public Const TILE_TYPE_CANON = 22
Public Const TILE_TYPE_BANK = 23
'Public Const TILE_TYPE_SKILL = 24
Public Const TILE_TYPE_GUILDBLOCK = 25
Public Const TILE_TYPE_HOOKSHOT = 26
Public Const TILE_TYPE_WALKTHRU = 27
Public Const TILE_TYPE_ROOF = 28
Public Const TILE_TYPE_ROOFBLOCK = 29
Public Const TILE_TYPE_ONCLICK = 30
Public Const TILE_TYPE_LOWER_STAT = 31

' Item constants
Public Const ITEM_TYPE_NONE = 0
Public Const ITEM_TYPE_WEAPON = 1
Public Const ITEM_TYPE_TWO_HAND = 2
Public Const ITEM_TYPE_ARMOR = 3
Public Const ITEM_TYPE_HELMET = 4
Public Const ITEM_TYPE_SHIELD = 5
Public Const ITEM_TYPE_LEGS = 6
Public Const ITEM_TYPE_RING = 7
Public Const ITEM_TYPE_NECKLACE = 8
Public Const ITEM_TYPE_POTIONADDHP = 9
Public Const ITEM_TYPE_POTIONADDMP = 10
Public Const ITEM_TYPE_POTIONADDSP = 11
Public Const ITEM_TYPE_POTIONSUBHP = 12
Public Const ITEM_TYPE_POTIONSUBMP = 13
Public Const ITEM_TYPE_POTIONSUBSP = 14
Public Const ITEM_TYPE_KEY = 15
Public Const ITEM_TYPE_CURRENCY = 16
Public Const ITEM_TYPE_SPELL = 17
Public Const ITEM_TYPE_SCRIPTED = 18
Public Const ITEM_TYPE_THROW = 19
Public Const ITEM_TYPE_WARP = 20

' Direction constants
Public Const DIR_UP = 0
Public Const DIR_DOWN = 1
Public Const DIR_LEFT = 2
Public Const DIR_RIGHT = 3

' Constants for player movement
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2

' Weather constants
Public Const WEATHER_NONE = 0
Public Const WEATHER_RAINING = 1
Public Const WEATHER_SNOWING = 2
Public Const WEATHER_THUNDER = 3

' Time constants
Public Const TIME_DAY = 0
Public Const TIME_NIGHT = 1

' Admin constants
Public Const ADMIN_MONITER = 1
Public Const ADMIN_MAPPER = 2
Public Const ADMIN_DEVELOPER = 3
Public Const ADMIN_CREATOR = 4

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED = 1
Public Const NPC_BEHAVIOR_FRIENDLY = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER = 3
Public Const NPC_BEHAVIOR_GUARD = 4
Public Const NPC_BEHAVIOR_SCRIPTED = 5

' Speach bubble constants
Public Const DISPLAY_BUBBLE_TIME = 2000 ' In milliseconds.
Public DISPLAY_BUBBLE_WIDTH As Byte ' not a constant
Public Const MAX_BUBBLE_WIDTH = 6 ' In tiles. Includes corners.
Public Const MAX_LINE_LENGTH = 23 ' In characters.
Public Const MAX_LINES = 3

' Spell constants
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_SCRIPTED = 6
Public Const SPELL_TYPE_TEMP = 7

Public Const BLACK = 0
Public Const BLUE = 1
Public Const GREEN = 2
Public Const CYAN = 3
Public Const RED = 4
Public Const MAGENTA = 5
Public Const BROWN = 6
Public Const GREY = 7
Public Const DARKGREY = 8
Public Const BRIGHTBLUE = 9
Public Const BRIGHTGREEN = 10
Public Const BRIGHTCYAN = 11
Public Const BRIGHTRED = 12
Public Const PINK = 13
Public Const YELLOW = 14
Public Const WHITE = 15

Public Const SayColor = GREY
Public Const GlobalColor = GREEN
Public Const BroadcastColor = WHITE
Public Const TellColor = WHITE
Public Const EmoteColor = WHITE
Public Const AdminColor = BRIGHTCYAN
Public Const HelpColor = WHITE
Public Const WhoColor = GREY
Public Const JoinLeftColor = GREY
Public Const NpcColor = WHITE
Public Const AlertColor = WHITE
Public Const NewMapColor = GREY

