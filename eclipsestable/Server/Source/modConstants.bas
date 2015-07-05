Attribute VB_Name = "modConstants"
Option Explicit

' Global Variables.
Public GAME_NAME As String
Public MAX_PLAYERS As Integer
Public MAX_CLASSES As Integer
Public MAX_SPELLS As Integer
Public MAX_SCRIPTSPELLS As Integer
Public MAX_ELEMENTS As Integer
Public MAX_MAPS As Integer
Public MAX_SHOPS As Integer
Public MAX_ITEMS As Integer
Public MAX_NPCS As Integer
Public MAX_MAP_ITEMS As Integer
Public MAX_GUILDS As Integer
Public MAX_GUILD_MEMBERS As Integer
Public MAX_PARTY_MEMBERS As Integer
Public MAX_EMOTICONS As Integer
Public MAX_LEVEL As Integer
Public MAX_SERVLINES As Long
Public SCRIPTING As Byte
Public PAPERDOLL As Byte
Public SPRITESIZE As Byte
Public CUSTOM_SPRITE As Integer
Public PKMINLVL As Integer
Public LEVEL As Integer
Public EMAIL_AUTH As Integer
Public ACC_VERIFY As Byte
Public HP_REGEN As Byte
Public HP_TIMER As Long
Public MP_REGEN As Byte
Public MP_TIMER As Long
Public SP_REGEN As Byte
Public SP_TIMER As Long
Public NPC_REGEN As Byte
Public SP_ENABLE As Byte
Public STAT1 As String
Public STAT2 As String
Public STAT3 As String
Public STAT4 As String
Public SAVETIME As Long
Public CLASSES As Byte
Public SP_ATTACK As Byte
Public SP_RUNNING As Byte

' Global Timers.
Public CHATLOG_TIMER As Long
Public SHUTDOWN_TIMER As Long
Public PLYRSAVE_TIMER As Long

' Max Settings.
Public Const MAX_ARROWS = 100
Public Const MAX_INV = 24
Public Const MAX_BANK = 50
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 66
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10
Public Const MAX_SHOP_ITEMS = 25

' Choice Constants.
Public Const NO = 0
Public Const YES = 1

' Account Constants.
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Basic Security.
Public Const SEC_CODE = "270"

' Gender Constants.
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map Coords.
Public MAX_MAPX As Long
Public MAX_MAPY As Long

' Map Morals.
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_NO_PENALTY = 2
Public Const MAP_MORAL_HOUSE = 3

' Tile Consants.
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

' Item Constants.
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

' Direction Constants.
Public Const DIR_UP = 0
Public Const DIR_DOWN = 1
Public Const DIR_LEFT = 2
Public Const DIR_RIGHT = 3

' Player Movement Constants.
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2

' Weather Constants.
Public Const WEATHER_NONE = 0
Public Const WEATHER_RAINING = 1
Public Const WEATHER_SNOWING = 2
Public Const WEATHER_THUNDER = 3

' Time Constants.
Public Const TIME_DAY = 0
Public Const TIME_NIGHT = 1

' Admin Constants.
Public Const ADMIN_MONITER = 1
Public Const ADMIN_MAPPER = 2
Public Const ADMIN_DEVELOPER = 3
Public Const ADMIN_CREATOR = 4

' NPC Constants.
Public Const NPC_BEHAVIOR_ATTACKONSIGHT = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED = 1
Public Const NPC_BEHAVIOR_FRIENDLY = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER = 3
Public Const NPC_BEHAVIOR_GUARD = 4
Public Const NPC_BEHAVIOR_SCRIPTED = 5

' Spell Constants.
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_GIVEITEM = 6
Public Const SPELL_TYPE_SCRIPTED = 6

' Target Type Constants.
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1

' Version Constants.
Public Const CLIENT_MAJOR = 2
Public Const CLIENT_MINOR = 8
Public Const CLIENT_REVISION = 0

Public Const ADMIN_LOG = "Logs\Admin.txt"
Public Const PLAYER_LOG = "Logs\Player.txt"

' Max Lines On Console.
Public Const MAX_LINES = 500

' Default System Colors.
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

' Default Message Colors.
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
