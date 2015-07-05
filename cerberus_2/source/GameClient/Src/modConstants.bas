Attribute VB_Name = "modConstants"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

' Winsock globals
Public GAME_IP As String
Public GAME_PORT As Long

' Website
Public GAME_WEBSITE As String

' Font variables
Public Const FONT_NAME = "fixedsys"
Public Const FONT_SIZE = 18

' Map Path and variables
Public Const MAP_PATH = "\maps\"
Public Const MAP_EXT = ".map"

' Gfx Path and variables
Public Const GFX_PATH = "\gfx\"
Public Const GFX_EXT = ".bmp"

' API constants
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086

' Key constants
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
'Public Const MENU_STATE_GETCHARS = 3
Public Const MENU_STATE_NEWCHAR = 4
Public Const MENU_STATE_ADDCHAR = 5
Public Const MENU_STATE_DELCHAR = 6
Public Const MENU_STATE_USECHAR = 7
'Public Const MENU_STATE_INIT = 8

' Speed moving vars
Public Const WALK_SPEED = 2
Public Const RUN_SPEED = 4

Public Const Black = 0
Public Const Blue = 1
Public Const Green = 2
Public Const Cyan = 3
Public Const Red = 4
Public Const Magenta = 5
Public Const Brown = 6
Public Const Grey = 7
Public Const DarkGrey = 8
Public Const BrightBlue = 9
Public Const BrightGreen = 10
Public Const BrightCyan = 11
Public Const BrightRed = 12
Public Const Pink = 13
Public Const Yellow = 14
Public Const White = 15

'Public Const SayColor = Grey
'Public Const GlobalColor = BrightGreen
'Public Const TellColor = Cyan
'Public Const EmoteColor = BrightCyan
Public Const HelpColor = Magenta
'Public Const WhoColor = Pink
'Public Const JoinLeftColor = DarkGrey
'Public Const NpcColor = Brown
Public Const AlertColor = Red
'Public Const NewMapColor = Pink

' Damage Variables
Public DmgDamage As Long
Public DmgColor As Byte
Public DmgTime As Long
Public NPCDmgDamage As Long
Public NPCDmgColor As Byte
Public NPCDmgTime As Long
Public NPCWho As Long
Public PKDmgDamage As Long
Public PKDmgColor As Byte
Public PKDmgTime As Long
Public PKWho As Long
Public ResourceDmgDamage As Long
Public ResourceDmgColor As Byte
Public ResourceDmgTime As Long
Public ResourceDmgWho As Long

' Blitted Message Variables
Public MsgMessage As String
Public MessageTime As Long
Public MessageColor As Byte
Public WarnMessage As String
Public WarnMsgTime As Long
Public WarnMsgColor As Byte
Public ResourceMsgMessage As String
Public ResourceMsgTime As Long
Public ResourceMsgColor As Byte
Public ResourceWho As Long
Public NpcMsgMessage As String
Public NpcMessageTime As Long
Public NpcMessageColor As Byte
Public NpcMsgWho As Long
Public PKMsgMessage As String
Public PKMsgTime As Long
Public PKMsgColor As Byte
Public PKMsgWho As Long
Public ItemMsgMessage As String
Public ItemMsgTime As Long
Public ItemMsgColor As Byte
Public ItemWho As Long

Public ii As Long, iii As Long, iv As Long, vi As Long, vii As Long, viii As Long, viv As Long
Public xi As Long, xii As Long, xiv As Long
Public sx As Long

' General constants
Public GAME_NAME As String
Public WEBSITE As String
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_SKILLS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_GUILDS As Long
Public MAX_GUILD_MEMBERS As Long
'Public MAX_EMOTICONS As Long
Public MAX_QUESTS As Long
'Public Const GAME_NAME = "Thallingorn"
'Public Const MAX_PLAYERS = 5
Public Const MAX_PLAYER_SPELLS = 10
Public Const MAX_PLAYER_SKILLS = 10
Public Const MAX_PLAYER_QUESTS = 5
Public Const MAX_PLAYER_ARROWS = 10
Public Const MAX_PLAYER_MAPS = 3
'Public Const MAX_ITEMS = 50
'Public Const MAX_NPCS = 50
Public Const MAX_INV = 30
'Public Const MAX_MAP_ITEMS = 20
Public Const MAX_MAP_NPCS = 5
Public Const MAX_MAP_RESOURCES = 20
'Public Const MAX_SHOPS = 50
'Public Const MAX_SPELLS = 50
'Public Const MAX_SKILLS = 50
'Public Const MAX_QUESTS = 50
Public Const MAX_TRADES = 8
Public Const MAX_GIVE_ITEMS = 5
Public Const MAX_GIVE_VALUE = 5
Public Const MAX_GET_ITEMS = 3
Public Const MAX_GET_VALUE = 3
Public Const MAX_NPC_QUESTS = 5
Public Const MAX_NPC_DROPS = 5
'Public Const MAX_GUILDS = 20
'Public Const MAX_GUILD_MEMBERS = 10
Public Const MAX_GUILD_MAPS = 5
Public Const MAX_GUILD_QUESTS = 10

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map constants
'Public Const MAX_MAPS = 50
Public Const MAX_MAPX = 31
Public Const MAX_MAPY = 23
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_PLAYER = 2
Public Const MAP_MORAL_GUILD = 3

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
Public Const TILE_TYPE_PUSHBLOCK = 7
Public Const TILE_TYPE_NSPAWN = 8
Public Const TILE_TYPE_RSPAWN = 9

' Item constants
Public Const ITEM_TYPE_NONE = 0
Public Const ITEM_TYPE_WEAPON = 1
Public Const ITEM_TYPE_ARMOR = 2
Public Const ITEM_TYPE_HELMET = 3
Public Const ITEM_TYPE_SHIELD = 4
Public Const ITEM_TYPE_POTIONADDHP = 5
Public Const ITEM_TYPE_POTIONADDMP = 6
Public Const ITEM_TYPE_POTIONADDSP = 7
Public Const ITEM_TYPE_POTIONSUBHP = 8
Public Const ITEM_TYPE_POTIONSUBMP = 9
Public Const ITEM_TYPE_POTIONSUBSP = 10
Public Const ITEM_TYPE_KEY = 11
Public Const ITEM_TYPE_CURRENCY = 12
Public Const ITEM_TYPE_SPELL = 13
Public Const ITEM_TYPE_TOOL = 14
Public Const ITEM_TYPE_AMULET = 15
Public Const ITEM_TYPE_RING = 16
Public Const ITEM_TYPE_SKILL = 17
Public Const ITEM_TYPE_ARROW = 18

' Weapon subtype constants
Public Const WEAPON_SUBTYPE_NONE = 0
Public Const WEAPON_SUBTYPE_DAGGER = 1
Public Const WEAPON_SUBTYPE_SWORD = 2
Public Const WEAPON_SUBTYPE_AXE = 3
Public Const WEAPON_SUBTYPE_MACE = 4
Public Const WEAPON_SUBTYPE_WAND = 5
Public Const WEAPON_SUBTYPE_STAFF = 6
Public Const WEAPON_SUBTYPE_BOW = 7

' Direction constants
Public Const DIR_UP = 0
Public Const DIR_DOWN = 1
Public Const DIR_LEFT = 2
Public Const DIR_RIGHT = 3

' Constants for player movement
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED = 1
Public Const NPC_BEHAVIOR_FRIENDLY = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER = 3
Public Const NPC_BEHAVIOR_GUARD = 4
Public Const NPC_BEHAVIOR_RESOURCE = 5

' Spell constants
Public Const SPELL_TYPE_STAT = 0
Public Const SPELL_TYPE_GIVEITEM = 1
' Spell subType Constants
Public Const SPELL_SUB_NONE = 0
Public Const SPELL_STAT_ADDHP = 1
Public Const SPELL_STAT_ADDMP = 2
Public Const SPELL_STAT_ADDSP = 3
Public Const SPELL_STAT_SUBHP = 4
Public Const SPELL_STAT_SUBMP = 5
Public Const SPELL_STAT_SUBSP = 6
Public Const SPELL_GIVE_ITEM = 7

' Skill Constants
Public Const SKILL_TYPE_ATTRIBUTE = 0
Public Const SKILL_TYPE_CHANCE = 1
Public Const SKILL_TYPE_BUFF = 2
Public Const SKILL_TYPE_QUALIFY = 3
Public Const SKILL_TYPE_PARTY = 4
' Skill subType Constants
Public Const SKILL_SUB_NONE = 0
Public Const SKILL_ATTRIBUTE_STR = 1
Public Const SKILL_ATTRIBUTE_DEF = 2
Public Const SKILL_ATTRIBUTE_MAGI = 3
Public Const SKILL_ATTRIBUTE_SPEED = 4
Public Const SKILL_ATTRIBUTE_DEX = 5
Public Const SKILL_CHANCE_CRIT = 6
Public Const SKILL_CHANCE_DROP = 7
Public Const SKILL_CHANCE_BLOCK = 8
Public Const SKILL_CHANCE_ACCU = 9

' Charm constants
Public Const CHARM_TYPE_ADDHP = 0
Public Const CHARM_TYPE_ADDMP = 1
Public Const CHARM_TYPE_ADDSP = 2
Public Const CHARM_TYPE_ADDSTR = 3
Public Const CHARM_TYPE_ADDDEF = 4
Public Const CHARM_TYPE_ADDMAGI = 5
Public Const CHARM_TYPE_ADDSPEED = 6
Public Const CHARM_TYPE_ADDDEX = 7
Public Const CHARM_TYPE_ADDCRIT = 8
Public Const CHARM_TYPE_ADDDROP = 9
Public Const CHARM_TYPE_ADDBLOCK = 10
Public Const CHARM_TYPE_ADDACCU = 11

' Quest type constants
Public Const QUEST_TYPE_KILL = 0
Public Const QUEST_TYPE_FETCH = 1
Public Const QUEST_TYPE_TRADE = 2
