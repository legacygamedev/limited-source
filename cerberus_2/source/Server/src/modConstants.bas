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

Public Const ADMIN_LOG = "admin.log"
Public Const PLAYER_LOG = "player.log"
Public Const HACK_LOG = "hack.log"
Public Const PACK_LOG = "packet.log"

' Version constants
Public Const CLIENT_MAJOR = 0
Public Const CLIENT_MINOR = 3
Public Const CLIENT_REVISION = 2
Public Const Quote = """"

Public Const MAX_LINES = 100

'Socket Globals
Public GAME_IP As String
Public GAME_PORT As Long

' Image constants
Public Const PIC_X = 32
Public Const PIC_Y = 32

' True/false constants
Public Const NO = 0
Public Const YES = 1

' Colour Constants
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

Public Const SayColor = DarkGrey
Public Const GlobalColor = Blue
Public Const BroadcastColor = Magenta
Public Const TellColor = Green
Public Const EmoteColor = Cyan
Public Const AdminColor = Cyan
Public Const HelpColor = Pink
Public Const WhoColor = Pink
Public Const JoinLeftColor = DarkGrey
Public Const NpcColor = Brown
Public Const AlertColor = Red
Public Const NewMapColor = Pink

'General constants
Public Const MAX_MAP_NPCS = 5
Public Const MAX_MAP_RESOURCES = 20
Public Const MAX_PLAYER_SPELLS = 10
Public Const MAX_PLAYER_SKILLS = 10
Public Const MAX_PLAYER_QUESTS = 5
Public Const MAX_PLAYER_MAPS = 3
Public Const MAX_INV = 30
Public Const MAX_TRADES = 8
Public Const MAX_GIVE_ITEMS = 5
Public Const MAX_GIVE_VALUE = 5
Public Const MAX_GET_ITEMS = 3
Public Const MAX_GET_VALUE = 3
Public Const MAX_NPC_QUESTS = 5
Public Const MAX_NPC_DROPS = 5
Public Const MAX_GUILD_MAPS = 5
Public Const MAX_GUILD_QUESTS = 10
Public Const MAX_GUIS = 3

' Admin constants
Public Const PLAYER_BASIC = 0
Public Const PLAYER_PARTY = 1
Public Const PLAYER_GUILD = 2
Public Const ADMIN_MONITER = 3
Public Const ADMIN_MAPPER = 4
Public Const ADMIN_DEVELOPER = 5
Public Const ADMIN_CREATOR = 6

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Direction constants
Public Const DIR_UP = 0
Public Const DIR_DOWN = 1
Public Const DIR_LEFT = 2
Public Const DIR_RIGHT = 3

' Constants for player movement
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2

' Target type constants
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1
Public Const TARGET_TYPE_RESOURCE = 2
Public Const TARGET_TYPE_ITEM = 3

' Map constants
Public Const MAX_MAPX = 31
Public Const MAX_MAPY = 23
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_PLAYER = 2
Public Const MAP_MORAL_GUILD = 3

' Start location
Public Const START_MAP = 1
Public Const START_X = MAX_MAPX / 2
Public Const START_Y = MAX_MAPY / 2

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
