Attribute VB_Name = "modGlobals"
Option Explicit

' Used for respawning items
Public SpawnSeconds As Long

'Player Saving Constants
Global PlayerI As Byte

' Used for weather effects
Public GameWeather As Long
Public WeatherSeconds As Long
Public GameTime As Long
Public TimeSeconds As Long

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back players and npcs hp
Public GiveHPTimer As Long
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

Public Map(1 To MAX_MAPS) As MapRec
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public Player(1 To MAX_PLAYERS) As AccountRec
Public Class() As ClassRec
Public Item(0 To MAX_ITEMS) As ItemRec
Public Npc(0 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Guild(1 To MAX_GUILDS) As GuildRec

