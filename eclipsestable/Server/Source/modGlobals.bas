Attribute VB_Name = "modGlobals"
Option Explicit

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

Public Map() As MapRec
Public MapCache() As String
Public TempTile() As TempTileRec
Public PlayersOnMap() As Long
Public Player() As AccountRec
Public ClassData() As ClassRec
Public Item() As ItemRec
Public NPC() As NpcRec
Public MapItem() As MapItemRec
Public MapNPC() As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Guild() As GuildRec
Public Emoticons() As EmoRec
Public Element() As ElementRec
Public Experience() As Long
Public CTimers As Collection

Public Arrows(1 To MAX_ARROWS) As ArrowRec

Public addHP As StatRec
Public addMP As StatRec
Public addSP As StatRec

Public temp As Integer

Public START_MAP As Long
Public START_X As Long
Public START_Y As Long

Global PlayerI As Byte

' Winsock globals
Public GAME_PORT As Long

' Map Control
Public IS_SCROLLING As Long

' Used for respawning items
Public SpawnSeconds As Long

' Used for weather effects
Public WeatherType As Long
Public GameTime As Long
Public WeatherLevel As Long
Public GameClock As String
Public Gamespeed As Long

Public Hours As Integer
Public Seconds As Long
Public Minutes As Integer

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back players and npcs hp
Public GiveHPTimer As Long
Public GiveMPTimer As Long
Public GiveSPTimer As Long
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean

Public TimeDisable As Boolean

' Our dll cls
Global MyScript As clsSadScript

' Our hardcoded commands
Public clsScriptCommands As clsCommands

' Our GameServer and Sockets objects
Public GameServer As clsServer
Public Sockets As colSockets
