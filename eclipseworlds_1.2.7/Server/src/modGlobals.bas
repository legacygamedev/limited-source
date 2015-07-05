Attribute VB_Name = "modGlobals"
Option Explicit

' Exp Rate
Public EXP_RATE As Byte

' Used for gradually giving back NPCs hp
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean

' Text variables
Public vbQuote As String

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines As Long

' Used to handle shutting down server with countdown.
Public IsShuttingDown As Boolean
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long
Public ElapsedTime As Long

' High Indexing
Public Player_HighIndex As Long

' CPS Lock
Public CPSUnlock As Boolean

' Packet Tracker
Public PacketsIn As Long

Public PacketsOut As Long

' Server Online Time
Public ServerSeconds As Byte

Public ServerMinutes As Byte

Public ServerHours As Long

' Data Sizes
Public MAX_MAPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_ANIMATIONS As Long
Public MAX_SHOPS As Long
Public MAX_SPELLS As Long
Public MAX_RESOURCES As Long
Public MAX_QUESTS As Long
Public MAX_BANS As Long
Public MAX_TITLES As Long
Public MAX_MORALS As Long
Public MAX_CLASSES As Long
Public MAX_EMOTICONS As Long
