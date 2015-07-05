Attribute VB_Name = "modGlobals"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

' Client options
Public GAME_NAME As String
Public GAME_WEBSITE As String
Public SPRITE_OFFSET As Byte
Public TOTAL_WALKFRAMES As Byte
Public TOTAL_ATTACKFRAMES As Byte
Public TOTAL_ANIMFRAMES As Byte
Public CONFIG_STANDFRAME As Byte
Public WALKANIM_SPEED As Integer
Public MAX_PLAYERS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_SPELLS As Long
Public MAX_SHOPS As Long
Public MAX_LEVELS As Long
Public MAX_MAPS As Long
Public MAX_SIGNS As Long
Public MAX_GUILDS As Long
Public MAX_ANIMS As Long
Public IP_Source As String
Public Guild_Creation_Cost As Long
Public Guild_Creation_Item As Long
Public StatBonus() As Long
Public VitalBonus() As Long
Public GAME_NEWS As String

' Server options
Public GAME_PORT As Integer
Public ACTUAL_IP As String

' Used for MOTD
Public MOTD As String

' Used for config file in client
Public CONFIG_PASSWORD As String

' Used for gradually giving back npcs hp
Public GiveNPCHPTimer As Currency

' Used for logging
Public ServerLog As Boolean

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Text vars
Public vbQuote As String

' Maximum classes
Public Max_Classes As Byte

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines As Long

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Public Secs As Long
