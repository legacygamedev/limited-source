Attribute VB_Name = "modGlobals"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

' Used for closing key doors again
Public KeyTimer As Long

' Used for MOTD
Public MOTD As String

' Used for gradually giving back npcs hp
Public GiveNPCHPTimer As Long

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

' online player variables
Public PlayersOnline() As Integer
Public High_Index As Integer
Public TotalPlayersOnline As Integer
