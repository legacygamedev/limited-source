Attribute VB_Name = "modGlobals"
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/16/2005  Shannara   Created module.
'****************************************************************

Option Explicit

'Set Game Name
Public GAME_NAME As String

'Public GAME_NAME As String
Public MAX_PLAYERS As Long
Public MAX_MAPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_SHOPS As Long
Public MAX_SPELLS As Long
Public MAX_GUILDS As Long
Public MAX_EXPERIENCE As Long

'Set Port Public Variable
Public Game_Port As Integer

' Server's time variables
Public Server_Second As Byte
Public Server_Minute As Byte
Public Server_Hour As Byte

' Used for respawning items
Public SpawnSeconds As Long

' Used to update shops
Public ShopTimer As Long

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
Public Max_Visible_Classes As Byte

'PList enable
Public PList As Integer

'Edit Variable
Public EditType As Integer

Public Map() As MapRec
Public TempTile() As TempTileRec
Public PlayersOnMap() As Long
Public Player() As AccountRec
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc() As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Guild() As GuildRec
    
'Exp Curve before loaded
Public Experience() As Long

'Creating instance of type for system tray
Public nid As NOTIFYICONDATA
