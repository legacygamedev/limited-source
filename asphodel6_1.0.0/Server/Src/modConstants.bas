Attribute VB_Name = "modConstants"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

' path constants
Public Const ADMIN_LOG As String = "admin.log"
Public Const PLAYER_LOG As String = "player.log"

Public Const MAX_LINES As Integer = 500 ' Used for frmServer.txtText

' ********************************************************
' * The values below must match with the client's values *
' ********************************************************

' General constants
Public Const MAX_INV As Byte = 24
Public Const MAX_MAP_ITEMS As Byte = 20
Public MAX_MAP_NPCS As Byte
Public Const MAX_PLAYER_SPELLS As Byte = 10
Public Const MAX_TRADES As Byte = 64

Public Const SayColor As Byte = Color.Grey
Public Const GlobalColor As Byte = Color.BrightBlue
Public Const BroadcastColor As Byte = Color.Pink
Public Const TellColor As Byte = Color.BrightGreen
Public Const EmoteColor As Byte = Color.BrightCyan
Public Const AdminColor As Byte = Color.BrightCyan
Public Const HelpColor As Byte = Color.Pink
Public Const WhoColor As Byte = Color.Pink
Public Const JoinLeftColor As Byte = Color.DarkGrey
Public Const NpcColor As Byte = Color.Brown
Public Const AlertColor As Byte = Color.Red
Public Const NewMapColor As Byte = Color.Pink

' on/off true/false set/cleared constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' Account constants
Public Const NAME_LENGTH As Byte = 20
Public Const MAX_CHARS As Byte = 15

' Map constants
Public Const MAX_MAPX As Byte = 19
Public Const MAX_MAPY As Byte = 14
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
