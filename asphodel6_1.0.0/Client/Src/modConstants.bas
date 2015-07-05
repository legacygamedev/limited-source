Attribute VB_Name = "modConstants"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

' used for transparent text boxes
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TRANSPARENT = &H20&

' max variable type values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767

' the default width of frmmaingame and width when map editor is opened
Public Const Default_MenuWidth As Long = 13500
Public Const MenuWidth_withEditor As Long = 17445

' used for TNL, HP, and MP Bars
Public Const TNLBar_Width As Long = 644
Public Const HPMPBar_Width As Long = 160

' Visual spell icons
Public Const IconX As Byte = 2
Public Const IconY As Byte = 2
Public Const IconOffsetX As Byte = 2
Public Const IconOffsetY As Byte = 2
Public Const IconsInRow As Byte = 1

' Visual inventory icons
Public Const ItemIconX As Byte = 1
Public Const ItemIconY As Byte = 1
Public Const ItemOffsetX As Byte = 1
Public Const ItemOffsetY As Byte = 1
Public Const ItemsInRow As Byte = 4

' Visual shop icons
Public Const ShopIconX As Byte = 1
Public Const ShopIconY As Byte = 1
Public Const ShopOffsetX As Byte = 1
Public Const ShopOffsetY As Byte = 1
Public Const ShopIconsInRow As Byte = 8

' Debug mode
Public Const DEBUG_MODE As Boolean = False

' path constants
Public Const SOUND_PATH As String = "\sound\"
Public Const MUSIC_PATH As String = "\music\"
Public Const MUSIC_EXT As String = ".mid"
Public Const SOUND_EXT As String = ".wav"

' Font variables
Public Const FONT_NAME As String = "fixedsys"
Public Const FONT_SIZE As Byte = 18

' Config file name and key for IP
' (needs to match key in config file maker)
Public Const CONFIG_FILE As String = "config.aph"
Public Const DEFAULT_KEY As String = "admin"

' Log Path and variables
Public Const LOG_DEBUG As String = "debug.txt"
Public Const LOG_PATH As String = "\logs\"

' Map Path and variables
Public Const MAP_PATH As String = "\maps\"
Public Const MAP_EXT As String = ".map"

' Gfx Path and variables
Public Const GFX_PATH As String = "\graphics\"
Public Const GFX_EXT As String = ".bmp"

' Key constants
Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

' Font sizes
Public Const FONT_WIDTH As Byte = 8
Public Const FONT_HEIGHT As Byte = 7

' Speed moving vars
Public Const WALK_SPEED As Byte = 2
Public Const RUN_SPEED As Byte = 3

' Image constants
Public Const PIC_X As Integer = 32
Public Const PIC_Y As Integer = 32

' Height of spell icons
Public Const SpellIconHeight As Integer = 32

' General constants
Public Const MAX_INV As Byte = 24
Public Const MAX_MAP_ITEMS As Byte = 20
Public MAX_MAP_NPCS As Byte
Public Const MAX_PLAYER_SPELLS As Byte = 10
Public Const MAX_TRADES As Byte = 64

' Text color constants
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
Public Const AlertColor As Byte = Color.red
Public Const NewMapColor As Byte = Color.Pink

' Yes and No constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' Account constants
Public Const NAME_LENGTH As Byte = 20
Public Const MAX_CHARS As Byte = 3

' Map constants
Public Const MAX_MAPX As Byte = 19
Public Const MAX_MAPY As Byte = 14
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1

