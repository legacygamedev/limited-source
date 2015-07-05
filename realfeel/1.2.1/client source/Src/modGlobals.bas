Attribute VB_Name = "modGlobals"
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Created module.
'****************************************************************

Option Explicit
'Game Name Global
Public GAME_NAME As String

' Winsock globals
Public GAME_IP As String
Public GAME_PORT As Integer

' TCP variables
Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean

'--------------------------------------------
'Added the DirectX8 object (04/21/07)
'Added the DirectMusic objects (04/21/07)
'Added the DirectSound objects (04/23/07)
'Added the Direct3D objects (12/29/07)
'-smchronos

' DirectX8 object
Public DX8 As New DirectX8
Public Direct3D As clsDirect3D
Public DirectMusic As clsDirectMusic
Public DirectSound As clsDirectSound
Public DirectShow As clsDirectShow

'DirectX8 3D objects
'Public backbuffer As Direct3DSurface8
Public spritetex As Direct3DTexture8
Public itemtex As Direct3DTexture8
Public tiletex As Direct3DTexture8
Public mapeditortex As Direct3DTexture8

'DirectX8 Sound object

'--------------------------------------------

' Text variables
Public TexthDC As Long
Public GameFont As Long

' Max data variables
Public MAX_PLAYERS As Long
Public MAX_MAPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_SHOPS As Long
Public MAX_SPELLS As Long

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' String vars
Public STRING_HP As String
Public STRING_MP As String
Public STRING_SP As String
Public STRING_STRENGTH As String
Public STRING_DEFENSE As String
Public STRING_MAGIC As String
Public STRING_SPEED As String

' Game text buffer
'Public MyText As String

' Index of actual player
Public MyIndex As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileXEnd As Long
Public EditorTileYEnd As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key opene ditor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Map for local use
Public SaveMap As MapRec
Public SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public EditorIndex As Long

' Game fps
Public GameFPS As Long

' Used for atmosphere
Public GameWeather As Long
Public GameTime As Long

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte
Public Max_Visible_Classes As Byte

' Public structure variables
Public Map As MapRec
Public TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
Public Player() As PlayerRec
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec

'Multi-Tile Selection variables
Public KeyShift As Boolean

'Mouse Variables
Public SOffsetX As Integer
Public SOffsetY As Integer
Public Mouse_X As Long
Public Mouse_Y As Long

'Picture dimensions
Public PicSize As BITMAPINFO

' Visual Item Tracker
Public TopBank As Byte
Public BottomBank As Byte
Public TopInv As Byte
Public BottomInv As Byte
Public ItemSelected As Byte
Public BankSelected As Byte
Public InvSelected As Byte

'tracker name
Public TrackName As String

'shop number for tile
Public EditorShopNum As Long

'health values for tile
Public EditorHealValue As Long
Public EditorDamageValue As Long

' DLL Path
Public DLL_PATH As String

' Log Path
Public LOG_PATH As String

' Map Path
Public MAP_PATH As String

' Gfx Path
Public GFX_PATH As String

' Sound Path
Public SOUND_PATH As String

' Music Path
Public MUSIC_PATH As String

' GUI Path
Public GUI_PATH As String

' This is used to pause the current map
Public PauseMap As Boolean
Public PauseMessage As String

' This is used for flexibility in mapping
Public AllowMovement As Boolean
Public AttributeDisplay As Boolean
Public DepictAttributeTiles As Boolean

' Global Direct3D colors
Public C_Black As Long
Public C_Blue As Long
Public C_Green As Long
Public C_Cyan As Long
Public C_Red As Long
Public C_Magenta As Long
Public C_Brown As Long
Public C_Grey As Long
Public C_DarkGrey As Long
Public C_BrightBlue As Long
Public C_BrightGreen As Long
Public C_BrightCyan As Long
Public C_BrightRed As Long
Public C_Pink As Long
Public C_Yellow As Long
Public C_White As Long

' Check if we need to continue looping
Public LoopIntro As Byte
