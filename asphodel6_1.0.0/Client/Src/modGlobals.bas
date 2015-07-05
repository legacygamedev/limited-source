Attribute VB_Name = "modGlobals"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

' Number of tiles in width
Public TILESHEET_WIDTH() As Integer

' Config variables
Public Game_Name As String
Public GAME_IP As String
Public GAME_PORT As Integer
Public GAME_WEBSITE As String
Public Password_Confirmed As Boolean
Public Sprite_Offset As Byte
Public Remember As Boolean
Public Music_On As Boolean
Public Sound_On As Boolean
Public FPS_Lock As Integer
Public TOTAL_SPRITES As Long
Public TOTAL_ANIMGFX As Long
Public Config_Received As Boolean
Public Total_SpriteFrames As Integer
Public Total_AnimFrames As Byte
Public MAX_PLAYERS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_SPELLS As Long
Public MAX_SHOPS As Long
Public MAX_MAPS As Long
Public MAX_SIGNS As Long
Public MAX_ANIMS As Long
Public Selected_Spell As Long
Public CurrentClassName As String
Public StatBuffed(1 To Stats.Stat_Count - 1) As Long
Public SignSection() As String
Public PingCounter As Currency
Public CurPing As Long
Public PingEnabled As Byte
Public WaitingonPing As Boolean
Public CurrentWindow As Long
Public Windows(0 To Window_State.Count - 1) As Form
Public MapAttribType As Long
Public MapAttribFormTitle As String
Public MapAttribData(1 To 3) As Long
Public MapAttribName(1 To 3) As String
Public MapAttribMin(1 To 3) As Long
Public MapAttribMax(1 To 3) As Long
Public SettingSpawn As Boolean
Public MapSpawnX() As Long
Public MapSpawnY() As Long
Public CheckedStuff As Boolean
Public CheckedTwice As Boolean
Public Char_Selected As Long
Public Char_Sprite(1 To 3) As Long
Public CharIsThere(1 To 3) As Boolean

' decide to show player or NPC names
Public ShowPNames As Boolean
Public ShowNNames As Boolean

' Lookup tables
Public MultiplyPicX(0 To MAX_BYTE) As Integer
Public ColorTable(0 To Color.Count - 1) As Long
Public ModularTable(0 To MAX_INTEGER, 0 To MAX_BYTE) As Integer

' Use for visual inventory and spells
Public IconPosX As Single
Public IconPosY As Single
Public InvPosX As Single
Public InvPosY As Single
Public ShopPosX As Single
Public ShopPosY As Single

' Used for repairing and selling
Public ReadyToRepair As Boolean
Public ReadyToSell As Boolean

' Used for MovePicture to drag picture boxes
Public SOffsetX As Integer
Public SOffsetY As Integer

' TCP variables
Public PlayerBuffer As String

' Controls main gameloop
Public InGame As Boolean
Public isLogging As Boolean

' Text variables
Public TexthDC As Long
Public GameFont As Long

' Draw map name location
Public DrawMapNameX As Single
Public DrawMapNameY As Single
Public DrawMapNameColor As Long

' Stops movement when updating a map
Public CanMoveNow As Boolean

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Game text buffer
Public MyText As String

' Index of actual player
Public MyIndex As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Currency

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if FPS needs to be drawn
Public BFPS As Boolean
Public BLoc As Boolean

' frames per second rendered
Public GameFPS As Long

' Text vars
Public vbQuote As String

' Mouse cursor tile location
Public CurX As Integer
Public CurY As Integer

' Game editors
Public Editor As Byte
Public EditorIndex As Long

Public MAX_TILESETS As Long
Public MAX_ITEMSETS As Long

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileX2 As Long
Public EditorTileY2 As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public EditorMapMusic As String
Public EditorNpcAttackSound As String
Public EditorNpcSpawnSound As String
Public EditorNpcDeathSound As String
Public EditorSpellSound As String
Public EditorSignNum As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long
Public ItemEditorAnim As Long

' Used for map shop editor
Public ShopEditorNum As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key open editor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Used for anim editor
Public AnimEditorAnim As Long

' Used for parsing String packets
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte
