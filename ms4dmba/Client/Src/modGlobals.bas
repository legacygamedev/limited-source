Attribute VB_Name = "modGlobals"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

' Player variables
Public MyIndex As Long ' Index of actual player
Public PlayerInv(1 To MAX_INV) As PlayerInvRec   ' Inventory
Public PlayerSpells(1 To MAX_PLAYER_SPELLS) As Byte

Public InventoryItemSelected As Integer
Public SpellSelected As Integer

' Stops movement when updating a map
Public CanMoveNow As Boolean

' Debug mode
Public DEBUG_MODE As Boolean

' Game text buffer
Public MyText As String

' TCP variables
Public PlayerBuffer As String

' Used for parsing String packets
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Controls main gameloop
Public InGame As Boolean
Public isLogging As Boolean

' Used for improved looping
Public High_Index As Integer
Public High_Npc_Index As Long
Public PlayersOnMapHighIndex As Long
Public PlayersOnMap() As Long

' Text variables
Public TexthDC As Long
Public GameFont As Long

' Draw map name location
Public DrawMapNameX As Single
Public DrawMapNameY As Single
Public DrawMapNameColor As Long

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Used for dragging Picture Boxes
Public SOffsetX As Integer
Public SOffsetY As Integer

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if FPS needs to be drawn
Public BFPS As Boolean
Public BLoc As Boolean

Public GameFPS As Long ' frames per second rendered

' Text vars
Public vbQuote As String

' Mouse cursor tile location
Public CurX As Integer
Public CurY As Integer

' Game editors
Public Editor As Byte
Public EditorIndex As Long

' Used to check if in editor or not and variables for use in editor
Public InMapEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key open editor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Maximum classes
Public Max_Classes As Byte


Public Camera As RECT
Public TileView As RECT

