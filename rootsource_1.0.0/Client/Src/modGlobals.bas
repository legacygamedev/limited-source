Attribute VB_Name = "modGlobals"
Option Explicit

' ******************************************
' **               rootSource               **
' ******************************************

Public DX7 As New DirectX7  ' Master DirectX7 Object, early binding

' Player variables
Public MyIndex As Long ' Index of actual player
Public PlayerInv(1 To MAX_INV) As PlayerInvRec   ' Inventory
Public PlayerSpells(1 To MAX_PLAYER_SPELLS) As Byte

Public InventoryItemSelected As Integer
Public SpellSelected As Integer

Public CharSprites() As Long

' Stops movement when updating a map
Public CanMoveNow As Boolean

Public tMap(1 To 9) As Long


' Debug mode
Public DebugMode As Boolean

'Scrolling
Public NewPlayerX As Long
Public NewPlayerY As Long
Public NewXOffset As Long
Public NewYOffset As Long
Public StaticX As Long
Public StaticY As Long
Public SyncX As Long 'used to check sync
Public SyncY As Long
Public SyncMap As Long
Public SentSync As Boolean

' Controls main gameloop
Public InGame As Boolean
Public isLogging As Boolean

' Used for improved looping
Public High_Index As Integer
Public High_Npc_Index As Integer
Public PlayersOnMapHighIndex As Long
Public PlayersOnMap() As Long

' Used for dragging Picture Boxes
Public SOffsetX As Integer
Public SOffsetY As Integer

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

Public GameFPS As Long ' frames per second rendered

' Mouse cursor tile location
Public CurX As Integer
Public CurY As Integer

' Maximum classes
Public Max_Classes As Byte
