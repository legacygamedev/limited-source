Attribute VB_Name = "modGlobals"
Option Explicit
'DX Globals
Public DX As New DirectX7
Public DD As DirectDraw7
Public DD_PrimarySurf As DirectDrawSurface7
Public DD_SpriteSurf As DirectDrawSurface7
Public DD_TileSurf As DirectDrawSurface7
Public DD_ItemSurf As DirectDrawSurface7
Public DD_BackBuffer As DirectDrawSurface7
Public DD_Clip As DirectDrawClipper

Public DDSD_Primary As DDSURFACEDESC2
Public DDSD_Sprite As DDSURFACEDESC2
Public DDSD_Tile As DDSURFACEDESC2
Public DDSD_Item As DDSURFACEDESC2
Public DDSD_BackBuffer As DDSURFACEDESC2

Public rec As RECT
Public rec_pos As RECT

'game connection/player buffer
Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean

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
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
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

' Used for map key opene ditor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Map for local use
Public SaveMap As MapRec
Public SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean

Public EditorIndex As Long

'Public offsets
Public SOffsetX As Integer
Public SOffsetY As Integer

' Game fps
Public GameFPS As Long

' Used for atmosphere
Public GameWeather As Long
Public GameTime As Long

'text globals
Public TexthDC As Long
Public GameFont As Long

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

Public Map As MapRec
Public TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec


Public Sub CharMove(ByVal X As Single, ByVal Y As Single)
'Finds the heading way with our mouse position
Dim iScrX As Integer
Dim iScrY As Integer
Dim lAngle As Long

iScrX = X - frmMirage.picScreen.Left - Player(MyIndex).X * 32
iScrY = Y - frmMirage.picScreen.top - Player(MyIndex).Y * 32
iScrY = -iScrY
If iScrY = 0 Then
    lAngle = 0
Else
    lAngle = Atn(iScrX / iScrY) * 180 / 3.14159265
End If

If (lAngle >= -45 And lAngle <= 0) Or (lAngle <= 45 And lAngle >= 0) Then
    If iScrY > 0 Then
        DirUp = True
        DirDown = False
        DirLeft = False
        DirRight = False
        If CanMove = True Then
             Call SetPlayerDir(MyIndex, DIR_UP)
             Call CheckMovement
        End If
    Else
        DirUp = False
        DirDown = True
        DirLeft = False
        DirRight = False
        If CanMove = True Then
             Call SetPlayerDir(MyIndex, DIR_DOWN)
             Call CheckMovement
        End If
    End If
ElseIf (lAngle > 45 And lAngle <= 90) Or (lAngle < -45 And lAngle >= -90) Then
    If iScrX < 0 Then
                 DirUp = False
        DirDown = False
        DirLeft = True
        DirRight = False
        If CanMove = True Then
             Call SetPlayerDir(MyIndex, DIR_LEFT)
             Call CheckMovement
        End If
    Else
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = True
        If CanMove = True Then
             Call SetPlayerDir(MyIndex, DIR_RIGHT)
             Call CheckMovement
        End If
    End If
End If
End Sub
