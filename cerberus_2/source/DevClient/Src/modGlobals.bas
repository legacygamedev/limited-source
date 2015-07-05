Attribute VB_Name = "modGlobals"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

' TCP variables
Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean

' DirectX variables
Public DX As New DirectX7
Public DD As DirectDraw7
Public DD_PrimarySurf As DirectDrawSurface7
Public DD_SpriteSurf As DirectDrawSurface7
Public DD_BigSpriteSurf As DirectDrawSurface7
Public DD_TreeSpriteSurf As DirectDrawSurface7
'Public DD_BuildingSpriteSurf As DirectDrawSurface7
Public DD_TileSurf As DirectDrawSurface7
Public DD_ItemSurf As DirectDrawSurface7
Public DD_SkillSurf As DirectDrawSurface7
Public DD_SpellSurf As DirectDrawSurface7
Public DD_DirectionSurf As DirectDrawSurface7
Public DD_BackBuffer As DirectDrawSurface7
Public DD_Clip As DirectDrawClipper

Public DDSD_Primary As DDSURFACEDESC2
Public DDSD_Sprite As DDSURFACEDESC2
Public DDSD_BigSprite As DDSURFACEDESC2
Public DDSD_TreeSprite As DDSURFACEDESC2
'Public DDSD_BuildingSprite As DDSURFACEDESC2
Public DDSD_Tile As DDSURFACEDESC2
Public DDSD_Item As DDSURFACEDESC2
Public DDSD_Spell As DDSURFACEDESC2
Public DDSD_Skill As DDSURFACEDESC2
Public DDSD_Direction As DDSURFACEDESC2
Public DDSD_BackBuffer As DDSURFACEDESC2

Public rec As RECT
Public rec1 As RECT
Public rec2 As RECT
Public rec_pos As RECT

' Text variables
Public TexthDC As Long
Public GameFont As Long

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

' Used for Drag and Drop
Public MouseXOffset As Integer
Public MouseYOffset As Integer

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key open editor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Used for map PushBlock editor
Public PushDir1 As Byte
Public PushDir2 As Byte
Public PushDir3 As Byte

' Used for placing spawn points
Public RSpawnNum As Byte
Public NSpawnNum As Byte

' Map for local use
Public SaveMap As MapRec
'Public SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public SaveMapItem() As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public SaveMapResource(1 To MAX_MAP_RESOURCES) As MapResourceRec

' Used for index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSkillEditor As Boolean
Public InSpellEditor As Boolean
Public InQuestEditor As Boolean
Public InGUIEditor As Boolean
Public EditorIndex As Long

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Used for HighIndex
Public HighIndex As Long

' Maximum classes
Public Max_Classes As Byte

' Public structure variables
Public Map As MapRec
Public TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
Public PushTile(0 To MAX_MAPX, 0 To MAX_MAPY) As PushTileRec
'Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Player() As PlayerRec
Public Spells(1 To MAX_PLAYER_SPELLS) As PlayerSpellRec
Public Skills(1 To MAX_PLAYER_SKILLS) As PlayerSkillRec
Public Quests(1 To MAX_PLAYER_QUESTS) As PlayerQuestRec
Public Maps(1 To MAX_PLAYER_MAPS) As PlayerMapRec
Public Class() As ClassRec
'Public Item(1 To MAX_ITEMS) As ItemRec
Public Item() As ItemRec
'Public Npc(1 To MAX_NPCS) As NpcRec
Public Npc() As NpcRec
'Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapItem() As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public MapResource(1 To MAX_MAP_RESOURCES) As MapResourceRec
'Public Shop(1 To MAX_SHOPS) As ShopRec
Public Shop() As ShopRec
'Public Spell(1 To MAX_SPELLS) As SpellRec
Public Spell() As SpellRec
'Public Skill(1 To MAX_SKILLS) As SkillRec
Public Skill() As SkillRec
'Public Quest(1 To MAX_QUESTS) As QuestRec
Public Quest() As QuestRec
Public GUI(1 To MAX_GUIS) As GUIRec
Public Background(1 To 7) As GUIBackgroundRec
Public Menu(1 To 5) As GUIDataRec
Public Login(1 To 4) As GUIDataRec
Public NewAcc(1 To 4) As GUIDataRec
Public DelAcc(1 To 4) As GUIDataRec
Public Credits(1 To 2) As GUIDataRec
Public Chars(1 To 5) As GUIDataRec
Public NewChar(1 To 14) As GUIDataRec

