Attribute VB_Name = "modGlobals"
Option Explicit

'Store walkfix setting
Public WalkFix As Long

'Public declarations of variables holding songs, samples and streams
Public MapSound As String
Public Sounds As String
Public BGSound As String

' Position buffer
Public NewPosX As Long
Public NewPosY As Long
Public IsNewPos As Boolean

' mouse cursor location
Public CurX As Long
Public CurY As Long

Public snumber As Integer

' Game text buffer
Public MyText As String

' Index of actual player
Public MyIndex As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' For Map editor
Public ScreenMode As Byte
Public NightMode As Byte
Public GridMode As Byte
Public MapEditorSelectedType As Byte

'House editor
Public HouseEditorSelectedType As Byte
Public InHouseEditor As Boolean

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public EditorSet As Byte

' Camera globals
Public ScreenX As Long
Public ScreenY As Long
Public ScreenX2 As Long
Public ScreenY2 As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key open editor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long
Public KeyOpenEditorMsg As String

' Map for local use
Public SaveMapItem() As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InElementEditor As Boolean
Public InEmoticonEditor As Boolean
Public InArrowEditor As Boolean
Public EditorIndex As Long

' Game fps
Public GameFPS As Long
Public BFPS As Boolean

' Used for atmosphere
Public GameWeather As Long
Public GameTime As Long
Public RainIntensity As Long

' Scrolling Variables
Public NewPlayerX As Long
Public NewPlayerY As Long
Public NewXOffset As Long
Public NewYOffset As Long
Public NewX As Long
Public NewY As Long

' Damage Variables
Public DmgDamage As Long
Public DmgTime As Long
Public NPCDmgDamage As Long
Public NPCDmgTime As Long
Public NPCWho As Long

Public EditorItemX As Long
Public EditorItemY As Long

Public EditorShopNum As Long

Public EditorItemNum1 As Byte
Public EditorItemNum2 As Byte
Public EditorItemNum3 As Byte

Public Arena1 As Byte
Public Arena2 As Byte
Public Arena3 As Byte

Public ii As Long, iii As Long
Public sx As Long

Public SpritePic As Long
Public SpriteItem As Long
Public SpritePrice As Long

Public HouseItem As Long
Public HousePrice As Long

Public SoundFileName As String


Public SignLine1 As String
Public SignLine2 As String
Public SignLine3 As String

Public ClassChange As Long
Public ClassChangeReq As Long

Public NoticeTitle As String
Public NoticeText As String
Public NoticeSound As String

Public ScriptNum As Long

Public Connected As Boolean

' Used for NPC spawn
Public NPCSpawnNum As Long

' Used for roof tile
Public RoofId As String

Public AutoLogin As Long

' Used to make sure we have all the data before logging in
Public AllDataReceived As Boolean

' Used for classes
Public ClassesOn As Byte

' Last Direction
Public LAST_DIR As Long

' Keep track of time
Public Hours As Integer
Public Minutes As Integer
Public Seconds As Integer
Public Gamespeed As Integer

' Font data
Public Font As String
Public fontsize As Byte

Public SOffsetX As Integer
Public SOffsetY As Integer

Public BLoc As Boolean

Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean
Public TradePlayer As Long

Public TexthDC As Long
Public GameFont As Long

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean


' General constants
Public GAME_NAME As String
Public WEBSITE As String
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_ELEMENTS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_EMOTICONS As Long
Public MAX_SPELL_ANIM As Long
Public MAX_BLT_LINE As Long
Public paperdoll As Long
Public SpriteSize As Long
Public MAX_SCRIPTSPELLS As Long
Public CustomPlayers As Long
Public CUSTOM_TITLE As String
Public CUSTOM_IS_CLOSABLE As Long
Public MAX_PARTY_MEMBERS As Long
Public temp As Long
Public lvl As Long
Public STAT1 As String
Public STAT2 As String
Public STAT3 As String
Public STAT4 As String

Public Anim1Data As Long
Public Anim2Data As Long
Public M2AnimData As Long
Public FAnimData As Long
Public F2AnimData As Long

' OnClick tile info
Public ClickScript As Integer

' Map constants
' Public Const MAX_MAPX = 30
' Public Const MAX_MAPY = 30
Public MAX_MAPX As Long
Public MAX_MAPY As Long

' Minus Stat values
Public MinusHp As Integer
Public MinusMp As Integer
Public MinusSp As Integer
Public MessageMinus As String

' Playing sound
Public CurrentSong As String
Public CurrentSound As String
Public MapMusicStarted As Boolean

' Bubble thing
Public Bubble() As ChatBubble
Public ScriptBubble() As ScriptBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

Public Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public Trading2(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

Public Map() As MapRec
Public TempTile() As TempTileRec
Public Player() As PlayerRec
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Element() As ElementRec
Public Emoticons() As EmoRec
Public MapReport() As MapRec
Public ScriptSpell() As ScriptSpellAnimRec

Public MAX_RAINDROPS As Long
Public BLT_RAIN_DROPS As Long
Public DropRain() As DropRainRec

Public BLT_SNOW_DROPS As Long
Public DropSnow() As DropRainRec
Public Trade(1 To 7) As TradeRec
Public Arrows(1 To MAX_ARROWS) As ArrowRec

Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Public ItemDur(1 To 4) As ItemDurRec

Public Inventory As Long
Public slot As Long

Public Direct As Long
Public GuildBlock As String

Public SpellMemorized As Long
