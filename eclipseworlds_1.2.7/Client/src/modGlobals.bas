Attribute VB_Name = "modGlobals"
Option Explicit

' Player constants
Public MAX_LEVEL As Long
Public MAX_STAT As Long

' Audo engine
Public Audio As New clsAudio

' The current song playing
Public CurrentMusic As String
Public CurrentMusicIndex As Long

' Paperdoll rendering order
Public PaperdollOrder() As Long

' Music & sound list cache
Public MusicCache() As String
Public SoundCache() As String
Public HasPopulated As Boolean

Public Last_Dir As Long

Public GAME_NAME As String
Public GAME_WEBSITE As String

' Animation
Public Const AnimColumns As Long = 5

' Global dialogue Index
Public DialogueIndex As Long
Public DialogueData1 As Long

' For directional blocking
Public DirArrowX(1 To 4) As Byte
Public DirArrowY(1 To 4) As Byte

' Buttons
Public LastButton_Main As Long
Public LastButton_Menu As Long
Public CurButton_Main As Long
Public CurButton_Menu As Byte

' Events
Public InEvent As Boolean
Public EventFace As Long

' Amount of blood decals
Public BloodCount As Long

' Main menu unloading
Public EnteringGame As Boolean

' GUI
Public HPBar_Width As Long
Public MPBar_Width As Long
Public EXPBar_Width As Long

' Party GUI
Public Party_HPWidth As Long
Public Party_MPWidth As Long

' Target
Public MyTarget As Byte
Public MyTargetType As Byte

' Equipment Panel
Public EquipSlotTop(1 To Equipment.Equipment_Count - 1) As Long
Public EquipSlotLeft(1 To Equipment.Equipment_Count - 1) As Long

' Animated autotiles
Public AutoAnim As Long

' Trading
Public TradeTimer As Long
Public InTrade As Long
Public TradeYourOffer(1 To MAX_INV) As PlayerItemRec
Public TradeTheirOffer(1 To MAX_INV) As PlayerItemRec
Public TradeX As Integer
Public TradeY As Integer

' Cache the Resources in an array
Public MapResource() As MapResourceRec
Public Resource_Index As Long
Public Resources_Init As Boolean

' Inventory drag and drop
Public DragInvSlot As Byte
Public InvX As Integer
Public InvY As Integer

' Bank drag and drop
Public DragBankSlot As Byte
Public BankX As Integer
Public BankY As Integer

' Spell drag and drop
Public DragSpellSlot As Byte

' Hotbar drag and drop
Public DragHotbarSlot As Byte
Public DragHotbarSpell As Byte

' GUI
Public EqX As Integer
Public EqY As Integer
Public SpellX As Integer
Public SpellY As Integer
Public ShopX As Integer
Public ShopY As Integer
Public InvItemFrame(1 To MAX_INV) As Byte ' Used for animated items
Public BankItemFrame(1 To MAX_BANK) As Byte ' Used for animated items
Public ShopItemFrame(1 To MAX_TRADES) As Byte ' Used for animated items
Public LastItemDesc As Long ' Stores the last item we showed in desc
Public LastSpellDesc As Long ' Stores the last spell we showed in desc
Public LastSpellSlotDesc As Byte
Public TmpCurrencyItem As Long
Public InShop As Long ' Is the player in a shop?
Public ShopAction As Byte ' stores the current shop action
Public TryingToFixItem As Boolean ' Stores the current shop action
Public InBank As Boolean
Public CurrencyMenu As Byte
Public InChat As Boolean

' Stops movement when updating a map
Public CanMoveNow As Boolean

' Lets the client now that everything has loaded
Public GameLoaded As Boolean

' Player variables
Public MyIndex As Long ' Index of actual player
Public PlayerInv(1 To MAX_INV) As PlayerItemRec ' Inventory
Public PlayerSpells(1 To MAX_PLAYER_SPELLS) As Long
Public InventoryItemSelected As Integer
Public SpellBuffer As Long
Public SpellBufferTimer As Long
Public SpellCD(1 To MAX_PLAYER_SPELLS) As Long
Public StunDuration As Long

' Game text Buffer
Public MyText As String

' TCP variables
Public PlayerBuffer As String

' Controls main gameloop
Public InGame As Boolean
Public IsLogging As Boolean

' Text variables
Public TexthDC As Long
Public GameFont As Long

' Draw map Name location
Public DrawMapNameX As Single
Public DrawMapNameY As Single
Public DrawMapNameColor As Long

' Game direction variables
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Mouse directions
Public MouseX As Integer
Public MouseY As Integer

' Used for dragging Picture Boxes
Public SOffsetX As Integer
Public SYOffset As Integer

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if the location/cursor/map positions need to be drawn
Public BFPS As Boolean
Public BLoc As Boolean
Public BPing As Boolean

' FPS and Time-based movement variables
Public ElapsedTime As Long
Public GameFPS As Long

' Text variables
Public vbQuote As String

' Mouse cursor tile location
Public CurX As Long
Public CurY As Long

' Game editors
Public Editor As Byte
Public EditorIndex As Long
Public EditorSave As Boolean
Public AnimEditorFrame(0 To 1) As Long
Public AnimEditorTimer(0 To 1) As Long

' Used to check if in editor or not and variables for use in editor
Public InMapEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileWidth As Long
Public EditorTileHeight As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public SpawnNPCNum As Byte
Public SpawnNPCDir As Byte
Public EditorShop As Long

' Storing width for HP/MP/Exp
Public OldHPBarWidth  As Double
Public CurrentHPBarWidth As Double
Public NewHPBarWidth As Double
Public OldMPBarWidth As Double
Public CurrentMPBarWidth  As Double
Public NewMPBarWidth  As Double
Public OldEXPBarWidth As Double
Public CurrentEXPBarWidth  As Double
Public NewEXPBarWidth  As Double
Public initHPBar  As Boolean
Public initMPBar As Boolean
Public initEXPBar   As Boolean

Public HPBarInit As Boolean
Public MPBarInit As Boolean
Public EXPBarInit As Boolean

Public HPBarWidth As Long
Public MPBarWidth As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for key on map
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for key open on map
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Used for sprite
Public MapEditorTileSprite As Integer

' Map Resources
Public ResourceEditorNum As Long

' Used for map editor heal & damage & slide & guild & on click tiles
Public MapEditorVitalType As Byte
Public MapEditorVitalAmount As Long
Public MapEditorSlideDir As Byte
Public MapEditorSound As String

' Used for map editor chat
Public MapEditorChatDir As Byte
Public MapEditorChatNPC As Long

' Camera
Public Camera As RECT
Public TileView As RECT

' Pinging
Public PingStart As Long
Public PingEnd As Long
Public Ping As Long

' Indexing
Public ActionMsgIndex As Byte
Public BloodIndex As Byte
Public ChatBubbleIndex As Byte
Public AnimationIndex As Byte

' Editor edited items array
Public Item_Changed() As Boolean
Public Quest_Changed() As Boolean
Public NPC_Changed() As Boolean
Public Resource_Changed() As Boolean
Public Animation_Changed() As Boolean
Public Spell_Changed() As Boolean
Public Shop_Changed() As Boolean
Public Ban_Changed() As Boolean
Public Title_Changed() As Boolean
Public Moral_Changed() As Boolean
Public Class_Changed() As Boolean
Public Emoticon_Changed() As Boolean

' New character
Public NewCharClass As Long

' Looping saves
Public Player_HighIndex As Byte
Public Action_HighIndex As Byte
Public Blood_HighIndex As Byte
Public ChatBubble_HighIndex As Byte

' Temp event storage
Public tmpEvent As EventRec
Public isEdit As Boolean

Public curPageNum As Long
Public curCommand As Long
Public GraphicSelX As Long
Public GraphicSelY As Long
Public GraphicSelX2 As Long
Public GraphicSelY2 As Long

Public EventTileX As Long
Public EventTileY As Long

Public EditorEvent As Long

Public GraphicSelType As Long 'Are we selecting a graphic for a move route? A page sprite? What???
Public TempMoveRouteCount As Long
Public TempMoveRoute() As MoveRouteRec
Public IsMoveRouteCommand As Boolean
Public ListOfEvents() As Long

Public EventReplyID As Long
Public EventReplyPage As Long

Public RenameType As Long
Public RenameIndex As Long
Public EventChatTimer As Long

Public AnotherChat As Long 'Determines if another showtext/showchoices is comming up, if so, dont close the event chatbox...

' Fog
Public fogOffsetX As Long
Public fogOffsetY As Long

'Weather Stuff... events take precedent OVER map settings so we will keep temp map weather settings here.
Public CurrentWeather As Long
Public CurrentWeatherIntensity As Long
Public CurrentFog As Long
Public CurrentFogSpeed As Long
Public CurrentFogOpacity As Long
Public CurrentTintR As Long
Public CurrentTintG As Long
Public CurrentTintB As Long
Public CurrentTintA As Long
Public DrawThunder As Long

' Autotiling
Public autoInner(1 To 4) As PointRec
Public autoNW(1 To 4) As PointRec
Public autoNE(1 To 4) As PointRec
Public autoSW(1 To 4) As PointRec
Public autoSE(1 To 4) As PointRec

' Map animations
Public waterfallFrame As Long
Public autoTileFrame As Long

' Chat bubble
Public ChatBubble(1 To MAX_BYTE) As ChatBubbleRec

' Hotbar
Public Hotbar(1 To MAX_HOTBAR) As HotbarRec

' Swear filter
Public SwearArray() As String
Public ReplaceSwearArray() As String

' Chat
Public ChatLocked As Boolean
Public CurrentChatChannel As Long

' Character Creation Arrays
Public ClassSelection() As Byte

' Option Buttons
Public OptionButton(1 To OptionButtons.Opt_Count - 1) As ButtonRec

' GUI Toggles
Public GUIVisible As Boolean
Public ButtonsVisible As Boolean

Public FadeType As Long
Public FadeAmount As Long
Public FlashTimer As Long

Public hwndLastActiveWnd As Long

' Data Sizes
Public MAX_MAPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_ANIMATIONS As Long
Public MAX_SHOPS As Long
Public MAX_SPELLS As Long
Public MAX_RESOURCES As Long
Public MAX_QUESTS As Long
Public MAX_BANS As Long
Public MAX_TITLES As Long
Public MAX_MORALS As Long
Public MAX_CLASSES As Long
Public MAX_EMOTICONS As Long
