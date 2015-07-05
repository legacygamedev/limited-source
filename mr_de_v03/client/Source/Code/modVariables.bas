Attribute VB_Name = "modVariables"
Option Explicit
'//////////////
'// Declares //
'//////////////
'Text Declares
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

' Sound Declares
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Game Declares
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)

Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
  
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

'///////////////
'// Constants //
'///////////////

Public Const GAME_IP As String = "127.0.0.1"
Public Const GAME_PORT As Long = 4000

Public Const GAME_NAME As String = "Mirage Realms Developer Editions"
Public Const GAME_WEBSITE As String = "www.mirage-realms.com"
Public Const FONT_TYPE As String = "fixedsys"
Public Const FONT_SIZE As Byte = 18
Public Const MAX_PLAYERS As Byte = 100
Public Const MAX_ITEMS As Integer = 1000
Public Const MAX_NPCS As Integer = 1000
Public Const MAX_SHOPS As Integer = 1000
Public Const MAX_SPELLS As Integer = 1000
Public Const MAX_EMOTICONS As Byte = 255
Public Const MAX_ANIMATIONS As Byte = 255
Public Const MAX_INV As Byte = 16
Public Const MAX_MAP_ITEMS As Integer = 20
Public Const MAX_PLAYER_SPELLS As Byte = 20
Public Const MAX_TRADES As Byte = 12
Public Const MAX_MOBS As Byte = 20

Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

Public Const NO As Byte = 0
Public Const YES As Byte = 1
Public Const OKAY As Byte = 2

' Account constants
Public Const NAME_LENGTH As Byte = 20
Public Const MAX_CHARS As Byte = 3

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public Const MAX_MAPS As Integer = 1000
Public Const MAX_MAPX As Byte = 15
Public Const MAX_MAPY As Byte = 11
Public Const MAP_MORAL_SAFE As Byte = 0
Public Const MAP_MORAL_NONE As Byte = 1

' Image constants
Public Const PIC_X As Integer = 32
Public Const PIC_Y As Integer = 32

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_MOBSPAWN As Byte = 7

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_EQUIPMENT As Byte = 1
Public Const ITEM_TYPE_POTION As Byte = 2
Public Const ITEM_TYPE_KEY As Byte = 3
Public Const ITEM_TYPE_SPELL As Byte = 4

' Shop Constants
Public Const SHOP_TYPE_SHOP As Byte = 0
Public Const SHOP_TYPE_INN As Byte = 1

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' Weather constants
Public Const WEATHER_NONE As Byte = 0
Public Const WEATHER_RAINING As Byte = 1
Public Const WEATHER_SNOWING As Byte = 2

' Time constants
Public Const TIME_DAY As Byte = 0
Public Const TIME_NIGHT As Byte = 1

' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOR_QUEST As Byte = 4

' Spell constants
Public Const SPELL_TYPE_VITAL As Byte = 0
Public Const SPELL_TYPE_OVERTIME As Byte = 1
Public Const SPELL_TYPE_BUFF As Byte = 2
Public Const SPELL_TYPE_REVIVE As Byte = 3

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

'Text Variables
Public Const Quote = """"

Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15

Public Const SayColor As Byte = Grey
Public Const TellColor As Byte = Cyan
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = BrightCyan
Public Const NewMapColor As Byte = Pink

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10


Public Const VK_UP = &H26
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_ALT = &H12
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
'Public Const MOVING_RUNNING = 2

' Speed moving vars
Public Const WALK_Speed As Byte = 4
'Public Const RUN_Speed = 8
Public Const NPC_Speed As Byte = 4
Public Const NPC_Speed_FAST As Byte = 8
Public Const NPC_Speed_FASTEST As Byte = 16

' Consts for visual inventory
Public Const InvLeft As Byte = 11
Public Const InvTop As Byte = 30
Public Const InvOffsetX As Byte = 17
Public Const InvOffsetY As Byte = 16
Public Const InvColumns As Byte = 4

' Consts for Action Msgs
Public Const ACTIONMSG_STATIC As Byte = 0
Public Const ACTIONMSG_SCROLL As Byte = 1
Public Const ACTIONMSG_SCREEN As Byte = 2

'//////////////////////
'// Public Variables //
'//////////////////////

' Used for faster loops
Public MapPlayers() As Long
Public MapPlayersCount As Long

' PM variable
Public PMName As String

Public TexthDC As Long
Public GameFont As Long

' Sound Variables
Public CurrentSong As String

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public AltDown As Boolean
Public ControlDown As Boolean

' Game text buffer
Public MyText As String

' Index of actual player
Public MyIndex As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Boolean
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
Public ItemEditorBlocked As Byte

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key opene ditor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long
Public KeyOpenPressure As Byte

' Used for spawn npc editor
Public SpawnNpcNum As Long
Public SpawnNpcNumDir As Byte

' Used for Index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InEmoticonEditor As Boolean
Public InAnimationEditor As Boolean
Public InQuestEditor As Boolean
Public EditorIndex As Long

'Location of pointer on picscreen
Public CurX As Integer
Public CurY As Integer
Public MouseX As Integer
Public MouseY As Integer

' Target
Public MyTarget As Byte
Public MyTargetType As Byte

'Visual Inventory
Public InvVisible As Boolean
Public DragInvSlotNum As Byte
Public DropNum As Byte
Public InvX As Long
Public InvY As Long
Public TradeX As Long
Public TradeY As Long

' TCP
Public PlayerBuffer As clsBuffer

Public InGame As Boolean

' For action Messages
Public ActionMsgIndex As Byte
Public AnimIndex As Byte

' For the shop system
Public InShop As Byte
Public ShopNpcNum As Byte

' Used for sprite selection
Public spriteMale() As String
Public spriteFemale() As String
Public currSpriteNum As Long

Public EncryptPackets As Byte           ' Flag for encrypting packets - 1 for Yes - 0 for No
Public PacketInIndex As Byte            ' Holds the Index of what packetkey for incoming packets
Public PacketOutIndex As Byte           ' Holds the Index of what packetkey for outgoing packets
Public PacketKeys() As String           ' Holds our array of packet encryption keys

Public CastingSpell As Long
Public CastTime As Long

Public LockFPS As Boolean
Public ShowFPS As Boolean
Public GameFPS As Long

' Config options
Public ShowItemLinks As Boolean

' For item hover caches
Public ItemReq(1 To MAX_ITEMS) As String
Public ItemDesc(1 To MAX_ITEMS) As String

Public Camera As RECT
Public TileView As RECT
Public Const HalfX As Integer = ((MAX_MAPX + 1) / 2) * PIC_X
Public Const HalfY As Integer = ((MAX_MAPY + 1) / 2) * PIC_Y
Public Const ScreenX As Integer = (MAX_MAPX + 1) * PIC_X
Public Const ScreenY As Integer = (MAX_MAPY + 1) * PIC_Y
Public Const StartXValue As Integer = ((MAX_MAPX + 1) / 2)
Public Const StartYValue As Integer = ((MAX_MAPY + 1) / 2)
Public Const EndXValue As Integer = (MAX_MAPX + 1) + 1
Public Const EndYValue As Integer = (MAX_MAPY + 1) + 1
Public Const Half_PIC_X As Integer = PIC_X / 2
Public Const Half_PIC_Y As Integer = PIC_Y / 2

Public RandomTile(0 To 3) As Integer
Public RandomTileSelected As Byte

Public TILE_WIDTH As Long

' States
Public CurrentState As MenuStates

Public MapNpcCount As Long

Public HotBarSpell(0 To 3) As Long   ' Will hold the spellslot number

' Below are used to track the size of the specified UDT
' Set in game init
Public AnimationSize As Long    ' Animation UDT
Public EmoticonSize As Long     ' Emoticon UDT
Public ItemSize As Long         ' Item UDT
Public NpcSize As Long          ' Npc UDT
Public ShopSize As Long         ' Shop UDT
Public SpellSize As Long        ' Spell UDT
Public QuestSize As Long        ' Quest UDT

Public MAX_CLASSES As Byte


'//////////////////
'// Public Enums //
'//////////////////

' Menu states
Public Enum MenuStates
    MainMenu = 0
    NewAccount
    Login
    GetChars
    NewChar
    AddChar
    DelChar
    UseChar
    Chars
    Shutdown
End Enum

Public Enum Stats
    Strength = 1
    Dexterity
    Vitality
    Intelligence
    Wisdom
    ' Make sure Stat_Count is below everything else
    Stat_Count = Wisdom
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    SP
    ' Mak sure Vital_Count is below everything else
    Vital_Count = SP
End Enum

Public Enum Slots
    Weapon = 1
    Armor
    Helmet
    Shield
    ' Make sure Slot_Count is below everything else
    Slot_Count = Shield
End Enum

Public Enum Targets
    Target_None = 1
    Target_SelfOnly = 2    ' Will only be cast on you - no matter what other flags are set
    Target_PlayerHostile = 4    ' Can be cast on other players - on pvp map
    Target_PlayerBeneficial = 8 ' Can be cast on other players - on any map
    Target_Npc = 16             ' Can be cast on all npcs
    Target_PlayerParty = 32     ' Can only be cast on party members - Will override other flags except selfonly
    ' Make sure Target_Count is below everything else
    Target_Count = 6
End Enum

Public Enum ItemBind
    None = 0
    BindOnEquip
    BindOnPickUp
    ItemBind_Count = BindOnPickUp
End Enum

'*********************
'   For Packet Data
'*********************
Public Enum CMsgTypes                   ' Server -> Client // Make sure it's the same for the server
    CMsgAlertMsg = 1
    CMsgClientMsg
    CMsgAllChars
    CMsgLoginOk
    CMsgNewCharClasses
    CMsgClassesData
    CMsgInGame
    CMsgPlayerInv
    CMsgPlayerInvUpdate
    CMsgPlayerWornEq
    CMsgPlayerVital
    CMsgPlayerStats
    CMsgPlayerData
    CMsgPlayerMove
    CMsgNpcMove
    CMsgPlayerDir
    CMsgNpcDir
    CMsgPlayerXY
    CMsgAttack
    CMsgNpcAttack
    CMsgCheckForMap
    CMsgMapData
    CMsgMapItemData
    CMsgMapNpcData
    CMsgMapDone
    CMsgChatMsg
    CMsgSpawnItem
    CMsgItemEditor
    CMsgUpdateItem
    CMsgUpdateItems
    CMsgEditItem
    CMsgEditEmoticon
    CMsgUpdateEmoticon
    CMsgUpdateEmoticons
    CMsgEmoticonEditor
    CMsgCheckEmoticon
    CMsgNewTarget
    CMsgSpawnNpc
    CMsgNpcDead
    CMsgNpcEditor
    CMsgUpdateNpc
    CMsgUpdateNpcs
    CMsgEditNpc
    CMsgMapKey
    CMsgEditMap
    CMsgShopEditor
    CMsgUpdateShop
    CMsgUpdateShops
    CMsgEditShop
    CMsgSpellEditor
    CMsgUpdateSpell
    CMsgUpdateSpells
    CMsgEditSpell
    CMsgAnimationEditor
    CMsgUpdateAnimation
    CMsgUpdateAnimations
    CMsgEditAnimation
    CMsgTrade
    CMsgSpells
    CMsgActionMsg
    CMsgAnimation
    CMsgPlayerGuild
    CMsgPlayerExp
    CMsgCancelSpell
    CMsgSpellReady
    CMsgSpellCooldown
    CMsgLeftGame
    CMsgPlayerDead
    CMsgPlayerGold
    CMsgPlayerRevival
    ' Quest
    CMsgQuestEditor
    CMsgUpdateQuest
    CMsgUpdateQuests
    CMsgEditQuest
    CMsgAvailableQuests
    CMsgPlayerQuests
    CMsgPlayerQuest
    'The following enum member automatically stores the number of messages,
    'since it is last. Any new messages must be placed above this entry.
    CMSG_COUNT
End Enum

Public Enum SMsgTypes                   ' Client -> Server // Make sure it's the same for the server
    SMsgGetClasses = 1
    SMsgNewAccount
    SMsgLogin
    SMsgRequestEditEmoticon
    SMsgEditEmoticon
    SMsgSaveEmoticon
    SMsgCheckEmoticon
    SMsgAddChar
    SMsgDelChar
    SMsgUseChar
    SMsgSayMsg
    SMsgEmoteMsg
    SMsgGlobalMsg
    SMsgAdminMsg
    SMsgPartyMsg
    SMsgPlayerMsg
    SMsgPlayerMove
    SMsgPlayerDir
    SMsgUseItem
    SMsgUnequipSlot
    SMsgAttack
    SMsgUseStatPoint
    SMsgPlayerInfoRequest
    SMsgWarpMeTo
    SMsgWarpToMe
    SMsgWarpTo
    SMsgSetSprite
    SMsgGetStats
    SMsgClickWarp
    SMsgRequestNewMap
    SMsgMapData
    SMsgNeedMap
    SMsgMapGetItem
    SMsgMapDropItem
    SMsgMapRespawn
    SMsgMapReport
    SMsgKickPlayer
    SMsgListBans
    SMsgBanDestroy
    SMsgBanPlayer
    SMsgRequestEditMap
    SMsgRequestEditItem
    SMsgEditItem
    SMsgSaveItem
    SMsgRequestEditNpc
    SMsgEditNpc
    SMsgSaveNpc
    SMsgRequestEditShop
    SMsgEditShop
    SMsgSaveShop
    SMsgRequestEditSpell
    SMsgEditSpell
    SMsgSaveSpell
    SMsgRequestEditAnimation
    SMsgEditAnimation
    SMsgSaveAnimation
    SMsgSetAccess
    SMsgWhosOnline
    SMsgSetMOTD
    SMsgTradeRequest
    SMsgSearch
    SMsgParty
    SMsgJoinParty
    SMsgLeaveParty
    SMsgCast
    SMsgRequestLocation
    SMsgFix
    SMsgChangeInvSlots
    SMsgClearTarget
    SMsgGCreate
    SMsgSetGMOTD
    SMsgGQuit
    SMsgGDelete
    SMsgGPromote
    SMsgGDemote
    SMsgGKick
    SMsgGInvite
    SMsgGJoin
    SMsgGDecline
    SMsgGuildMsg
    SMsgKill
    SMsgSetBound
    SMsgCancelSpell
    SMsgRelease
    SMsgRevive
    SMsgRequestEditQuest
    SMsgEditQuest
    SMsgSaveQuest
    SMsgAcceptQuest
    SMsgCompleteQuest
    SMsgDropQuest
    'The following enum member automatically stores the number of messages,
    'since it is last. Any new messages must be placed above this entry.
    SMSG_COUNT
End Enum

' Has to be below the enums
Public HandleDataSub(CMSG_COUNT) As Long

