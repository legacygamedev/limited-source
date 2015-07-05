Attribute VB_Name = "modTypes"

' Copyright (c) 2008 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit
Public HelmetLogin(1 To 3) As Integer
Public LegsLogin(1 To 3) As Integer
Public ArmorLogin(1 To 3) As Integer
Public WeaponLogin(1 To 3) As Integer
Public ShieldLogin(1 To 3) As Integer
Public DropIndex As Long
Public Const MAX_PARTY_MEMBERS As Byte = 4
Public Const MAX_PARTY_INV_SLOTS As Byte = 10
Public InQuestMenu As Byte
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Const Quote = """"
Public Const Black = 0
Public Const Blue = 1
Public Const Green = 2
Public Const Cyan = 3
Public Const Red = 4
Public Const Magenta = 5
Public Const Brown = 6
Public Const Grey = 7
Public Const DarkGrey = 8
Public Const BrightBlue = 9
Public Const BrightGreen = 10
Public Const BrightCyan = 11
Public Const BrightRed = 12
Public Const Pink = 13
Public Const Yellow = 14
Public Const White = 15
Public Const SayColor = Grey
Public Const GlobalColor = Green
Public Const BroadcastColor = White
Public Const TellColor = White
Public Const EmoteColor = White
Public Const AdminColor = BrightCyan
Public Const HelpColor = White
Public Const WhoColor = Grey
Public Const JoinLeftColor = Grey
Public Const NpcColor = White
Public Const AlertColor = White
Public Const NewMapColor = Grey
Public TexthDC As Long
Public GameFont As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public CurrentSong As String
Public Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    reserved As Long
End Type
Public Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Public Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors.
End Type
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Public Const RASTERCAPS As Long = 38
Public Const RC_PALETTE As Long = &H100
Public Const SIZEPALETTE As Long = 104
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function GetSystemPaletteEntries Lib "gdi32.dll" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Public Declare Function CreatePalette Lib "gdi32.dll" (lpLogPalette As LOGPALETTE) As Long

Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function GetForegroundWindow Lib "USER32.DLL" () As Long
Public Declare Function SelectPalette Lib "gdi32.dll" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Public Declare Function RealizePalette Lib "gdi32.dll" (ByVal hDC As Long) As Long
Public Declare Function GetWindowDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function GetDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "USER32.DLL" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "USER32.DLL" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Public Declare Function GetDesktopWindow Lib "USER32.DLL" () As Long
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal filename$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal filename$)
Public Const MAX_PATH = 260
Private Const ERROR_NO_MORE_FILES = 18
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public SOffsetX As Integer
Public SOffsetY As Integer
Public Type WMcolors
  bgClr As Long
  frClr As Long
  fntProp As Long
End Type
Public ClrData(19) As WMcolors
Public AFileName As String
Public QuestIndex As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public CurrentQuestNum As Long
Public CurrentQuestNpcNum As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_ALPHA = &H2&
Public Const VK_UP = &H26
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11
Public mclsStyle As clsWindowed
Public Const MENU_STATE_NEWACCOUNT = 0
Public Const MENU_STATE_LOGIN = 1
Public Const MENU_STATE_GETCHARS = 2
Public Const MENU_STATE_NEWCHAR = 3
Public Const MENU_STATE_ADDCHAR = 4
Public Const MENU_STATE_DELCHAR = 5
Public Const MENU_STATE_USECHAR = 6
Public Const MENU_STATE_INIT = 7

' Speed moving vars
Public Const WALK_SPEED = 4
Public Const GM_WALK_SPEED = 4

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

' Alignment
Public AlignmentBarTime As Long

Public CorpseIndex As Integer

Public FishNumber As Integer
Public ToolNumber As Integer
Public FishName As String
Public ToolName As String

'Used for smithing
Public SmithNumber As Integer
Public SToolNumber As Integer
Public SmithName As String
Public SToolName As String

'Used for Mining
Public OreNumber As Integer
Public OreName As String

'Used for Mining
Public LogNumber As Integer
Public LogName As String

'Used for Foraging
Public SeedNumber As Integer
Public SeedName As String

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public EditorSet As Byte

Public EditorSpellX As Long
Public EditorSpellY As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map furniture editor
Public FurnitureNum As Long

Public RoofId As String

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key opene ditor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long
Public KeyOpenEditorMsg As String

' Map for local use


' Used for index based editors
Public InSpeechEditor As Boolean
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InElementEditor As Boolean
Public InEmoticonEditor As Boolean
Public InArrowEditor As Boolean
Public InQuestEditor As Boolean
Public InSpawnEditor As Boolean
Public EditorIndex As Long

' Used to know what npc we are choosing the spawn for
Public SpawnLocator As Long

' Game fps
Public GameFPS As Long

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

Public MiniMap As Boolean

Public EditorShopNum As Long

Public EditorItemNum1 As Byte
Public EditorItemNum2 As Byte
Public EditorItemNum3 As Byte

Public Arena1 As Byte
Public Arena2 As Byte
Public Arena3 As Byte

Public ii As Long, iii As Long
Public sx As Long

Public MouseDownX As Long
Public MouseDownY As Long

Public SpritePic As Long
Public SpriteItem As Long
Public SpritePrice As Long

Public HouseItem As Long
Public HousePrice As Long

Public SelectorWidth As Long
Public SelectorHeight As Long

Public SoundFileName As String

Public ScreenMode As Byte

Public SignLine1 As String
Public SignLine2 As String
Public SignLine3 As String

Public ClassChange As Long
Public ClassChangeReq As Long

Public NoticeTitle As String
Public NoticeText As String
Public NoticeSound As String

Public ScriptNum As Long

Public Connucted As Boolean

Public SpeechEditorCurrentNumber As Long
Public SpeechConvo1 As Long
Public SpeechConvo2 As Long
Public SpeechConvo3 As Long

Public ShopNum As Long

Public GoDebug As Long

Public MouseX As Long
Public MouseY As Long
Public XToGo As Long
Public YToGo As Long
Public GAME_NAME As String
Public WEBSITE As String
Public PAPERDOLL As Integer
Public SPRITESIZE As Integer
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_EMOTICONS As Long
Public MAX_SPELL_ANIM As Long
Public MAX_BLT_LINE As Long
Public MAX_SPEECH As Long
Public MAX_ELEMENTS As Long
Public Const MAX_ARROWS = 100
Public Const MAX_PLAYER_ARROWS = 100
Public Const MAX_INV = 24
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 66
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10
Public Const MAX_SPEECH_OPTIONS = 20
Public Const MAX_FRIENDS = 20
Public Const MAX_BANK = 50
Public Const MAX_QUESTS = 500
Public Const NO = 0
Public Const YES = 1
Public Const CLIENT_MAJOR = 1
Public Const CLIENT_MINOR = 1
Public Const CLIENT_REVISION = 1
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3
Public Const SEC_CODE = "89h89hr98hewf9wfnd3nf98b9s8enfs09fn390jnf83n"
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1
Public MAX_MAPX As Variant
Public MAX_MAPY As Variant
Public Const SCREEN_X = 24
Public Const SCREEN_Y = 18
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_NO_PENALTY = 2
Public Const MAP_MORAL_HOUSE = 3
Public Const PIC_X = 32
Public Const PIC_Y = 32
Public SIZE_X As Integer
Public SIZE_Y As Integer
Public Const TILE_TYPE_WALKABLE = 0
Public Const TILE_TYPE_BLOCKED = 1
Public Const TILE_TYPE_WARP = 2
Public Const TILE_TYPE_ITEM = 3
Public Const TILE_TYPE_NPCAVOID = 4
Public Const TILE_TYPE_KEY = 5
Public Const TILE_TYPE_KEYOPEN = 6
Public Const TILE_TYPE_HEAL = 7
Public Const TILE_TYPE_KILL = 8
Public Const TILE_TYPE_SHOP = 9
Public Const TILE_TYPE_CBLOCK = 10
Public Const TILE_TYPE_ARENA = 11
Public Const TILE_TYPE_SOUND = 12
Public Const TILE_TYPE_SPRITE_CHANGE = 13
Public Const TILE_TYPE_SIGN = 14
Public Const TILE_TYPE_DOOR = 15
Public Const TILE_TYPE_NOTICE = 16
Public Const TILE_TYPE_CHEST = 17
Public Const TILE_TYPE_CLASS_CHANGE = 18
Public Const TILE_TYPE_SCRIPTED = 19
Public Const TILE_TYPE_NONE = 20
Public Const TILE_TYPE_BANK = 23
Public Const TILE_TYPE_HOUSE_BUY = 24
Public Const TILE_TYPE_HOUSE = 25
Public Const TILE_TYPE_FURNITURE = 26
Public Const TILE_TYPE_ROOF = 27
Public Const TILE_TYPE_ROOFBLOCK = 28
Public Const TILE_TYPE_SPAWNGATE = 29
Public Const TILE_TYPE_FISH = 30
Public Const TILE_TYPE_MINE = 31
Public Const TILE_TYPE_LJACKING = 32
Public Const ITEM_TYPE_NONE = 0
Public Const ITEM_TYPE_WEAPON = 1
Public Const ITEM_TYPE_ARMOR = 2
Public Const ITEM_TYPE_HELMET = 3
Public Const ITEM_TYPE_SHIELD = 4
Public Const ITEM_TYPE_POTIONADDHP = 5
Public Const ITEM_TYPE_POTIONADDMP = 6
Public Const ITEM_TYPE_POTIONADDSP = 7
Public Const ITEM_TYPE_POTIONSUBHP = 8
Public Const ITEM_TYPE_POTIONSUBMP = 9
Public Const ITEM_TYPE_POTIONSUBSP = 10
Public Const ITEM_TYPE_KEY = 11
Public Const ITEM_TYPE_CURRENCY = 12
Public Const ITEM_TYPE_SPELL = 13
Public Const ITEM_TYPE_PET = 14
Public Const ITEM_TYPE_FURNITURE = 15
Public Const ITEM_TYPE_SCRIPTED = 16
Public Const ITEM_TYPE_LEGS = 17
Public Const ITEM_TYPE_BOOTS = 18
Public Const ITEM_TYPE_GLOVES = 19
Public Const ITEM_TYPE_RING1 = 20
Public Const ITEM_TYPE_RING2 = 21
Public Const ITEM_TYPE_AMULET = 22
Public Const ITEM_TYPE_GUILDDEED = 23
Public Const ITEM_TYPE_HOUSEKEY = 24
Public Const ITEM_TYPE_FOOD = 25
Public Const ITEM_TYPE_ARROWS = 26
Public Const DIR_UP = 0
Public Const DIR_DOWN = 1
Public Const DIR_LEFT = 2
Public Const DIR_RIGHT = 3
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2
Public Const WEATHER_NONE = 0
Public Const WEATHER_RAINING = 1
Public Const WEATHER_SNOWING = 2
Public Const WEATHER_THUNDER = 3
Public Const TIME_DAY = 0
Public Const TIME_NIGHT = 1
Public Const ADMIN_MONITER = 1
Public Const ADMIN_MAPPER = 2
Public Const ADMIN_DEVELOPER = 3
Public Const ADMIN_CREATOR = 4
Public Const NPC_BEHAVIOR_ATTACKONSIGHT = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED = 1
Public Const NPC_BEHAVIOR_FRIENDLY = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER = 3
Public Const NPC_BEHAVIOR_GUARD = 4
Public Const NPC_BEHAVIOR_SCRIPTED = 5
Public Const NPC_BEHAVIOR_QUEST = 6
Public Const NPC_BEHAVIOR_BANKER = 7
Public Const NPC_BEHAVIOR_SPELLCASTER = 8
Public Const DISPLAY_BUBBLE_TIME As Long = 2000 ' In milliseconds.
Public DISPLAY_BUBBLE_WIDTH As Byte
Public Const MAX_BUBBLE_WIDTH As Byte = 6 ' In tiles. Includes corners.
Public Const MAX_LINE_LENGTH As Byte = 23 ' In characters.
Public Const MAX_LINES As Byte = 3
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_PET = 6
Public Const SPELL_TYPE_SCRIPTED = 7
Public Const TARGET_TYPE_PLAYER = 0
Public Const TARGET_TYPE_NPC = 1
Public Const TARGET_TYPE_LOCATION = 2
Public Const TARGET_TYPE_PET = 3
Public Const EMOTICON_TYPE_IMAGE = 0
Public Const EMOTICON_TYPE_SOUND = 1
Public Const EMOTICON_TYPE_BOTH = 2
Public Const GFX_PASSWORD = "test"

Type PartyRec
    MemberIndex(1 To MAX_PARTY_MEMBERS) As Long
    MemberNames(1 To MAX_PARTY_MEMBERS) As String
    MemberSprite(1 To MAX_PARTY_MEMBERS) As Long
    Level(1 To MAX_PARTY_MEMBERS) As Long
    Leader As String
End Type

Type QuestRec
Name As String
LevelIsReq As Byte
ClassIsReq As Byte
StartOn As Byte
LevelReq As Integer
ClassReq As Integer

StartItem As Long
Startval As Long
ItemReq As Long
ItemVal As Long
RewardNum As Long
RewardVal As Long
Start As String
End As String
During As String
NotHasItem As String
Before As String
After As String
QuestExpReward As Long
End Type

Type ElementRec
    Name As String * NAME_LENGTH
    Strong As Long
    Weak As Long
End Type

Type BankRec
    Num As Long
    Value As Long
    Dur As Long
End Type

Type ChatBubble
    Text As String
    Created As Long
End Type

Type PlayerInvRec
    Num As Long
    Value As Long
    Dur As Long
End Type

Type SpellAnimRec
    CastedSpell As Byte
    
    SpellTime As Long
    SpellVar As Long
    SpellDone As Long
    
    Target As Long
    TargetType As Long
End Type

Type PlayerArrowRec
    Arrow As Byte
    ArrowNum As Long
    ArrowAnim As Long
    ArrowTime As Long
    ArrowVarX As Long
    ArrowVarY As Long
    ArrowX As Long
    ArrowY As Long
    ArrowPosition As Byte
End Type

Type PetRec
    Sprite As Long
   
    Alive As Byte
   
    HP As Long
    MaxHP As Long
   
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
   
    Moving As Byte
    XOffset As Long
    YOffset As Long
   
    AttackTimer As Long
    Attacking As Byte
   
    LastAttack As Long
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Class As Long
    Sprite As Long
    Level As Long
    EXP As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    FP As Long
    
    ' Stats
    STR As Long
    DEF As Long
    speed As Long
    MAGI As Long
    POINTS As Long
    
    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    LegsSlot As Long
    BootsSlot As Long
    GlovesSlot As Long
    Ring1Slot As Long
    Ring2Slot As Long
    AmuletSlot As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    Bank(1 To MAX_BANK) As BankRec
       
    ' Position
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    
    ' Pet!
    Pet As PetRec
    
    ' Client use only
    MaxHP As Long
    MaxMP As Long
    MaxSP As Long
    MaxFP As Long
    XOffset As Integer
    YOffset As Integer
    MovingH As Integer
    MovingV As Integer
    Attacking As Byte
    AttackTimer As Long
    LastAttack As Long
    MapGetTimer As Long
    CastedSpell As Byte
    
    SpellNum As Long
    SpellAnim() As SpellAnimRec

    EmoticonNum As Long
    EmoticonSound As String
    EmoticonType As Long
    EmoticonTime As Long
    EmoticonVar As Long
    EmoticonPlayed As Boolean
    
    LevelUp As Long
    LevelUpT As Long
    
    ArmorNum As Long
    WeaponNum As Long
    ShieldNum As Long
    HelmetNum As Long
    LegsNum As Long
    BootsNum As Long
    GlovesNum As Long
    Ring1Num As Long
    Ring2Num As Long
    AmuletNum As Long

    Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec
    Hands As Long
    
    Alignment As Long
    SenseAlignment As Long
    SenseAlignmentTime As Long
    
    CorpseMap As Integer
    CorpseX As Byte
    CorpseY As Byte
    CorpseLoot(1 To 4) As PlayerInvRec
    FishEXP As Long
    MineEXP As Long
    LJackingEXP As Long
    LargeBladesEXP As Long
    SmallBladesEXP As Long
    BluntWeaponsEXP As Long
    PolesEXP As Long
    AxesEXP As Long
    ThrownEXP As Long
    XbowsEXP As Long
    BowsEXP As Long
    FishLevel As Long
    MineLevel As Long
    LJackingLevel As Long
    LargeBladesLevel As Long
    SmallBladesLevel As Long
    BluntWeaponsLevel As Long
    PolesLevel As Long
    AxesLevel As Long
    ThrownLevel As Long
    XbowsLevel As Long
    BowsLevel As Long
    Race As Long
    SpawnGateMap As Long
    SpawnGateX As Long
    SpawnGateY As Long
    ArrowsAmount As Long
    QuestFlags(1 To MAX_QUESTS) As Long
    InParty As Boolean
    Party As PartyRec
    Poisoned As Byte
    Diseased As Byte
End Type
    
Type TileRec
    Ground As Long
    Mask As Long
    Anim As Long
    Mask2 As Long
    M2Anim As Long
    Fringe As Long
    FAnim As Long
    Fringe2 As Long
    F2Anim As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    String2 As String
    String3 As String
    Light As Long
    GroundSet As Long
    MaskSet As Long
    AnimSet As Long
    Mask2Set As Long
    M2AnimSet As Long
    FringeSet As Long
    FAnimSet As Long
    Fringe2Set As Long
    F2AnimSet As Long
End Type

Type LocRec
    Used As Byte
    X As Long
    Y As Long
End Type

Type MapRec
    Name As String * 40
    Revision As Long
    Moral As Byte
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    Music As String
    BootMap As Long
    BootX As Byte
    BootY As Byte
    Shop As Long
    Indoors As Byte
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    NpcSpawn(1 To MAX_MAP_NPCS) As LocRec
    Owner As String
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    MaleSprite As Long
    FemaleSprite As Long
    
    Locked As Long
    
    STR As Long
    DEF As Long
    speed As Long
    MAGI As Long
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    desc As String * 150
    
    pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    StrReq As Long
    DefReq As Long
    SpeedReq As Long
    MagicReq As Long
    ClassReq As Long
    AccessReq As Byte
    
    AddHP As Long
    AddMP As Long
    AddSP As Long
    AddStr As Long
    AddDef As Long
    AddMagi As Long
    AddSpeed As Long
    AddEXP As Long
    AttackSpeed As Long
    Price As Long
    Stackable As Long
    Bound As Long
    LevelReq As Long
    Element As Long
    StamRemove As Long
    Rarity As String * 11
    BowsReq As Long
    LargeBladesReq As Long
    SmallBladesReq As Long
    BluntWeaponsReq As Long
    PoleArmsReq As Long
    AxesReq As Long
    ThrownReq As Long
    XbowsReq As Long
    LBA As Long
    SBA As Long
    BWA As Long
    PAA As Long
    AA As Long
    TWA As Long
    XBA As Long
    BA As Long
    Poison As Long
    Disease As Long
    AilmentDamage As Long
    AilmentInterval As Long
    AilmentMS As Long
End Type

Type MapItemRec
    Num As Long
    Value As Long
    Dur As Long
    
    X As Byte
    Y As Byte
End Type

Type NPCEditorRec
    itemnum As Long
    itemvalue As Long
    Chance As Long
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    
    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
        
    STR  As Long
    DEF As Long
    speed As Long
    MAGI As Long
    Big As Long
    MaxHP As Long
    EXP As Long
    SpawnTime As Long
    
    Speech As Long
    
    Script As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
    Element As Long
    Poison As Long
    AP As Long
    Disease As Long
    Quest As Integer
    NpcDIR As Byte
    AilmentDamage As Long
    AilmentInterval As Long
    AilmentMS As Long
    Spell As Long
End Type

Type MapNpcRec
    Num As Long
    
    Target As Long
    
    HP As Long
    MaxHP As Long
    MP As Long
    SP As Long
    
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    Big As Byte

    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    LastAttack As Long
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type TradeItemsRec
    Value(1 To MAX_TRADES) As TradeItemRec
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
    FixesItems As Byte
    TradeItem(1 To 6) As TradeItemsRec
End Type

Type SpellRec
    Name As String * NAME_LENGTH
    ClassReq As Long
    LevelReq As Long
    Sound As Long
    MPCost As Long
    Type As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Range As Byte
    
    SpellAnim As Long
    SpellTime As Long
    SpellDone As Long
    
    AE As Long
    pic As Long
    Element As Long
End Type

Type TempTileRec
    DoorOpen As Byte
    
    Ground As Long
    Mask As Long
    Anim As Long
    Mask2 As Long
    M2Anim As Long
    Fringe As Long
    FAnim As Long
    Fringe2 As Long
    F2Anim As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    String2 As String
    String3 As String
    Light As Long
    GroundSet As Byte
    MaskSet As Byte
    AnimSet As Byte
    Mask2Set As Byte
    M2AnimSet As Byte
    FringeSet As Byte
    FAnimSet As Byte
    Fringe2Set As Byte
    F2AnimSet As Byte
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
    InvVal As Long
End Type
Public Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public Trading2(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

Type EmoRec
    pic As Long
    Sound As String
    Command As String
    Type As Byte
End Type

Type OptionRec
    Text As String
    GoTo As Long
    Exit As Byte
End Type

Type InvSpeechRec
    Exit As Byte
    Text As String
    SaidBy As Byte
    Respond As Byte
    Script As Long
    Responces(1 To 3) As OptionRec
End Type

Type SpeechRec
    Name As String
    Num(0 To MAX_SPEECH_OPTIONS) As InvSpeechRec
End Type

Type DropRainRec
    X As Long
    Y As Long
    Randomized As Boolean
    speed As Byte
End Type

' Bubble thing
Public Bubble() As ChatBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1
Public NEXT_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

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
Public Speech() As SpeechRec
Public Quest(0 To MAX_QUESTS) As QuestRec
Public SaveMapItem() As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

Public MAX_RAINDROPS As Long
Public BLT_RAIN_DROPS As Long
Public DropRain() As DropRainRec

Public BLT_SNOW_DROPS As Long
Public DropSnow() As DropRainRec

Type ItemTradeRec
    ItemGetNum As Long
    ItemGiveNum As Long
    ItemGetVal As Long
    ItemGiveVal As Long
End Type
Type TradeRec
    Items(1 To MAX_TRADES) As ItemTradeRec
    Selected As Long
    SelectedItem As Long
End Type
Public Trade(1 To 6) As TradeRec

Type ArrowRec
    Name As String
    pic As Long
    Range As Byte
End Type
Public Arrows(1 To MAX_ARROWS) As ArrowRec

Type BattleMsgRec
    Msg As String
    Index As Byte
    Color As Byte
    Time As Long
    Done As Byte
    Y As Long
End Type
Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Type ItemDurRec
    Item As Long
    Dur As Long
    Done As Byte
End Type
Public ItemDur(1 To 4) As ItemDurRec

Public TempNpcSpawn(1 To MAX_MAP_NPCS) As LocRec
Public picDrop(1 To MAX_NPC_DROPS) As Long
Public Inventory As Long
Public SpellIndex As Long
Public SpellMemorized As Long
Public charselsprite(MAX_CHARS) As Double


Sub ClearTempTile()
Dim X As Long, Y As Long

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            TempTile(X, Y).DoorOpen = NO
            
            TempTile(X, Y).Ground = 0
            TempTile(X, Y).Mask = 0
            TempTile(X, Y).Anim = 0
            TempTile(X, Y).Mask2 = 0
            TempTile(X, Y).M2Anim = 0
            TempTile(X, Y).Fringe = 0
            TempTile(X, Y).FAnim = 0
            TempTile(X, Y).Fringe2 = 0
            TempTile(X, Y).F2Anim = 0
            TempTile(X, Y).Type = TILE_TYPE_NONE
            TempTile(X, Y).Data1 = 0
            TempTile(X, Y).Data2 = 0
            TempTile(X, Y).Data3 = 0
            TempTile(X, Y).String1 = ""
            TempTile(X, Y).String2 = ""
            TempTile(X, Y).String3 = ""
            TempTile(X, Y).Light = 0
            TempTile(X, Y).GroundSet = 0
            TempTile(X, Y).MaskSet = 0
            TempTile(X, Y).AnimSet = 0
            TempTile(X, Y).Mask2Set = 0
            TempTile(X, Y).M2AnimSet = 0
            TempTile(X, Y).FringeSet = 0
            TempTile(X, Y).FAnimSet = 0
            TempTile(X, Y).Fringe2Set = 0
            TempTile(X, Y).F2AnimSet = 0
        Next X
    Next Y
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim I As Long
Dim n As Long

    Player(Index).Name = ""
    Player(Index).Guild = ""
    Player(Index).Guildaccess = 0
    Player(Index).Class = 1
    Player(Index).Level = 0
    Player(Index).Sprite = 0
    Player(Index).EXP = 0
    Player(Index).Access = 0
    Player(Index).PK = NO
        
    Player(Index).HP = 0
    Player(Index).MP = 0
    Player(Index).SP = 0
    Player(Index).FP = 0
        
    Player(Index).STR = 0
    Player(Index).DEF = 0
    Player(Index).speed = 0
    Player(Index).MAGI = 0
        
    For n = 1 To MAX_INV
        Player(Index).Inv(n).Num = 0
        Player(Index).Inv(n).Value = 0
        Player(Index).Inv(n).Dur = 0
    Next n
        
    For n = 1 To MAX_BANK
        Player(Index).Bank(n).Num = 0
        Player(Index).Bank(n).Value = 0
        Player(Index).Bank(n).Dur = 0
    Next n
        
    Player(Index).ArmorSlot = 0
    Player(Index).WeaponSlot = 0
    Player(Index).HelmetSlot = 0
    Player(Index).ShieldSlot = 0
    Player(Index).Hands = 0
    Player(Index).LegsSlot = 0
    Player(Index).BootsSlot = 0
    Player(Index).GlovesSlot = 0
    Player(Index).Ring1Slot = 0
    Player(Index).Ring2Slot = 0
    Player(Index).AmuletSlot = 0
        
    Player(Index).Map = 0
    Player(Index).X = 0
    Player(Index).Y = 0
    Player(Index).Dir = 0
    
    ' Client use only
    Player(Index).MaxHP = 0
    Player(Index).MaxMP = 0
    Player(Index).MaxSP = 0
    Player(Index).XOffset = 0
    Player(Index).YOffset = 0
    Player(Index).MovingH = 0
    Player(Index).MovingV = 0
    Player(Index).Attacking = 0
    Player(Index).AttackTimer = 0
    Player(Index).MapGetTimer = 0
    Player(Index).CastedSpell = NO
    Player(Index).EmoticonNum = -1
    Player(Index).EmoticonSound = ""
    Player(Index).EmoticonType = 0
    Player(Index).EmoticonTime = 0
    Player(Index).EmoticonVar = 0
    Player(Index).EmoticonPlayed = True
    
    For I = 1 To MAX_SPELL_ANIM
        Player(Index).SpellAnim(I).CastedSpell = NO
        Player(Index).SpellAnim(I).SpellTime = 0
        Player(Index).SpellAnim(I).SpellVar = 0
        Player(Index).SpellAnim(I).SpellDone = 0
        
        Player(Index).SpellAnim(I).Target = 0
        Player(Index).SpellAnim(I).TargetType = TARGET_TYPE_PLAYER
    Next I
    
    Player(Index).SpellNum = 0
    
    Player(Index).ArmorNum = 0
    Player(Index).WeaponNum = 0
    Player(Index).ShieldNum = 0
    Player(Index).HelmetNum = 0
    Player(Index).LegsNum = 0
    Player(Index).BootsNum = 0
    Player(Index).GlovesNum = 0
    Player(Index).Ring1Num = 0
    Player(Index).Ring2Num = 0
    Player(Index).AmuletNum = 0
    
    For I = 1 To MAX_BLT_LINE
        BattlePMsg(I).Index = 1
        BattlePMsg(I).Time = I
        BattleMMsg(I).Index = 1
        BattleMMsg(I).Time = I
    Next I
    
    Player(Index).CorpseMap = 0
    Player(Index).CorpseX = 0
    Player(Index).CorpseY = 0
    For I = 1 To 4
    Player(Index).CorpseLoot(I).Dur = 0
    Player(Index).CorpseLoot(I).Num = 0
    Player(Index).CorpseLoot(I).Value = 0
    Next I
    
    If CorpseIndex = Index Then
    CorpseIndex = 0
    End If
    
    Player(Index).LargeBladesLevel = 0
    Player(Index).SmallBladesLevel = 0
    Player(Index).BluntWeaponsLevel = 0
    Player(Index).PolesLevel = 0
    Player(Index).AxesLevel = 0
    Player(Index).ThrownLevel = 0
    Player(Index).XbowsLevel = 0
    Player(Index).BowsLevel = 0
    Player(Index).LargeBladesEXP = 0
    Player(Index).SmallBladesEXP = 0
    Player(Index).BluntWeaponsEXP = 0
    Player(Index).PolesEXP = 0
    Player(Index).AxesEXP = 0
    Player(Index).ThrownEXP = 0
    Player(Index).XbowsEXP = 0
    Player(Index).BowsEXP = 0
    Player(Index).Race = 1
    Player(Index).SpawnGateMap = 0
    Player(Index).SpawnGateY = 0
    Player(Index).SpawnGateX = 0
    Player(Index).ArrowsAmount = 0
    Player(Index).FishLevel = 0
    Player(Index).MineLevel = 0
    Player(Index).LJackingLevel = 0
    Player(Index).FishEXP = 0
    Player(Index).MineEXP = 0
    Player(Index).LJackingEXP = 0
    
    Inventory = 1
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index).Name = ""
    Item(Index).desc = ""
    
    Item(Index).Type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
    Item(Index).StrReq = 0
    Item(Index).DefReq = 0
    Item(Index).SpeedReq = 0
    Item(Index).MagicReq = 0
    Item(Index).ClassReq = 0
    Item(Index).AccessReq = 0
    
    Item(Index).AddHP = 0
    Item(Index).AddMP = 0
    Item(Index).AddSP = 0
    Item(Index).AddStr = 0
    Item(Index).AddDef = 0
    Item(Index).AddMagi = 0
    Item(Index).AddSpeed = 0
    Item(Index).AddEXP = 0
    Item(Index).AttackSpeed = 1000
    Item(Index).Stackable = 0
    Item(Index).Bound = 0
    Item(Index).LevelReq = 0
    Item(Index).Element = 0
    Item(Index).StamRemove = 0
    Item(Index).Rarity = ""
End Sub

Sub ClearItems()
Dim I As Long

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next I
End Sub

Sub ClearMapItem(ByVal Index As Long)
    MapItem(Index).Num = 0
    MapItem(Index).Value = 0
    MapItem(Index).Dur = 0
    MapItem(Index).X = 0
    MapItem(Index).Y = 0
End Sub

Sub ClearMaps()
Dim I As Long

For I = 1 To MAX_MAPS
    Call ClearMap(I)
Next I
End Sub

Sub ClearMap(ByVal MapNum As Long)
Dim I, X, Y As Long

    I = MapNum
    Map(I).Name = ""
    Map(I).Revision = 0
    Map(I).Moral = 0
    Map(I).Up = 0
    Map(I).Down = 0
    Map(I).Left = 0
    Map(I).Right = 0
    Map(I).Indoors = 0
        
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map(I).Tile(X, Y).Ground = 0
            Map(I).Tile(X, Y).Mask = 0
            Map(I).Tile(X, Y).Anim = 0
            Map(I).Tile(X, Y).Mask2 = 0
            Map(I).Tile(X, Y).M2Anim = 0
            Map(I).Tile(X, Y).Fringe = 0
            Map(I).Tile(X, Y).FAnim = 0
            Map(I).Tile(X, Y).Fringe2 = 0
            Map(I).Tile(X, Y).F2Anim = 0
            Map(I).Tile(X, Y).Type = 0
            Map(I).Tile(X, Y).Data1 = 0
            Map(I).Tile(X, Y).Data2 = 0
            Map(I).Tile(X, Y).Data3 = 0
            Map(I).Tile(X, Y).String1 = ""
            Map(I).Tile(X, Y).String2 = ""
            Map(I).Tile(X, Y).String3 = ""
            Map(I).Tile(X, Y).Light = 0
            Map(I).Tile(X, Y).GroundSet = -1
            Map(I).Tile(X, Y).MaskSet = -1
            Map(I).Tile(X, Y).AnimSet = -1
            Map(I).Tile(X, Y).Mask2Set = -1
            Map(I).Tile(X, Y).M2AnimSet = -1
            Map(I).Tile(X, Y).FringeSet = -1
            Map(I).Tile(X, Y).FAnimSet = -1
            Map(I).Tile(X, Y).Fringe2Set = -1
            Map(I).Tile(X, Y).F2AnimSet = -1
        Next X
    Next Y
End Sub

Sub ClearMapItems()
Dim X As Long

    For X = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(X)
    Next X
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    MapNpc(Index).Num = 0
    MapNpc(Index).Target = 0
    MapNpc(Index).HP = 0
    MapNpc(Index).MP = 0
    MapNpc(Index).SP = 0
    MapNpc(Index).Map = 0
    MapNpc(Index).X = 0
    MapNpc(Index).Y = 0
    MapNpc(Index).Dir = 0
    
    ' Client use only
    MapNpc(Index).XOffset = 0
    MapNpc(Index).YOffset = 0
    MapNpc(Index).Moving = 0
    MapNpc(Index).Attacking = 0
    MapNpc(Index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
Dim I As Long

    For I = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(I)
    Next I
End Sub

Sub ClearSpeech(ByVal Index As Long)
Dim I As Long
Dim O As Long

    Speech(Index).Name = ""

    For O = 0 To MAX_SPEECH_OPTIONS
        Speech(Index).Num(O).Exit = 0
        Speech(Index).Num(O).Respond = 0
        Speech(Index).Num(O).SaidBy = 0
        Speech(Index).Num(O).Text = "Write what you want to be said here."
        Speech(Index).Num(O).Script = 0
    
        For I = 1 To 3
            Speech(Index).Num(O).Responces(I).Exit = 0
            Speech(Index).Num(O).Responces(I).GoTo = 0
            Speech(Index).Num(O).Responces(I).Text = "Write a responce here."
        Next I
    Next O
End Sub

Sub ClearSpeeches()
Dim I As Long

    For I = 1 To MAX_SPEECH
        Call ClearSpeech(I)
    Next I
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim(Player(Index).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Guildaccess As Long)
    Player(Index).Guildaccess = Guildaccess
End Sub


Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    Player(Index).EXP = EXP
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).HP = HP
    
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).HP = GetPlayerMaxHP(Index)
    End If
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).MP = GetPlayerMaxMP(Index)
    End If
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).SP = GetPlayerMaxSP(Index)
    End If
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    GetPlayerMaxHP = Player(Index).MaxHP
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    GetPlayerMaxMP = Player(Index).MaxMP
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    GetPlayerMaxSP = Player(Index).MaxSP
End Function

Function GetPlayerSTR(ByVal Index As Long) As Long
    GetPlayerSTR = Player(Index).STR
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).DEF
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal speed As Long)
    Player(Index).speed = speed
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).MAGI
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal MAGI As Long)
    Player(Index).MAGI = MAGI
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
If Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal itemnum As Long)
    Player(Index).Inv(InvSlot).Num = itemnum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal itemvalue As Long)
    Player(Index).Inv(InvSlot).Value = itemvalue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).ShieldSlot = InvNum
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    If BankSlot > MAX_BANK Then Exit Function
    GetPlayerBankItemNum = Player(Index).Bank(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal itemnum As Long)
    Player(Index).Bank(BankSlot).Num = itemnum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal itemvalue As Long)
    Player(Index).Bank(BankSlot).Value = itemvalue
End Sub

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemDur = Player(Index).Bank(BankSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
    Player(Index).Bank(BankSlot).Dur = ItemDur
End Sub

Sub SetPlayerHands(ByVal Index As Long, Item As Long)
    Player(Index).Hands = Item
    Call SendData("SETHANDS" & SEP_CHAR & Player(Index).Hands & SEP_CHAR & END_CHAR)
End Sub

Sub SetPlayerAlignment(ByVal Index As Long, Num As Long)
    Player(Index).Alignment = Num
End Sub

Function GetPlayerAlignment(ByVal Index As Long) As Long
    GetPlayerAlignment = Player(Index).Alignment
End Function

Function GetPlayerLegsSlot(ByVal Index As Long) As Long
    GetPlayerLegsSlot = Player(Index).LegsSlot
End Function

Sub SetPlayerLegsSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).LegsSlot = InvNum
End Sub

Function GetPlayerBootsSlot(ByVal Index As Long) As Long
    GetPlayerBootsSlot = Player(Index).BootsSlot
End Function

Sub SetPlayerBootsSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).BootsSlot = InvNum
End Sub

Function GetPlayerGlovesSlot(ByVal Index As Long) As Long
    GetPlayerGlovesSlot = Player(Index).GlovesSlot
End Function

Sub SetPlayerGlovesSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).GlovesSlot = InvNum
End Sub

Function GetPlayerRing1Slot(ByVal Index As Long) As Long
    GetPlayerRing1Slot = Player(Index).Ring1Slot
End Function

Sub SetPlayerRing1Slot(ByVal Index As Long, InvNum As Long)
    Player(Index).Ring1Slot = InvNum
End Sub

Function GetPlayerRing2Slot(ByVal Index As Long) As Long
    GetPlayerRing2Slot = Player(Index).Ring2Slot
End Function

Sub SetPlayerRing2Slot(ByVal Index As Long, InvNum As Long)
    Player(Index).Ring2Slot = InvNum
End Sub

Function GetPlayerAmuletSlot(ByVal Index As Long) As Long
    GetPlayerAmuletSlot = Player(Index).AmuletSlot
End Function

Sub SetPlayerAmuletSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).AmuletSlot = InvNum
End Sub

Function GetPlayerRace(ByVal Index As Long) As Long
    GetPlayerRace = Player(Index).Race
End Function

Sub SetPlayerRace(ByVal Index As Long, ByVal RaceNum As Long)
    Player(Index).Race = RaceNum
End Sub

Function GetPlayerLargeBladesLevel(ByVal Index As Long) As Long
    GetPlayerLargeBladesLevel = Player(Index).LargeBladesLevel
End Function

Sub SetPlayerLargeBladesLevel(ByVal Index As Long, ByVal LargeBladesLevel As Long)
    Player(Index).LargeBladesLevel = LargeBladesLevel
End Sub

Function GetPlayerLargeBladesExp(ByVal Index As Long) As Long
    GetPlayerLargeBladesExp = Player(Index).LargeBladesEXP
End Function

Sub SetPlayerLargeBladesExp(ByVal Index As Long, ByVal LargeBladesEXP As Long)
    Player(Index).LargeBladesEXP = LargeBladesEXP
End Sub

Function GetPlayerSmallBladesLevel(ByVal Index As Long) As Long
    GetPlayerSmallBladesLevel = Player(Index).SmallBladesLevel
End Function

Sub SetPlayerSmallBladesLevel(ByVal Index As Long, ByVal SmallBladesLevel As Long)
    Player(Index).SmallBladesLevel = SmallBladesLevel
End Sub

Function GetPlayerSmallBladesExp(ByVal Index As Long) As Long
    GetPlayerSmallBladesExp = Player(Index).SmallBladesEXP
End Function

Sub SetPlayerSmallBladesExp(ByVal Index As Long, ByVal SmallBladesEXP As Long)
    Player(Index).SmallBladesEXP = SmallBladesEXP
End Sub

Function GetPlayerBluntWeaponsLevel(ByVal Index As Long) As Long
    GetPlayerBluntWeaponsLevel = Player(Index).BluntWeaponsLevel
End Function

Sub SetPlayerBluntWeaponsLevel(ByVal Index As Long, ByVal BluntWeaponsLevel As Long)
    Player(Index).BluntWeaponsLevel = BluntWeaponsLevel
End Sub

Function GetPlayerBluntWeaponsExp(ByVal Index As Long) As Long
    GetPlayerBluntWeaponsExp = Player(Index).BluntWeaponsEXP
End Function

Sub SetPlayerBluntWeaponsExp(ByVal Index As Long, ByVal BluntWeaponsEXP As Long)
    Player(Index).BluntWeaponsEXP = BluntWeaponsEXP
End Sub

Function GetPlayerPolesLevel(ByVal Index As Long) As Long
    GetPlayerPolesLevel = Player(Index).PolesLevel
End Function

Sub SetPlayerPolesLevel(ByVal Index As Long, ByVal PolesLevel As Long)
    Player(Index).PolesLevel = PolesLevel
End Sub

Function GetPlayerPolesExp(ByVal Index As Long) As Long
    GetPlayerPolesExp = Player(Index).PolesEXP
End Function

Sub SetPlayerPolesExp(ByVal Index As Long, ByVal PolesEXP As Long)
    Player(Index).PolesEXP = PolesEXP
End Sub

Function GetPlayerAxesLevel(ByVal Index As Long) As Long
    GetPlayerAxesLevel = Player(Index).AxesLevel
End Function

Sub SetPlayerAxesLevel(ByVal Index As Long, ByVal AxesLevel As Long)
    Player(Index).AxesLevel = AxesLevel
End Sub

Function GetPlayerAxesExp(ByVal Index As Long) As Long
    GetPlayerAxesExp = Player(Index).AxesEXP
End Function

Sub SetPlayerAxesExp(ByVal Index As Long, ByVal AxesEXP As Long)
    Player(Index).AxesEXP = AxesEXP
End Sub

Function GetPlayerThrownLevel(ByVal Index As Long) As Long
    GetPlayerThrownLevel = Player(Index).ThrownLevel
End Function

Sub SetPlayerThrownLevel(ByVal Index As Long, ByVal ThrownLevel As Long)
    Player(Index).ThrownLevel = ThrownLevel
End Sub

Function GetPlayerThrownExp(ByVal Index As Long) As Long
    GetPlayerThrownExp = Player(Index).ThrownEXP
End Function

Sub SetPlayerThrownExp(ByVal Index As Long, ByVal ThrownEXP As Long)
    Player(Index).ThrownEXP = ThrownEXP
End Sub

Function GetPlayerXbowsLevel(ByVal Index As Long) As Long
    GetPlayerXbowsLevel = Player(Index).XbowsLevel
End Function

Sub SetPlayerXbowsLevel(ByVal Index As Long, ByVal XbowsLevel As Long)
    Player(Index).XbowsLevel = XbowsLevel
End Sub

Function GetPlayerXbowsExp(ByVal Index As Long) As Long
    GetPlayerXbowsExp = Player(Index).XbowsEXP
End Function

Sub SetPlayerXbowsExp(ByVal Index As Long, ByVal XbowsEXP As Long)
    Player(Index).XbowsEXP = XbowsEXP
End Sub

Function GetPlayerBowsLevel(ByVal Index As Long) As Long
    GetPlayerBowsLevel = Player(Index).BowsLevel
End Function

Sub SetPlayerBowsLevel(ByVal Index As Long, ByVal BowsLevel As Long)
    Player(Index).BowsLevel = BowsLevel
End Sub

Function GetPlayerBowsExp(ByVal Index As Long) As Long
    GetPlayerBowsExp = Player(Index).BowsEXP
End Function

Sub SetPlayerBowsExp(ByVal Index As Long, ByVal BowsEXP As Long)
    Player(Index).BowsEXP = BowsEXP
End Sub

Function GetPlayerSpawnGateMap(ByVal Index As Long) As Long
If Index <= 0 Then Exit Function
    GetPlayerSpawnGateMap = Player(Index).SpawnGateMap
End Function

Sub SetPlayerSpawnGateMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).SpawnGateMap = MapNum
End Sub

Function GetPlayerSpawnGateX(ByVal Index As Long) As Long
    GetPlayerSpawnGateX = Player(Index).SpawnGateX
End Function

Sub SetPlayerSpawnGateX(ByVal Index As Long, ByVal X As Long)
    Player(Index).SpawnGateX = X
End Sub

Function GetPlayerSpawnGateY(ByVal Index As Long) As Long
    GetPlayerSpawnGateY = Player(Index).SpawnGateY
End Function

Sub SetPlayerSpawnGateY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).SpawnGateY = Y
End Sub

Function GetPlayerFP(ByVal Index As Long) As Long
    GetPlayerFP = Player(Index).FP
End Function

Sub SetPlayerFP(ByVal Index As Long, ByVal FP As Long)
    Player(Index).FP = FP
    
    If GetPlayerFP(Index) > GetPlayerMaxFP(Index) Then
        Player(Index).HP = GetPlayerMaxFP(Index)
    End If
End Sub

Function GetPlayerMaxFP(ByVal Index As Long) As Long
    GetPlayerMaxFP = Player(Index).MaxFP
End Function

Sub SetPlayerArrowsAmount(ByVal Index As Long, ArrowsAmount As Long)
    Player(Index).ArrowsAmount = ArrowsAmount
End Sub

Function GetPlayerArrowsAmount(ByVal Index As Long) As Long
    GetPlayerArrowsAmount = Player(Index).ArrowsAmount
End Function

Function GetPlayerFishLevel(ByVal Index As Long) As Long
    GetPlayerFishLevel = Player(Index).FishLevel
End Function

Function GetPlayerMineLevel(ByVal Index As Long) As Long
    GetPlayerMineLevel = Player(Index).MineLevel
End Function

Sub SetPlayerFishLevel(ByVal Index As Long, ByVal FishLevel As Long)
    Player(Index).FishLevel = FishLevel
End Sub

Sub SetPlayerMineLevel(ByVal Index As Long, ByVal MineLevel As Long)
    Player(Index).MineLevel = MineLevel
End Sub

Function GetPlayerFishExp(ByVal Index As Long) As Long
    GetPlayerFishExp = Player(Index).FishEXP
End Function

Function GetPlayerMineExp(ByVal Index As Long) As Long
    GetPlayerMineExp = Player(Index).MineEXP
End Function

Sub SetPlayerFishExp(ByVal Index As Long, ByVal FishEXP As Long)
    Player(Index).FishEXP = FishEXP
End Sub

Sub SetPlayerMineExp(ByVal Index As Long, ByVal MineEXP As Long)
    Player(Index).MineEXP = MineEXP
End Sub

Function GetPlayerLJackingExp(ByVal Index As Long) As Long
    GetPlayerLJackingExp = Player(Index).LJackingEXP
End Function

Function GetPlayerLJackingLevel(ByVal Index As Long) As Long
    GetPlayerLJackingLevel = Player(Index).LJackingLevel
End Function

Sub SetPlayerLJackingExp(ByVal Index As Long, ByVal LJackingEXP As Long)
    Player(Index).LJackingEXP = LJackingEXP
End Sub

Sub SetPlayerLJackingLevel(ByVal Index As Long, ByVal LJackingLevel As Long)
    Player(Index).LJackingLevel = LJackingLevel
End Sub
