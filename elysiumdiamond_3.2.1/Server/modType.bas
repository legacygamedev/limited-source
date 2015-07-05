Attribute VB_Name = "modType"
Option Explicit

Public Const MAX_FRIENDS As Byte = 20
Public Const MAX_CHARS As Byte = 3
Public Const MAX_INV As Byte = 24
Public Const MAX_PLAYER_TRADES As Byte = 8
Public Const MAX_PLAYER_SPELLS As Byte = 20

Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long

Public Const YES As Byte = 1
Public Const NO As Byte = 0

Public START_MAP As Long
Public START_X As Long
Public START_Y As Long

Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

Public Player() As AccountRec
Public Spell() As SpellRec

Type SpellRec
    Name As String
    ClassReq As Long
    LevelReq As Long
    MPCost As Long
    sound As Long

    Type As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Range As Byte
    SpellAnim As Long
    SpellTime As Long
    SpellDone As Long
    AE As Long
End Type

Type PlayerInvRec
    num As Long
    Value As Long
    Dur As Long
End Type

Type PlayerRec

    ' General
    Name As String
    Guild As String
    Guildaccess As Byte
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Long
    Exp As Long
    Access As Byte
    PK As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long

    ' Stats
    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long

    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long

    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    Friends(1 To MAX_FRIENDS) As String
End Type

Type PetRec
    Sprite As Long
    Alive As Byte
    Map As Long
    x As Long
    y As Long
    Dir As Byte
    Level As Long
    HP As Long
    MapToGo As Long
    XToGo As Long
    YToGo As Long
    Target As Long
    TargetType As Byte
    AttackTimer As Long
End Type

Type AccountRec

    ' Account
    Login As String
    Password As String

    ' Characters (we use 0 to prevent a crash that still needs to be figured out)
    Char(0 To MAX_CHARS) As PlayerRec

    ' None saved local vars
    Buffer As String
    IncBuffer As String
    CharNum As Byte
    InGame As Byte
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    PartyID As Long
    InParty As Byte
    Invited As Long
    TargetType As Byte
    Target As Long
    CastedSpell As Byte
    SpellVar As Long
    SpellDone As Long
    SpellNum As Long
    GettingMap As Byte
    Emoticon As Long
    InTrade As Byte
    TradePlayer As Long
    TradeOk As Byte
    TradeItemMax As Byte
    TradeItemMax2 As Byte
    'Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
    InChat As Byte
    ChatPlayer As Long
    Mute As Boolean
    Pet As PetRec
End Type
