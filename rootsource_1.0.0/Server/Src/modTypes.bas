Attribute VB_Name = "modTypes"
Option Explicit

' ********************************************
' **               rootSource               **
' ********************************************

' Public data structures
Public Map() As MapRec
Public MapCache() As Cache
Public TempTile() As TempTileRec
Public PlayersOnMap() As Long
Public Player() As AccountRec
Public TempPlayer() As TempPlayerRec
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc() As MapNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec

Public Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Public Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Sex As Byte
    Class As Byte
    Sprite As Integer
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    
    Guild As String
    GuildAccess As Long
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Byte
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Byte
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
    
    ' Position
    Map As Integer
    X As Byte
    y As Byte
    Dir As Byte
End Type
Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
       
    ' Characters
    ' 0 is used to prevent an RTE9 when accessing a cleared account
    Char(0 To MAX_CHARS) As PlayerRec
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    Buffer As clsBuffer
    CharNum As Byte
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    PartyPlayer As Long
    InParty As Byte
    TargetType As Byte
    Target As Byte
    CastedSpell As Byte
    PartyStarter As Byte
    GettingMap As Byte
End Type

Public Type TileRec
    Num(0 To 8) As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type
Public Type MapRec
    Name As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    TileSet As Integer
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    Music As Byte
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Shop As Byte
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Byte
End Type

Public Type ClassRec
    Name As String * NAME_LENGTH
    
    Sprite As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Byte
End Type

Public Type ItemRec
    Name As String * NAME_LENGTH
    
    Pic As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Public Type MapItemRec
    Num As Byte
    Value As Long
    Dur As Integer
    
    X As Byte
    y As Byte
End Type

Public Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Byte
End Type

Public Type MapNpcRec
    Num As Integer
    
    Target As Integer
    
    Vital(1 To Vitals.Vital_Count - 1) As Long
        
    X As Byte
    y As Byte
    Dir As Integer
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
End Type

Public Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Public Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 255
    LeaveSay As String * 255
    FixesItems As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type
    
Public Type SpellRec
    Name As String * NAME_LENGTH
    Pic As Integer
    MPReq As Integer
    ClassReq As Byte
    LevelReq As Byte
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Public Type TempTileRec
    DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY) As Byte
    DoorTimer As Long
End Type

Public Type Cache
    Cache() As Byte
End Type
