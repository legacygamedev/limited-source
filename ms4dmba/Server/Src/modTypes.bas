Attribute VB_Name = "modTypes"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec
Public MapCache(1 To MAX_MAPS) As Cache
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public Player(1 To MAX_PLAYERS) As AccountRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
'Public MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
Public MapNpc(1 To MAX_MAPS) As MapDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec

Private Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Private Type Cache
    Data() As Byte
End Type

Private Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Sex As Byte
    Class As Byte
    Sprite As Integer
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    
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
    x As Byte
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

Private Type TileRec
    Ground As Integer
    Mask As Integer
    Anim As Integer
    Mask2 As Integer
    Fringe As Integer
    Fringe2 As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Private Type MapRec
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
    MaxX As Byte
    MaxY As Byte
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Integer
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    
    Sprite As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Byte
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    
    Pic As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Private Type MapItemRec
    Num As Byte
    Value As Long
    Dur As Integer
    
    x As Byte
    y As Byte
End Type

Private Type NpcRec
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

Private Type MapNpcRec
    Num As Integer
    
    Target As Integer
    
    Vital(1 To Vitals.Vital_Count - 1) As Long
        
    x As Byte
    y As Byte
    Dir As Integer
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
End Type

Private Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 255
    LeaveSay As String * 255
    FixesItems As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type
    
Private Type SpellRec
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

Private Type TempTileRec
    DoorOpen() As Byte
    DoorTimer As Long
End Type

Private Type MapDataRec
    Npc() As MapNpcRec
End Type
