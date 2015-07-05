Attribute VB_Name = "Types"
Option Explicit

' *------------*
' | rootSource |
' *------------*

' Public data structures
Public map(1 To 9) As MapRec
Public MapTilePosition(-1 To MAX_MAPX + 1, -1 To MAX_MAPY + 1) As TilePosRec
Public TempTile(-15 To MAX_MAPX + 15, -15 To MAX_MAPY + 15) As TempTileRec
Public Class() As ClassRec
Public Player() As PlayerRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public MapItem() As MapItemRec
Public MapNpc() As MapNpcRec
Public GameData As DataRec
Public Lookup() As TileLookupStruct

Public Type TileLookupStruct
    map As Long
    x As Long
    y As Long
End Type

Public Type RECTx
    xpos As Long
    ypos As Long
End Type

Type DataRec
  IP As String * NAME_LENGTH
  Port As Integer
  Autoupdater As Byte
  SaveLogin As Byte
  Username As String * NAME_LENGTH
  Password As String * NAME_LENGTH
  Music As Byte
  Sound As Byte
  PlayerNames As Byte
  NpcNames As Byte
  SpellGFX As Byte
  MusicExt As String * 5
  ScreenNum As Integer
  WebAddress As String * 500
  GameName As String * 255
  VerProcess As Long
End Type

Public Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Public Type SpellAnim
    spellnum As Integer
    Timer As Long
    FramePointer As Long
End Type
Public Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
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

    ' Position
    map As Integer
    x As Long
    y As Long
    Dir As Byte
    
    ' Client use only
    MaxHP As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
    
    SpellAnimations(1 To MAX_SPELLANIM) As SpellAnim
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
    
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
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
    
    x As Byte
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
    Num As Byte
    
    Target As Byte
    
    Vital(1 To Vitals.Vital_Count - 1) As Long
        
    map As Integer
    x As Byte
    y As Byte
    Dir As Byte

    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Long
    Attacking As Byte
    AttackTimer As Long
    
    SpellAnimations(1 To MAX_SPELLANIM) As SpellAnim
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
    DoorOpen As Byte
End Type


Public Type VERTEX
    x As Single
    y As Single
    Z As Single
    RHW As Single
    Color As Long
    Specular As Long
    TU As Single
    TV As Single
End Type

Public Type Quad
    Vertice(0 To 3) As VERTEX
End Type

Public Type TilePosRec
    PosX As Long
    PosY As Long
    Texture As Long
    Layer(0 To 8) As RECTx
End Type


