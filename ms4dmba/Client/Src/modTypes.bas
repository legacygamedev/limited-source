Attribute VB_Name = "modTypes"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

' Public data structures
Public Map As MapRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec

Public Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Private Type SpellAnim
    SpellNum As Integer
    Timer As Long
    FramePointer As Long
End Type

Private Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
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

    ' Position
    Map As Integer
    X As Byte
    Y As Byte
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
    
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
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
    
    X As Byte
    Y As Byte
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
    Num As Byte
    
    Target As Byte
    
    Vital(1 To Vitals.Vital_Count - 1) As Long
        
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte

    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    
    SpellAnimations(1 To MAX_SPELLANIM) As SpellAnim
End Type

Private Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
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
    DoorOpen As Byte
End Type

