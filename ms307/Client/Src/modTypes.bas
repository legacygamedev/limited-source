Attribute VB_Name = "modTypes"
Option Explicit


Type PlayerInvRec
    Num As Long
    Value As Long
    Dur As Long
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Class As Byte
    Sprite As Integer
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte
    POINTS As Byte
    
    ' Worn equipment
    ArmorSlot As Long 'Byte
    WeaponSlot As Long 'Byte
    HelmetSlot As Long 'Byte
    ShieldSlot As Long 'Byte
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
       
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
End Type
    
Type TileRec
    Ground As Integer
    Mask As Integer
    Anim As Integer
    Fringe As Integer
    Type As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
End Type

Type MapRec
    Name As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
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

Type ClassRec
    Name As String * NAME_LENGTH
    Sprite As Integer
    
    STR As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    
    Pic As Integer
    Type As Long
    Data1 As Long 'Integer
    Data2 As Long 'Integer
    Data3 As Long 'Integer
    UnBreakable As Long 'Integer
    Locked As Long 'Integer
    Disabled As Long
    Assigned As String * NAME_LENGTH
End Type

Type MapItemRec
    Num As Long 'Byte
    Value As Long 'Long
    Dur As Long 'Integer
    
    X As Byte
    Y As Byte
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Long 'Integer
    DropItem As Long 'Byte
    DropItemValue As Long 'Integer
    
    STR  As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte
End Type

Type MapNpcRec
    Num As Byte
    
    Target As Byte
    
    HP As Long
    MP As Long
    SP As Long
        
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
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
    FixesItems As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Type SpellRec
    Name As String * NAME_LENGTH
    ClassReq As Byte
    LevelReq As Byte
    Type As Long 'Byte
    Data1 As Long 'Integer
    Data2 As Long 'Integer
    Data3 As Long 'Integer
End Type

Type TempTileRec
    DoorOpen As Byte
End Type

Type uAssign
    Name As String
    Access As Long
End Type
Public pAssign() As uAssign

Type uCache
    nDate As String
    nTime As String
    nCache As String
End Type
Public Cache(1 To 1)

' Used for parsing
Public PAD_CHAR As String * 1
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte

Public Map As MapRec
Public TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec

Public MyKEY As String

Sub ClearTempTile()
Dim X As Long, Y As Long

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            TempTile(X, Y).DoorOpen = NO
        Next X
    Next Y
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim I As Long
Dim N As Long

    Player(Index).Name = ""
    Player(Index).Class = 0
    Player(Index).Level = 0
    Player(Index).Sprite = 0
    Player(Index).Exp = 0
    Player(Index).Access = 0
    Player(Index).PK = NO
        
    Player(Index).HP = 0
    Player(Index).MP = 0
    Player(Index).SP = 0
        
    Player(Index).STR = 0
    Player(Index).DEF = 0
    Player(Index).SPEED = 0
    Player(Index).MAGI = 0
        
    For N = 1 To MAX_INV
        Player(Index).Inv(N).Num = 0
        Player(Index).Inv(N).Value = 0
        Player(Index).Inv(N).Dur = 0
    Next N
        
    Player(Index).ArmorSlot = 0
    Player(Index).WeaponSlot = 0
    Player(Index).HelmetSlot = 0
    Player(Index).ShieldSlot = 0
        
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
    Player(Index).Moving = 0
    Player(Index).Attacking = 0
    Player(Index).AttackTimer = 0
    Player(Index).MapGetTimer = 0
    Player(Index).CastedSpell = NO
End Sub

Sub ClearItem(ByVal Index As Long)
With Item(Index)
    .Name = ""
    .Type = 0
    .Data1 = 0
    .Data2 = 0
    .Data3 = 0
    .UnBreakable = 0
    .Locked = 0
    .Disabled = 0
    .Assigned = ""
End With
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

Sub ClearMap()
Dim I As Long
Dim X As Long
Dim Y As Long

    Map.Name = ""
    Map.Revision = 0
    Map.Moral = 0
    Map.Up = 0
    Map.Down = 0
    Map.Left = 0
    Map.Right = 0
        
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map.Tile(X, Y).Ground = 0
            Map.Tile(X, Y).Mask = 0
            Map.Tile(X, Y).Anim = 0
            Map.Tile(X, Y).Fringe = 0
            Map.Tile(X, Y).Type = 0
            Map.Tile(X, Y).Data1 = 0
            Map.Tile(X, Y).Data2 = 0
            Map.Tile(X, Y).Data3 = 0
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

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
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
    GetPlayerExp = Player(Index).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Exp = Exp
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
    GetPlayerSPEED = Player(Index).SPEED
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal SPEED As Long)
    Player(Index).SPEED = SPEED
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

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(InvSlot).Value = ItemValue
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

