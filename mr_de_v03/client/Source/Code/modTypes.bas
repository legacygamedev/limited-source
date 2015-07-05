Attribute VB_Name = "modTypes"
Option Explicit

Type AnimationRec
    Name            As String * NAME_LENGTH
    Animation       As Long
    AnimationFrames As Long
    AnimationSpeed  As Long
    AnimationSize   As Long
    AnimationLayer  As Long
End Type

Type PositionRec
    Map As Long
    X As Byte
    Y As Byte
End Type

Type AnimRec
    Anim As Byte
    MaxFrames As Byte
    Speed As Byte
    Size As Byte
    AnimNum As Byte
    Created As Long
    CurrFrame As Byte
    X As Long
    Y As Long
    Layer As Long
End Type

' Used for ActionMsg Messages
Type ActionMsgRec
    Message As String
    Created As Long
    Type As Long
    Color As Long
    Scroll As Long
    X As Long
    Y As Long
End Type

Type PlayerInvRec
    Num As Byte
    Value As Long
    Bound As Boolean
End Type

Type PlayerSpellRec
    SpellNum    As Long
    Cooldown    As Long
End Type

Type QuestProgressUDT
    QuestNum As Long
    Progress(1 To MAX_QUEST_NEEDS) As Long
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
    Vital(1 To Vitals.Vital_Count) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count) As Long
    Points As Long
    
    ' Worn equipment
    Equipment(1 To Slots.Slot_Count) As Long
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As PlayerSpellRec

    ' Position
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte
    
    ' For Death
    IsDead   As Boolean
    IsDeadTimer As Long
    
    ' For Quests
    ActiveQuestCount As Long
    CompletedQuests(1 To MAX_QUESTS) As Long
    QuestProgress(1 To MAX_PLAYER_QUESTS) As QuestProgressUDT
    
    ' Client use only
    MaxVital(1 To Vitals.Vital_Count) As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
    GuildName As String
    GuildAbbreviation As String
    NextLevel As Long
    
    ModStat(1 To Stats.Stat_Count) As Long
    
    EmoticonNum As Long
    EmoticonTime As Long
End Type

Type MobsRec
    NpcCount As Long    ' Count of npcs - used in total map_npc_count
    Npc() As Long       ' List of all npcs for this grouping
End Type

Type TileRec
    Ground As Integer
    Mask As Integer
    Anim As Integer
    Mask2 As Integer
    M2Anim As Integer
    Fringe As Integer
    FAnim As Integer
    Fringe2 As Integer
    F2Anim As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
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
    TileSet As Byte
    MaxX As Byte
    MaxY As Byte
    Tile() As TileRec
    Mobs(1 To MAX_MOBS) As MobsRec
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    
    MaleSprite As String
    FemaleSprite As String
    
    Vital(1 To Vitals.Vital_Count) As Long
    Stat(1 To Stats.Stat_Count) As Long
    BaseDodge As Long
    BaseCrit As Long
    BaseBlock As Long
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    Pic As Integer
    
    LevelReq As Long
    ClassReq As Integer ' Flags for each class
    StatReq(1 To Stats.Stat_Count) As Byte
    
    Type As Byte
    Rarity As Byte
    Bound As Long
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    
    ModVital(1 To Vitals.Vital_Count) As Long
    ModStat(1 To Stats.Stat_Count) As Long
    
    Stack As Byte
    StackMax As Long
End Type

Type MapItemRec
    Num As Long
    Value As Long
    X As Long
    Y As Long
End Type

Type DropItemRec
    Item As Long
    ItemValue As Long
    Chance As Byte
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 255
    
    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    MovementSpeed As Byte
    MovementFrequency As Byte
    
    Drop(1 To 4) As DropItemRec
    
    Stat(1 To Stats.Stat_Count) As Long
    MaxHP As Long
    MaxEXP As Long
    Level As Long
End Type

Type MapNpcRec
    Num As Long
    
    Target As Byte
    
    HP As Long
    MP As Long
    SP As Long
        
    Map As Integer
    X As Long
    Y As Long
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
    JoinSay As String * 500
    Type As Long                ' What type
    ' Below are used for inn
    BindPoint As PositionRec    ' Used for home point setting
    
    ' Below are used for shops
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Type SpellRec
    Name As String * NAME_LENGTH
    
    Type As Byte
    Range As Byte
    
    LevelReq As Byte
    ClassReq As Long
    VitalReq(1 To Vitals.Vital_Count) As Long
    
    TargetFlags As Long ' Refer to the "Targets" enum for flag bits
    
    CastTime As Long
    Cooldown As Long
    
    ModVital(1 To Vitals.Vital_Count) As Long
    ModStat(1 To Stats.Stat_Count) As Long
    
    TickCount As Long   ' overtime spells-how many ticks         buffs-always set to 1
    TickUpdate As Long  ' overtime spells-time in between ticks  buffs-total length of buff
    
    Animation As Long       ' What animation to play
End Type

Type EmoRec
    Pic As Long
    Command As String * NAME_LENGTH
End Type

Type TempTileRec
    Open As Boolean
End Type

Public Map As MapRec
'Public TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc() As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Emoticons(1 To MAX_EMOTICONS) As EmoRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
' Map for local use
Public SaveMap As MapRec
Public ActionMsg(0 To MAX_BYTE) As ActionMsgRec
Public Anim(0 To MAX_BYTE) As AnimRec

'Public Tiles(((MAX_MAPX + 1) * (MAX_MAPY + 1)) - 1) As TilesRec
'Public Tiles() As TilesRec

Public Sub ClearActionMsg(ByVal Index As Byte)
    ZeroMemory ByVal VarPtr(ActionMsg(Index)), LenB(ActionMsg(Index))
    ActionMsg(Index).Message = vbNullString
End Sub

Public Sub ClearAnim(ByVal AnimNum As Byte)
    ZeroMemory ByVal VarPtr(Anim(AnimNum)), LenB(Anim(AnimNum))
End Sub

Sub ClearTempTile()
Dim X As Long, Y As Long

    ReDim TempTile(0 To Map.MaxX, 0 To Map.MaxY)
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            TempTile(X, Y).Open = False
        Next
    Next
End Sub

Sub ClearPlayer(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Player(Index)), LenB(Player(Index))
    Player(Index).Name = vbNullString
    Player(Index).GuildName = vbNullString
    Player(Index).GuildAbbreviation = vbNullString
    
    Player(Index).EmoticonNum = -1
End Sub

Sub ClearItem(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Item(Index)), LenB(Item(Index))
    Item(Index).Name = vbNullString
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
End Sub

Sub ClearMapItem(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index))
End Sub

Sub ClearMaps()
Dim i As Long
Dim X As Long
Dim Y As Long

    '''''''
    ' MAP '
    '''''''
    ZeroMemory ByVal VarPtr(Map), LenB(Map)
    Map.Name = vbNullString
    
    ' set the min value for Maxx and Maxy
    Map.MaxX = MAX_MAPX
    Map.MaxY = MAX_MAPY
    
    ' set the Tile()
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    
    
    '''''''''''
    ' SAVEMAP '
    '''''''''''
    ZeroMemory ByVal VarPtr(SaveMap), LenB(SaveMap)
    SaveMap.Name = vbNullString
    
        ' set the min value for Maxx and Maxy
    SaveMap.MaxX = MAX_MAPX
    SaveMap.MaxY = MAX_MAPY
    
    ' set the Tile()
    ReDim SaveMap.Tile(0 To SaveMap.MaxX, 0 To SaveMap.MaxY)
End Sub

Sub ClearMapItems()
Dim X As Long

    For X = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(X)
    Next
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index))
End Sub

Sub ClearMapNpcs()
Dim i As Long
    'ReDim Preserve MapNpc(MapNpcCount)
    For i = 1 To MapNpcCount
        Call ClearMapNpc(i)
    Next
End Sub

Function Current_Name(ByVal Index As Long) As String
    Current_Name = Trim$(Player(Index).Name)
End Function

Sub Update_Name(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function Current_Class(ByVal Index As Long) As Long
    Current_Class = Player(Index).Class
End Function

Sub Update_Class(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function Current_Sprite(ByVal Index As Long) As Long
    Current_Sprite = Player(Index).Sprite
End Function

Sub Update_Sprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
End Sub

Function Current_Level(ByVal Index As Long) As Long
    Current_Level = Player(Index).Level
End Function

Sub Update_Level(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
End Sub

Function Current_NextLevel(ByVal Index As Long) As Long
    Current_NextLevel = Player(Index).NextLevel
End Function

Function Current_Exp(ByVal Index As Long) As Long
    Current_Exp = Player(Index).Exp
End Function

Sub Update_Exp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Exp = Exp
End Sub

Function Current_Access(ByVal Index As Long) As Long
    Current_Access = Player(Index).Access
End Function

Sub Update_Access(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function Current_PK(ByVal Index As Long) As Long
    Current_PK = Player(Index).PK
End Function

Sub Update_PK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Public Function Current_Vital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Current_Vital = Player(Index).Vital(Vital)
End Function

Public Sub Update_Vital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Vital(Vital) = Value
End Sub

Function Current_MaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Current_MaxVital = Player(Index).MaxVital(Vital)
End Function

Public Sub Update_MaxVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).MaxVital(Vital) = Value
End Sub

Public Function Current_BaseStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    Current_BaseStat = Player(Index).Stat(Stat)
End Function

Public Sub Update_BaseStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(Index).Stat(Stat) = Value
End Sub

Public Function Current_ModStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    Current_ModStat = Player(Index).ModStat(Stat)
End Function

Public Sub Update_ModStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(Index).ModStat(Stat) = Value
End Sub

Public Function Current_Stat(ByVal Index As Long, ByVal Stat As Stats) As Long
    Current_Stat = Current_BaseStat(Index, Stat) + Current_ModStat(Index, Stat)
End Function

Function Current_Points(ByVal Index As Long) As Long
    Current_Points = Player(Index).Points
End Function

Sub Update_Points(ByVal Index As Long, ByVal Points As Long)
    Player(Index).Points = Points
End Sub

Function Current_Map(ByVal Index As Long) As Long
    Current_Map = Player(Index).Map
End Function

Sub Update_Map(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
End Sub

Function Current_X(ByVal Index As Long) As Long
    Current_X = Player(Index).X
End Function

Sub Update_X(ByVal Index As Long, ByVal X As Long)
    Player(Index).X = X
End Sub

Function Current_Y(ByVal Index As Long) As Long
    Current_Y = Player(Index).Y
End Function

Sub Update_Y(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Y = Y
End Sub

Function Current_Dir(ByVal Index As Long) As Byte
    Current_Dir = Player(Index).Dir
End Function

Sub Update_Dir(ByVal Index As Long, ByVal Dir As Byte)
    Player(Index).Dir = Dir
End Sub

Function Current_InvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    Current_InvItemNum = Player(Index).Inv(InvSlot).Num
End Function

Sub Update_InvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(InvSlot).Num = ItemNum
End Sub

Function Current_InvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    Current_InvItemValue = Player(Index).Inv(InvSlot).Value
End Function

Sub Update_InvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(InvSlot).Value = ItemValue
End Sub

'***************************************
' Inv Item Bound
'***************************************
Public Function Current_InvItemBound(ByVal Index As Long, ByVal InvSlot As Long) As Boolean
    Current_InvItemBound = Player(Index).Inv(InvSlot).Bound
End Function
Sub Update_InvItemBound(ByVal Index As Long, ByVal InvSlot As Long, ByVal Bound As Boolean)
    Player(Index).Inv(InvSlot).Bound = Bound
End Sub

Public Function Current_EquipmentSlot(ByVal Index As Long, ByVal EquipmentSlot As Slots) As Long
    Current_EquipmentSlot = Player(Index).Equipment(EquipmentSlot)
End Function

Public Sub Update_EquipmentSlot(ByVal Index As Long, ByVal EquipmentSlot As Slots, ByVal InvNum As Long)
    Player(Index).Equipment(EquipmentSlot) = InvNum
End Sub

Public Function Current_Damage(ByVal Index As Long) As Long
    Current_Damage = Clamp((Current_Stat(Index, Stats.Strength) \ 2) + (Current_Stat(Index, Stats.Dexterity) \ 5.5), 0, MAX_LONG)
End Function

Public Function Current_Protection(ByVal Index As Long) As Long
    Current_Protection = Clamp(Current_Stat(Index, Stats.Vitality) \ 2.5, 0, MAX_LONG)
End Function

Public Function Current_MagicDamage(ByVal Index As Long) As Long
    Current_MagicDamage = Clamp(((Current_Stat(Index, Stats.Intelligence) \ 2) + (Current_Stat(Index, Stats.Wisdom) \ 5.5)) \ 2, 0, MAX_LONG)
End Function

Public Function Current_MagicProtection(ByVal Index As Long) As Long
    Current_MagicProtection = Clamp(((Current_Stat(Index, Stats.Intelligence) \ 5.5) + (Current_Stat(Index, Stats.Wisdom) \ 5.5)) \ 2, 0, MAX_LONG)
End Function



Function Current_GuildName(ByVal Index As Long) As String
    Current_GuildName = Trim$(Player(Index).GuildName)
End Function

Sub Update_GuildName(ByVal Index As Long, ByVal Guild As String)
    Player(Index).GuildName = Guild
End Sub

Function Current_GuildAbbreviation(ByVal Index As Long) As String
    Current_GuildAbbreviation = Trim$(Player(Index).GuildAbbreviation)
End Function

Sub Update_GuildAbbreviation(ByVal Index As Long, ByVal GuildAbbreviation As String)
    Player(Index).GuildAbbreviation = GuildAbbreviation
End Sub

Function Current_IsDead(ByVal Index) As Boolean
    Current_IsDead = Player(Index).IsDead
End Function
Sub Update_IsDead(ByVal Index As Long, ByVal Dead As Boolean)
    Player(Index).IsDead = Dead
    ' Display release box
    If Index = MyIndex Then
        If Current_IsDead(MyIndex) Then
            ' Display the death stuff
            AddText "You are dead. Type /release to be released or wait for someone to revive you.", White
            AlertMessage "You are dead. Would you like to be released?", AddressOf Release_Click, False
        End If
    End If
End Sub

Function Current_IsDeadTimer(ByVal Index) As Long
    Current_IsDeadTimer = Player(Index).IsDeadTimer
End Function
Sub Update_IsDeadTimer(ByVal Index As Long, ByVal Value As Long)
    Player(Index).IsDeadTimer = Value
End Sub
