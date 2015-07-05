Attribute VB_Name = "modTypes"
'   Copyright (c) 2006 Joshua Bendig
'   This file is part of Asgard.
'
'    Asgard is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Asgard is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Asgard; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

' General constants
Public GAME_NAME As String
Public WEBSITE As String
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

Public Const MAX_ARROWS = 100
Public Const MAX_PLAYER_ARROWS = 100

Public Const MAX_INV = 24
Public Const MAX_MAP_NPCS = 15
Public Const MAX_PLAYER_SPELLS = 20
Public Const MAX_TRADES = 66
Public Const MAX_PLAYER_TRADES = 8
Public Const MAX_NPC_DROPS = 10

Public Const NO = 0
Public Const YES = 1

' Account constants
Public Const NAME_LENGTH = 20
Public Const MAX_CHARS = 3

' Basic Security Passwords, You cant connect without it
Public Const SEC_CODE1 = "jwehiehfojcvnvnsdinaoiwheoewyriusdyrflsdjncjkxzncisdughfusyfuapsipiuahfpaijnflkjnvjnuahguiryasbdlfkjblsahgfauygewuifaunfauf"
Public Const SEC_CODE2 = "ksisyshentwuegeguigdfjkldsnoksamdihuehfidsuhdushdsisjsyayejrioehdoisahdjlasndowijapdnaidhaioshnksfnifohaifhaoinfiwnfinsaihfas"
Public Const SEC_CODE3 = "saiugdapuigoihwbdpiaugsdcapvhvinbudhbpidusbnvduisysayaspiufhpijsanfioasnpuvnupashuasohdaiofhaosifnvnuvnuahiosaodiubasdi"
Public Const SEC_CODE4 = "88978465734619123425676749756722829121973794379467987945762347631462572792798792492416127957989742945642672"

' Sex constants
Public Const SEX_MALE = 0
Public Const SEX_FEMALE = 1

' Map constants
'Public Const MAX_MAPX = 30
'Public Const MAX_MAPY = 30
Public MAX_MAPX As Variant
Public MAX_MAPY As Variant
Public Const MAP_MORAL_NONE = 0
Public Const MAP_MORAL_SAFE = 1
Public Const MAP_MORAL_NO_PENALTY = 2

' Image constants
Public Const PIC_X = 32
Public Const PIC_Y = 32

' Tile consants
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
Public Const TILE_TYPE_NPC_SPAWN = 20

' Item constants
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

' Direction constants
Public Const DIR_UP = 0
Public Const DIR_DOWN = 1
Public Const DIR_LEFT = 2
Public Const DIR_RIGHT = 3

' Constants for player movement
Public Const MOVING_WALKING = 1
Public Const MOVING_RUNNING = 2

' Weather constants
Public Const WEATHER_NONE = 0
Public Const WEATHER_RAINING = 1
Public Const WEATHER_SNOWING = 2
Public Const WEATHER_THUNDER = 3

' Time constants
Public Const TIME_DAY = 0
Public Const TIME_NIGHT = 1

' Admin constants
Public Const ADMIN_MONITER = 1
Public Const ADMIN_MAPPER = 2
Public Const ADMIN_DEVELOPER = 3
Public Const ADMIN_CREATOR = 4

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED = 1
Public Const NPC_BEHAVIOR_FRIENDLY = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER = 3
Public Const NPC_BEHAVIOR_GUARD = 4

' Speach bubble constants
Public Const DISPLAY_BUBBLE_TIME = 2000 ' In milliseconds.
Public DISPLAY_BUBBLE_WIDTH As Byte
Public Const MAX_BUBBLE_WIDTH = 6 ' In tiles. Includes corners.
Public Const MAX_LINE_LENGTH = 23 ' In characters.
Public Const MAX_LINES = 3

' Spell constants
Public Const SPELL_TYPE_ADDHP = 0
Public Const SPELL_TYPE_ADDMP = 1
Public Const SPELL_TYPE_ADDSP = 2
Public Const SPELL_TYPE_SUBHP = 3
Public Const SPELL_TYPE_SUBMP = 4
Public Const SPELL_TYPE_SUBSP = 5
Public Const SPELL_TYPE_CONDITION = 6

Type ChatBubble
    Text As String
    Created As Long
End Type

Type PlayerInvRec
    num As Long
    value As Long
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

Type PlayerRec
    ' General
    name As String * NAME_LENGTH
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
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
       
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    
    ' Client use only
    MaxHp As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte
    
    SpellNum As Long
    SpellAnim() As SpellAnimRec

    EmoticonNum As Long
    EmoticonTime As Long
    EmoticonVar As Long
    
    LevelUp As Long
    LevelUpT As Long

    Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec
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

Type MapRec
    name As String * 40
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
End Type

Type ClassRec
    name As String * NAME_LENGTH
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
    name As String * NAME_LENGTH
    desc As String * 150
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    StrReq As Long
    DefReq As Long
    SpeedReq As Long
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
End Type

Type MapItemRec
    num As Long
    value As Long
    Dur As Long
    
    x As Byte
    y As Byte
End Type

Type NPCEditorRec
    ItemNum As Long
    ItemValue As Long
    Chance As Long
End Type

Type NpcRec
    name As String * NAME_LENGTH
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
    MaxHp As Long
    EXP As Long
    SpawnTime As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
End Type

Type MapNpcRec
    num As Long
    
    Target As Long
    
    HP As Long
    MaxHp As Long
    MP As Long
    SP As Long
    
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    Big As Byte

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

Type TradeItemsRec
    value(1 To MAX_TRADES) As TradeItemRec
End Type

Type ShopRec
    name As String * NAME_LENGTH
    JoinSay As String * 100
    LeaveSay As String * 100
    FixesItems As Byte
    TradeItem(1 To 6) As TradeItemsRec
End Type

Type SpellRec
    name As String * NAME_LENGTH
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
End Type

Type TempTileRec
    DoorOpen As Byte
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
End Type
Public Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public Trading2(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

Type EmoRec
    Pic As Long
    Command As String
End Type

Type DropRainRec
    x As Long
    y As Long
    Randomized As Boolean
    speed As Byte
End Type

' Bubble thing
Public Bubble() As ChatBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

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
Public Emoticons() As EmoRec
Public MapReport() As MapRec

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
    name As String
    Pic As Long
    Range As Byte
End Type
Public Arrows(1 To MAX_ARROWS) As ArrowRec

Type BattleMsgRec
    Msg As String
    index As Byte
    Color As Byte
    Time As Long
    Done As Byte
    y As Long
End Type
Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Type ItemDurRec
    Item As Long
    Dur As Long
    Done As Byte
End Type
Public ItemDur(1 To 4) As ItemDurRec

Public Inventory As Long


Public Const MAX_ATTRIBUTE_NPCS = 25
Public MapAttributeNpc() As MapNpcRec
Public SaveMapAttributeNpc() As MapNpcRec
Public charselsprite(MAX_CHARS) As Double


Sub ClearTempTile()
Dim x As Long, y As Long

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            TempTile(x, y).DoorOpen = NO
        Next x
    Next y
End Sub

Sub ClearPlayer(ByVal index As Long)
Dim i As Long
Dim n As Long

    Player(index).name = ""
    Player(index).Guild = ""
    Player(index).Guildaccess = 0
    Player(index).Class = 0
    Player(index).Level = 0
    Player(index).Sprite = 0
    Player(index).EXP = 0
    Player(index).Access = 0
    Player(index).PK = NO
        
    Player(index).HP = 0
    Player(index).MP = 0
    Player(index).SP = 0
        
    Player(index).STR = 0
    Player(index).DEF = 0
    Player(index).speed = 0
    Player(index).MAGI = 0
        
    For n = 1 To MAX_INV
        Player(index).Inv(n).num = 0
        Player(index).Inv(n).value = 0
        Player(index).Inv(n).Dur = 0
    Next n
        
    Player(index).ArmorSlot = 0
    Player(index).WeaponSlot = 0
    Player(index).HelmetSlot = 0
    Player(index).ShieldSlot = 0
        
    Player(index).Map = 0
    Player(index).x = 0
    Player(index).y = 0
    Player(index).Dir = 0
    
    ' Client use only
    Player(index).MaxHp = 0
    Player(index).MaxMP = 0
    Player(index).MaxSP = 0
    Player(index).XOffset = 0
    Player(index).YOffset = 0
    Player(index).Moving = 0
    Player(index).Attacking = 0
    Player(index).AttackTimer = 0
    Player(index).MapGetTimer = 0
    Player(index).CastedSpell = NO
    Player(index).EmoticonNum = -1
    Player(index).EmoticonTime = 0
    Player(index).EmoticonVar = 0
    
    For i = 1 To MAX_SPELL_ANIM
        Player(index).SpellAnim(i).CastedSpell = NO
        Player(index).SpellAnim(i).SpellTime = 0
        Player(index).SpellAnim(i).SpellVar = 0
        Player(index).SpellAnim(i).SpellDone = 0
        
        Player(index).SpellAnim(i).Target = 0
        Player(index).SpellAnim(i).TargetType = 0
    Next i
    
    Player(index).SpellNum = 0
    
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).index = 1
        BattleMMsg(i).Time = i
    Next i
    
    Inventory = 1
End Sub

Sub ClearItem(ByVal index As Long)
    Item(index).name = ""
    Item(index).desc = ""
    
    Item(index).Type = 0
    Item(index).Data1 = 0
    Item(index).Data2 = 0
    Item(index).Data3 = 0
    Item(index).StrReq = 0
    Item(index).DefReq = 0
    Item(index).SpeedReq = 0
    Item(index).ClassReq = -1
    Item(index).AccessReq = 0
    
    Item(index).AddHP = 0
    Item(index).AddMP = 0
    Item(index).AddSP = 0
    Item(index).AddStr = 0
    Item(index).AddDef = 0
    Item(index).AddMagi = 0
    Item(index).AddSpeed = 0
    Item(index).AddEXP = 0
    Item(index).AttackSpeed = 1000
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearMapItem(ByVal index As Long)
    MapItem(index).num = 0
    MapItem(index).value = 0
    MapItem(index).Dur = 0
    MapItem(index).x = 0
    MapItem(index).y = 0
End Sub

Sub ClearMap()
Dim i As Long
Dim x As Long
Dim y As Long

For i = 1 To MAX_MAPS
    Map(i).name = ""
    Map(i).Revision = 0
    Map(i).Moral = 0
    Map(i).Up = 0
    Map(i).Down = 0
    Map(i).Left = 0
    Map(i).Right = 0
    Map(i).Indoors = 0
        
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Map(i).Tile(x, y).Ground = 0
            Map(i).Tile(x, y).Mask = 0
            Map(i).Tile(x, y).Anim = 0
            Map(i).Tile(x, y).Mask2 = 0
            Map(i).Tile(x, y).M2Anim = 0
            Map(i).Tile(x, y).Fringe = 0
            Map(i).Tile(x, y).FAnim = 0
            Map(i).Tile(x, y).Fringe2 = 0
            Map(i).Tile(x, y).F2Anim = 0
            Map(i).Tile(x, y).Type = 0
            Map(i).Tile(x, y).Data1 = 0
            Map(i).Tile(x, y).Data2 = 0
            Map(i).Tile(x, y).Data3 = 0
            Map(i).Tile(x, y).String1 = ""
            Map(i).Tile(x, y).String2 = ""
            Map(i).Tile(x, y).String3 = ""
            Map(i).Tile(x, y).Light = 0
            Map(i).Tile(x, y).GroundSet = 0
            Map(i).Tile(x, y).MaskSet = 0
            Map(i).Tile(x, y).AnimSet = 0
            Map(i).Tile(x, y).Mask2Set = 0
            Map(i).Tile(x, y).M2AnimSet = 0
            Map(i).Tile(x, y).FringeSet = 0
            Map(i).Tile(x, y).FAnimSet = 0
            Map(i).Tile(x, y).Fringe2Set = 0
            Map(i).Tile(x, y).F2AnimSet = 0
        Next x
    Next y
Next i
End Sub

Sub ClearMapItems()
Dim x As Long

    For x = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(x)
    Next x
End Sub

Sub ClearMapAttributeNpc(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    MapAttributeNpc(index, x, y).num = 0
    MapAttributeNpc(index, x, y).Target = 0
    MapAttributeNpc(index, x, y).HP = 0
    MapAttributeNpc(index, x, y).MP = 0
    MapAttributeNpc(index, x, y).SP = 0
    MapAttributeNpc(index, x, y).Map = 0
    MapAttributeNpc(index, x, y).x = 0
    MapAttributeNpc(index, x, y).y = 0
    MapAttributeNpc(index, x, y).Dir = 0
    
    ' Client use only
    MapAttributeNpc(index, x, y).XOffset = 0
    MapAttributeNpc(index, x, y).YOffset = 0
    MapAttributeNpc(index, x, y).Moving = 0
    MapAttributeNpc(index, x, y).Attacking = 0
    MapAttributeNpc(index, x, y).AttackTimer = 0
End Sub

Sub ClearMapAttributeNpcs()
Dim i As Long, x As Long, y As Long
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            For i = 1 To MAX_ATTRIBUTE_NPCS
                Call ClearMapAttributeNpc(i, x, y)
            Next i
        Next x
    Next y
End Sub

Sub ClearMapNpc(ByVal index As Long)
    MapNpc(index).num = 0
    MapNpc(index).Target = 0
    MapNpc(index).HP = 0
    MapNpc(index).MP = 0
    MapNpc(index).SP = 0
    MapNpc(index).Map = 0
    MapNpc(index).x = 0
    MapNpc(index).y = 0
    MapNpc(index).Dir = 0
    
    ' Client use only
    MapNpc(index).XOffset = 0
    MapNpc(index).YOffset = 0
    MapNpc(index).Moving = 0
    MapNpc(index).Attacking = 0
    MapNpc(index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next i
End Sub

Function GetPlayerName(ByVal index As Long) As String
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim(Player(index).name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal name As String)
    Player(index).name = name
End Sub

Function GetPlayerGuild(ByVal index As Long) As String
    GetPlayerGuild = Trim(Player(index).Guild)
End Function

Sub SetPlayerGuild(ByVal index As Long, ByVal Guild As String)
    Player(index).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal index As Long) As Long
    GetPlayerGuildAccess = Player(index).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal Guildaccess As Long)
    Player(index).Guildaccess = Guildaccess
End Sub


Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    GetPlayerSprite = Player(index).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    GetPlayerLevel = Player(index).Level
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
    Player(index).Level = Level
End Sub

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).EXP
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal EXP As Long)
    Player(index).EXP = EXP
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    GetPlayerAccess = Player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    GetPlayerPK = Player(index).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).PK = PK
End Sub

Function GetPlayerHP(ByVal index As Long) As Long
    GetPlayerHP = Player(index).HP
End Function

Sub SetPlayerHP(ByVal index As Long, ByVal HP As Long)
    Player(index).HP = HP
    
    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then
        Player(index).HP = GetPlayerMaxHP(index)
    End If
End Sub

Function GetPlayerMP(ByVal index As Long) As Long
    GetPlayerMP = Player(index).MP
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
    Player(index).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        Player(index).MP = GetPlayerMaxMP(index)
    End If
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
    GetPlayerSP = Player(index).SP
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
    Player(index).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        Player(index).SP = GetPlayerMaxSP(index)
    End If
End Sub

Function GetPlayerMaxHP(ByVal index As Long) As Long
    GetPlayerMaxHP = Player(index).MaxHp
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
    GetPlayerMaxMP = Player(index).MaxMP
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long
    GetPlayerMaxSP = Player(index).MaxSP
End Function

Function GetPlayerSTR(ByVal index As Long) As Long
    GetPlayerSTR = Player(index).STR
End Function

Sub SetPlayerSTR(ByVal index As Long, ByVal STR As Long)
    Player(index).STR = STR
End Sub

Function GetPlayerDEF(ByVal index As Long) As Long
    GetPlayerDEF = Player(index).DEF
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal DEF As Long)
    Player(index).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal index As Long) As Long
    GetPlayerSPEED = Player(index).speed
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal speed As Long)
    Player(index).speed = speed
End Sub

Function GetPlayerMAGI(ByVal index As Long) As Long
    GetPlayerMAGI = Player(index).MAGI
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal MAGI As Long)
    Player(index).MAGI = MAGI
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    GetPlayerPOINTS = Player(index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    Player(index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
If index <= 0 Then Exit Function
    GetPlayerMap = Player(index).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
    Player(index).Map = MapNum
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    GetPlayerX = Player(index).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    Player(index).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    GetPlayerY = Player(index).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    Player(index).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    GetPlayerDir = Player(index).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(index).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(index).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(index).Inv(InvSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(InvSlot).value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(index).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(index).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerArmorSlot(ByVal index As Long) As Long
    GetPlayerArmorSlot = Player(index).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)
    Player(index).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal index As Long) As Long
    GetPlayerWeaponSlot = Player(index).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)
    Player(index).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long
    GetPlayerHelmetSlot = Player(index).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)
    Player(index).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long
    GetPlayerShieldSlot = Player(index).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)
    Player(index).ShieldSlot = InvNum
End Sub

