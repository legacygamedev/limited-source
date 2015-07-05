Attribute VB_Name = "modTypes"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

' ****************************
' ** Player & Account types **
' ****************************

Type PlayerInvRec
    Num As Long
    Value As Long
    Dur As Long
End Type

Type PlayerSkillRec
    Num As Long
    Level As Long
    EXP As Long
End Type

Type PlayerSpellRec
    Num As Long
    Level As Long
    EXP As Long
End Type

Type PlayerQuestRec
    Num As Long
    SetMap As Long
    SetBy As Long
    Value As Long
    Count As Long
End Type

Type PlayerMapRec
    Num As Long
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Sex As Byte
    Class As Long
    Sprite As Long
    Level As Long
    Access As Byte
    EXP As Long
    PK As Byte
    Guild As Long
    
    ' Vitals
    HP As Long
    MP As Long
    SP As Long
    
    ' Stats
    STR As Long
    DEF As Long
    SPEED As Long
    MAGI As Long
    DEX As Long
    POINTS As Byte
    
    ' Worn equipment
    WeaponSlot As Long
    ArmorSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    AmuletSlot As Long
    RingSlot As Long
    ArrowSlot As Long
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Skills(1 To MAX_PLAYER_SKILLS) As PlayerSkillRec
    Spells(1 To MAX_PLAYER_SPELLS) As PlayerSpellRec
    Quests(1 To MAX_PLAYER_QUESTS) As PlayerQuestRec
    Maps(1 To MAX_PLAYER_MAPS) As PlayerMapRec
End Type
    
Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
       
    ' Characters (we use 0 to prevent a crash that still needs to be figured out)
    Char(0 To MAX_CHARS) As PlayerRec
    
    '' None saved local vars
    Buffer As String
    'IncBuffer As String
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

Type ClassRec
    Name As String * NAME_LENGTH
    
    Sprite As Long
    
    STR As Byte
    DEF As Byte
    SPEED As Byte
    MAGI As Byte
    DEX As Byte
    
    StartMap As Long
    StartX As Byte
    StartY As Byte
End Type


' ***************
' ** Map types **
' ***************

Type RSpawnRec
    RSx As Byte
    RSy As Byte
End Type

Type NSpawnRec
    NSx As Byte
    NSy As Byte
End Type

Type MapItemRec
    Num As Long
    Value As Long
    Dur As Long
    
    x As Byte
    y As Byte
End Type

Type MapNpcRec
    Num As Long
    
    Target As Long
    
    HP As Long
    MP As Long
    SP As Long
        
    x As Byte
    y As Byte
    Dir As Byte
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
End Type

Type MapResourceRec
    Num As Long
    
    HP As Long
    
    x As Byte
    y As Byte
    
    ' For server use only
    SpawnWait As Long
End Type

Type TileRec
    Tileset As Byte
    Ground As Byte
    Mask As Byte
    Mask2 As Byte
    Anim As Byte
    Fringe As Byte
    Fringe2 As Byte
    FAnim As Byte
    Light As Byte
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    WalkUp As Byte
    WalkDown As Byte
    WalkLeft As Byte
    WalkRight As Byte
    Build As Byte
End Type

Type MapRec
    Name As String * NAME_LENGTH
    Owner As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    Music As Byte
    BootMap As Long
    BootX As Byte
    BootY As Byte
    Indoors As Byte
    'Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
    Resource(1 To MAX_MAP_RESOURCES) As Long
    NSpawn(1 To MAX_MAP_NPCS) As NSpawnRec
    RSpawn(1 To MAX_MAP_RESOURCES) As RSpawnRec
    Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
End Type

Type TempTileRec
    DoorOpen() As Byte
    DoorTimer As Long
End Type

Type PushTileRec
    Pushed() As Byte
    PushedTimer As Long
End Type


' *************************
' ** Party & Guild types **
' *************************

Type GuildRec
    Name As String * NAME_LENGTH
    Founder As String * NAME_LENGTH
    Member() As String * NAME_LENGTH
    Map(1 To MAX_GUILD_MAPS) As Long
    Quests(1 To MAX_GUILD_QUESTS) As Byte
End Type


' *******************
' ** General types **
' *******************

Type NPCDropRec
    ItemNum As Long
    ItemValue As Long
    Chance As Long
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    STR  As Long
    DEF As Long
    SPEED As Long
    MAGI As Long
    
    Big As Byte
    MaxHp As Long
    Respawn As Long
    HitOnlyWith As Long
    ShopLink As Long
    ExpType As Byte
    EXP As Long
    
    QuestNPC(1 To MAX_NPC_QUESTS) As Long
    ItemNPC(1 To MAX_NPC_DROPS) As NPCDropRec
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
End Type

Type TradeItemRec
    GiveItem(1 To MAX_GIVE_ITEMS) As Long
    GiveValue(1 To MAX_GIVE_VALUE) As Long
    GetItem(1 To MAX_GET_ITEMS) As Long
    GetValue(1 To MAX_GET_VALUE) As Long
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    FixesItems As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
    ItemStock(1 To MAX_TRADES) As Integer
End Type

Type SkillRec
    Name As String * NAME_LENGTH
    SkillSprite As Long
    ClassReq As Long
    LevelReq As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
End Type

Type SpellRec
    Name As String * NAME_LENGTH
    SpellSprite As Long
    ClassReq As Long
    LevelReq As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
End Type

Type QuestRec
    Name As String
    Description As String
    SetBy As Long
    ClassReq As Long
    LevelMin As Long
    LevelMax As Long
    Type As Byte
    Reward As Long
    RewardValue As Integer
    Data1 As Long
    Data2 As Long
    Data3 As Long
End Type

' **********************
' ** Menu & GUI types **
' **********************

Type GUIDataRec
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    Data4 As Integer
End Type

Type GUIBackgroundRec
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    Data4 As Integer
    Data5 As Integer
End Type

Type GUIRec
    Name As String
    Designer As String
    Revision As Long
    Background(1 To 7) As GUIBackgroundRec
    Menu(1 To 5) As GUIDataRec
    Login(1 To 4) As GUIDataRec
    NewAcc(1 To 4) As GUIDataRec
    DelAcc(1 To 4) As GUIDataRec
    Credits(1 To 2) As GUIDataRec
    Chars(1 To 5) As GUIDataRec
    NewChar(1 To 14) As GUIDataRec
End Type
