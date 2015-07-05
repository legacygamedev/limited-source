Attribute VB_Name = "modTypes"

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.
Option Explicit

Global PlayerI As Byte

' Timers enabled?
Public PlayerTimer As Byte
Public tmrPlayerSave As Byte

' Winsock globals
Public GAME_PORT As Long

' Server started when?
Public ServerStartTime As Long
Public ServerOn As Byte
Public LastGetTickCount As Long

' Used with sending songs, what type of SendData to use
Public Const SDT As Byte = 1 'SendDataTo
Public Const SDTM As Byte = 2 'SendDataToMap
Public Const SDTAB As Byte = 3 'SendDataToAllBut

' The sound IDs
Public Const MAGIC_SOUND As Byte = 1
Public Const THUNDER_SOUND As Byte = 2
Public Const SERVERSHUTDOWN_SOUND As Byte = 3
Public Const DEAD_SOUND As Byte = 4
Public Const PAIN_SOUND As Byte = 5
Public Const NEWLEVEL_SOUND As Byte = 6
Public Const LOGINTOSERVER_SOUND As Byte = 7
Public Const PLAYERJOINED_SOUND As Byte = 8
Public Const LOGOUTOFSERVER_SOUND As Byte = 9
Public Const PLAYERHASLEFT_SOUND As Byte = 10
Public Const KEY_SOUND As Byte = 11
Public Const WARP_SOUND As Byte = 12
Public Const NEWVERSIONRELEASED_SOUND As Byte = 13
Public Const MISS_SOUND As Byte = 14
Public Const STRENGTHRAISED_SOUND As Byte = 15
Public Const DEFENSERAISED_SOUND As Byte = 16
Public Const MAGICRAISED_SOUND As Byte = 17
Public Const SPEEDRAISED_SOUND As Byte = 18
Public Const ATTACK_SOUND As Byte = 19
Public Const CRITICALHIT_SOUND As Byte = 20

' Data coming in and out
Public DataKBIn As Long
Public DataKBOut As Long

' GetVar variables
Public HPRegenOn As Byte
Public MPRegenOn As Byte
Public SPRegenOn As Byte

' General constants
Public GAME_NAME As String
Public MAX_PLAYERS As Long
Public MAX_SPELLS As Long
Public MAX_MAPS As Long
Public MAX_SHOPS As Long
Public MAX_ITEMS As Long
Public MAX_NPCS As Long
Public MAX_MAP_ITEMS As Long
Public MAX_GUILDS As Long
Public MAX_GUILD_MEMBERS As Long
Public MAX_EMOTICONS As Long
Public MAX_LEVEL As Long
Public SCRIPTING As Long
Public MAX_PARTIES As Long
Public MAX_PARTY_MEMBERS As Long
Public MAX_SPEECH As Long

Public Const MAX_ARROWS As Byte = 100
Public Const MAX_SPEECH_OPTIONS As Byte = 20
Public Const MAX_INV As Byte = 24
Public Const MAX_MAP_NPCS As Byte = 15
Public Const MAX_PLAYER_SPELLS As Byte = 20
Public Const MAX_TRADES As Byte = 66
Public Const MAX_PLAYER_TRADES As Byte = 8
Public Const MAX_NPC_DROPS As Byte = 10
Public Const MAX_FRIENDS As Byte = 20
Public Const NO  As Byte = 0
Public Const YES As Byte = 1

' Account constants
Public Const NAME_LENGTH As Byte = 20
Public Const MAX_CHARS As Byte = 3

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
'Public Const MAX_MAPX = 30
'Public Const MAX_MAPY = 30
Public MAX_MAPX As Variant
Public MAX_MAPY As Variant

Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_NO_PENALTY As Byte = 2

' Image constants
Public Const PIC_X As Long = 32
Public Const PIC_Y As Long = 32

' Monster Constants
Public Const MON_X As Byte = 32
Public Const MON_Y As Byte = 32

' Packet constants
Public MAXINFO_CHAR As String * 1
Public INFO_CHAR As String * 1
Public NPCHP_CHAR As String * 1
Public ALERTMSG_CHAR As String * 1
Public PLAINMSG_CHAR As String * 1
Public ALLCHARS_CHAR As String * 1
Public LOGINOK_CHAR As String * 1
Public NEWCHARCLASSES_CHAR As String * 1
Public CLASSESDATA_CHAR As String * 1
Public INGAME_CHAR As String * 1
Public PLAYERINV_CHAR As String * 1
Public PLAYERINVUPDATE_CHAR As String * 1
Public PLAYERWORNEQ_CHAR As String * 1
Public PLAYERPOINTS_CHAR As String * 1
Public PLAYERHP_CHAR As String * 1
Public PETHP_CHAR As String * 1
Public PLAYERMP_CHAR As String * 1
Public MAPMSG2_CHAR As String * 1
Public PLAYERSP_CHAR As String * 1
Public PLAYERSTATSPACKET_CHAR As String * 1
Public PLAYERDATA_CHAR As String * 1
Public PETDATA_CHAR As String * 1
Public PLAYERMOVE_CHAR As String * 1
Public PETMOVE_CHAR As String * 1
Public NPCMOVE_CHAR As String * 1
Public PLAYERDIR_CHAR As String * 1
Public NPCDIR_CHAR As String * 1
Public PLAYERXY_CHAR As String * 1
Public ATTACKPLAYER_CHAR As String * 1
Public ATTACKNPC_CHAR As String * 1
Public PETATTACKNPC_CHAR As String * 1
Public NPCATTACK_CHAR As String * 1
Public NPCATTACKPET_CHAR As String * 1
Public CHECKFORMAP_CHAR As String * 1
Public MAPDATA_CHAR As String * 1
Public MAPITEMDATA_CHAR As String * 1
Public MAPNPCDATA_CHAR As String * 1
Public MAPDONE_CHAR As String * 1
Public SAYMSG_CHAR As String * 1
Public SPAWNITEM_CHAR As String * 1
Public ITEMEDITOR_CHAR As String * 1
Public UPDATEITEM_CHAR As String * 1
Public EDITITEM_CHAR As String * 1
Public SPAWNNPC_CHAR As String * 1
Public NPCDEAD_CHAR As String * 1
Public NPCEDITOR_CHAR As String * 1
Public UPDATENPC_CHAR As String * 1
Public EDITNPC_CHAR As String * 1
Public MAPKEY_CHAR As String * 1
Public EDITMAP_CHAR As String * 1
Public SHOPEDITOR_CHAR As String * 1
Public UPDATESHOP_CHAR As String * 1
Public EDITSHOP_CHAR As String * 1
Public MAINEDITOR_CHAR As String * 1
Public SPELLEDITOR_CHAR As String * 1
Public UPDATESPELL_CHAR As String * 1
Public EDITSPELL_CHAR As String * 1
Public TRADE_CHAR As String * 1
Public STARTSPEECH_CHAR As String * 1
Public SPELLS_CHAR As String * 1
Public WEATHER_CHAR As String * 1
Public TIME_CHAR As String * 1
Public ONLINELIST_CHAR As String * 1
Public BLITPLAYERDMG_CHAR As String * 1
Public BLITNPCDMG_CHAR As String * 1
Public PPTRADING_CHAR As String * 1
Public QTRADE_CHAR As String * 1
Public UPDATETRADEITEM_CHAR As String * 1
Public TRADING_CHAR As String * 1
Public PPCHATING_CHAR As String * 1
Public QCHAT_CHAR As String * 1
Public SENDCHAT_CHAR As String * 1
Public SOUND_CHAR As String * 1
Public SPRITECHANGE_CHAR As String * 1
Public CHANGEDIR_CHAR As String * 1
Public CHANGEPETDIR_CHAR As String * 1
Public FLASHEVENT_CHAR As String * 1
Public PROMPT_CHAR As String * 1
Public SPEECHEDITOR_CHAR As String * 1
Public SPEECH_CHAR As String * 1
Public EDITSPEECH_CHAR As String * 1
Public EMOTICONEDITOR_CHAR As String * 1
Public UPDATEEMOTICON_CHAR As String * 1
Public EDITEMOTICON_CHAR As String * 1
Public CLEARTEMPTILE_CHAR As String * 1
Public FRIENDLIST_CHAR As String * 1
Public ARROWEDITOR_CHAR As String * 1
Public UPDATEARROW_CHAR As String * 1
Public EDITARROW_CHAR As String * 1
Public CHECKARROWS_CHAR As String * 1
Public CHECKSPRITE_CHAR As String * 1
Public MAPREPORT_CHAR As String * 1
Public SPELLANIM_CHAR As String * 1
Public CHECKEMOTICONS_CHAR As String * 1
Public DAMAGEDISPLAY_CHAR As String * 1
Public ITEMBREAK_CHAR As String * 1
Public GETINFO_CHAR As String * 1
Public GATCLASSES_CHAR As String * 1
Public NEWFACCOUNTIED_CHAR As String * 1
Public DELIMACCOUNTED_CHAR As String * 1
Public LOGINATION_CHAR As String * 1
Public ADDACHARA_CHAR As String * 1
Public DELIMBOCHARU_CHAR As String * 1
Public USAGAKRIM_CHAR As String * 1
Public GUILDCHANGEACCESS_CHAR As String * 1
Public GUILDDISOWN_CHAR As String * 1
Public GUILDLEAVE_CHAR As String * 1
Public MAKEGUILD_CHAR As String * 1
Public GUILDMEMBER_CHAR As String * 1
Public GUILDTRAINEE_CHAR As String * 1
Public EMOTEMSG_CHAR As String * 1
Public BROADCASTMSG_CHAR As String * 1
Public GLOBALMSG_CHAR As String * 1
Public ADMINMSG_CHAR As String * 1
Public PLAYERMSG_CHAR As String * 1
Public USEITEM_CHAR As String * 1
Public ATTACK_CHAR As String * 1
Public USESTATPOINT_CHAR As String * 1
Public PLAYERINFOREQUEST_CHAR As String * 1
Public SETSPRITE_CHAR As String * 1
Public SETPLAYERSPRITE_CHAR As String * 1
Public GETSTATS_CHAR As String * 1
Public REQUESTNEWMAP_CHAR As String * 1
Public NEEDMAP_CHAR As String * 1
Public MAPGETITEM_CHAR As String * 1
Public MAPDROPITEM_CHAR As String * 1
Public MAPRESPAWN_CHAR As String * 1
Public KICKPLAYER_CHAR As String * 1
Public BANLIST_CHAR As String * 1
Public BANDESTROY_CHAR As String * 1
Public BANPLAYER_CHAR As String * 1
Public REQUESTEDITMAP_CHAR As String * 1
Public REQUESTEDITITEM_CHAR As String * 1
Public SAVEITEM_CHAR As String * 1
Public REQUESTEDITNPC_CHAR As String * 1
Public SAVENPC_CHAR As String * 1
Public REQUESTEDITQUEST_CHAR As String * 1
Public REQUESTEDITSHOP_CHAR As String * 1
Public ADDFRIEND_CHAR As String * 1
Public REMOVEFRIEND_CHAR As String * 1
Public SAVESHOP_CHAR As String * 1
Public REQUESTEDITMAIN_CHAR As String * 1
Public REQUESTEDITSPELL_CHAR As String * 1
Public SAVESPELL_CHAR As String * 1
Public SETACCESS_CHAR As String * 1
Public WHOSONLINE_CHAR As String * 1
Public SETMOTD_CHAR As String * 1
Public TRADEREQUEST_CHAR As String * 1
Public FIXITEM_CHAR As String * 1
Public SEARCH_CHAR As String * 1
Public PLAYERCHAT_CHAR As String * 1
Public ACHAT_CHAR As String * 1
Public DCHAT_CHAR As String * 1
Public ATRADE_CHAR As String * 1
Public DTRADE_CHAR As String * 1
Public UPDATETRADEINV_CHAR As String * 1
Public SWAPITEMS_CHAR As String * 1
Public PARTY_CHAR As String * 1
Public JOINPARTY_CHAR As String * 1
Public LEAVEPARTY_CHAR As String * 1
Public PARTYCHAT_CHAR As String * 1
Public GUILDCHAT_CHAR As String * 1
Public NEWMAIN_CHAR As String * 1
Public REQUESTBACKUPMAIN_CHAR As String * 1
Public CAST_CHAR As String * 1
Public REQUESTLOCATION_CHAR As String * 1
Public KILLPET_CHAR As String * 1
Public REFRESH_CHAR As String * 1
Public PETMOVESELECT_CHAR As String * 1
Public BUYSPRITE_CHAR As String * 1
Public CHECKCOMMANDS_CHAR As String * 1
Public REQUESTEDITARROW_CHAR As String * 1
Public SAVEARROW_CHAR As String * 1
Public SPEECHSCRIPT_CHAR As String * 1
Public REQUESTEDITSPEECH_CHAR As String * 1
Public SAVESPEECH_CHAR As String * 1
Public NEEDSPEECH_CHAR As String * 1
Public REQUESTEDITEMOTICON_CHAR As String * 1
Public SAVEEMOTICON_CHAR As String * 1
Public GMTIME_CHAR As String * 1
Public WARPTO_CHAR As String * 1
Public WARPTOME_CHAR As String * 1
Public WARPPLAYER_CHAR As String * 1
Public ARROWHIT_CHAR As String * 1
Public PPCHATTING_CHAR As String * 1
Public TEMPTILE_CHAR As String * 1
Public TEMPATTRIBUTE_CHAR As String * 1
Public LEVELUP_CHAR As String * 1
Public GATGLASSES_CHAR As String * 1
Public USAGAKARIM_CHAR As String * 1
Public MAPMSG_CHAR As String * 1
Public PPTRADE_CHAR As String * 1
Public NEWPARTY_CHAR As String * 1
Public FORGETSPELL_CHAR As String * 1
Public RETURNSCRIPT_CHAR As String * 1
'Public CLOSINGDOWN_CHAR As String * 1

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_HEAL As Byte = 7
Public Const TILE_TYPE_KILL As Byte = 8
Public Const TILE_TYPE_SHOP As Byte = 9
Public Const TILE_TYPE_CBLOCK As Byte = 10
Public Const TILE_TYPE_ARENA As Byte = 11
Public Const TILE_TYPE_SOUND As Byte = 12
Public Const TILE_TYPE_SPRITE_CHANGE As Byte = 13
Public Const TILE_TYPE_SIGN As Byte = 14
Public Const TILE_TYPE_DOOR As Byte = 15
Public Const TILE_TYPE_NOTICE As Byte = 16
Public Const TILE_TYPE_CHEST As Byte = 17
Public Const TILE_TYPE_CLASS_CHANGE As Byte = 18
Public Const TILE_TYPE_SCRIPTED As Byte = 19
Public Const TILE_TYPE_NONE As Byte = 20

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_POTIONADDHP As Byte = 5
Public Const ITEM_TYPE_POTIONADDMP As Byte = 6
Public Const ITEM_TYPE_POTIONADDSP As Byte = 7
Public Const ITEM_TYPE_POTIONSUBHP As Byte = 8
Public Const ITEM_TYPE_POTIONSUBMP As Byte = 9
Public Const ITEM_TYPE_POTIONSUBSP As Byte = 10
Public Const ITEM_TYPE_KEY As Byte = 11
Public Const ITEM_TYPE_CURRENCY As Byte = 12
Public Const ITEM_TYPE_SPELL As Byte = 13
Public Const ITEM_TYPE_PET As Byte = 14

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Weather constants
Public Const WEATHER_NONE As Byte = 0
Public Const WEATHER_RAINING As Byte = 1
Public Const WEATHER_SNOWING As Byte = 2
Public Const WEATHER_THUNDER As Byte = 3

' Time constants
Public Const TIME_DAY As Byte = 0
Public Const TIME_NIGHT As Byte = 1

' Admin constants
Public Const ADMIN_MONITER As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOR_GUARD As Byte = 4

' Spell constants
Public Const SPELL_TYPE_ADDHP As Byte = 0
Public Const SPELL_TYPE_ADDMP As Byte = 1
Public Const SPELL_TYPE_ADDSP As Byte = 2
Public Const SPELL_TYPE_SUBHP As Byte = 3
Public Const SPELL_TYPE_SUBMP As Byte = 4
Public Const SPELL_TYPE_SUBSP As Byte = 5

'Public Const SPELL_TYPE_GIVEITEM = 6
Public Const SPELL_TYPE_PET As Byte = 6

' Target type constants
Public Const TARGET_TYPE_PLAYER As Byte = 0
Public Const TARGET_TYPE_NPC As Byte = 1
Public Const TARGET_TYPE_LOCATION As Byte = 2
Public Const TARGET_TYPE_PET As Byte = 3

' Emoticon type constants
Public Const EMOTICON_TYPE_IMAGE As Byte = 0
Public Const EMOTICON_TYPE_SOUND As Byte = 1
Public Const EMOTICON_TYPE_BOTH As Byte = 2

Type PlayerInvRec
    num As Long
    Value As Long
    Dur As Long
End Type

Type PetRec
    Sprite As Long
    Alive As Byte
    Map As Long
    X As Long
    Y As Long
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

Type PlayerRec

    ' General
    Name As String * NAME_LENGTH
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
    X As Byte
    Y As Byte
    Dir As Byte
    Friends(1 To MAX_FRIENDS) As String
    
    LoggedIn As Byte
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
End Type

Type PartyRec
    Member() As Long
End Type

Type AccountRec

    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH

    ' Characters (we use 0 to prevent a crash that still needs to be figured out)
    Char(0 To MAX_CHARS) As PlayerRec

    ' None saved local vars
    Buffer As String
    IncBuffer As String
    CharNum As Byte
    InGame As String * 3
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
    Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
    InChat As Byte
    ChatPlayer As Long
    Mute As Boolean
    Pet As PetRec
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
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    AdvanceFrom As Long
    LevelReq As Long

    Type As Long
    Locked As Long
    MaleSprite As Long
    FemaleSprite As Long
    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long
    Map As Long
    X As Byte
    Y As Byte
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 150
    Pic As Long

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
End Type

Type MapGridRec
    Blocked As Boolean
End Type

Type GridRec
    Loc() As MapGridRec
End Type

Type MapItemRec
    num As Long
    Value As Long
    Dur As Long
    X As Byte
    Y As Byte
End Type

Type NPCEditorRec
    ItemNum As Long
    ItemValue As Long
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
    Speed As Long
    Magi As Long
    Big As Long
    MaxHp As Long
    Exp As Long
    SpawnTime As Long
    Speech As Long
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
End Type

Type MapNpcRec
    num As Long
    TargetType As Byte
    Target As Long
    HP As Long
    MP As Long
    SP As Long
    X As Byte
    Y As Byte
    Dir As Integer

    ' For server use only
    SpawnWait As Long
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

Type TempTileRec

    DoorOpen()  As Byte
    DoorTimer As Long
End Type

Type GuildRec
    Name As String * NAME_LENGTH
    Founder As String * NAME_LENGTH
    Member() As String * NAME_LENGTH
End Type

Type EmoRec
    Pic As Long
    sound As String
    Command As String

    Type As Byte
End Type

Type CMRec
    Title As String
    Message As String
End Type

Type OptionRec
    text As String
    GoTo As Long
    Exit As Byte
End Type

Type InvSpeechRec
    Exit As Byte
    text As String
    SaidBy As Byte
    Respond As Byte
    Script As Long
    Responces(1 To 3) As OptionRec
End Type

Type SpeechRec
    Name As String
    num(0 To MAX_SPEECH_OPTIONS) As InvSpeechRec
End Type

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1
Public NEXT_CHAR As String * 1

' Maximum classes
Public Max_Classes As Byte
Public Map() As MapRec
Public TempTile() As TempTileRec
Public PlayersOnMap() As Long
Public Player() As AccountRec
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public MapItem() As MapItemRec
Public MapNpc() As MapNpcRec
Public Grid() As GridRec
Public Shop() As ShopRec
Public Spell() As SpellRec
Public Guild() As GuildRec
Public Party() As PartyRec
Public Emoticons() As EmoRec
Public Experience() As Long
Public CMessages(1 To 6) As CMRec
Public Speech() As SpeechRec

Type ArrowRec
    Name As String
    Pic As Long
    Range As Byte
End Type

Public Arrows(1 To MAX_ARROWS) As ArrowRec

Type StatRec
    Level As Long
    STR As Long
    DEF As Long
    Magi As Long
    Speed As Long
End Type

Public AddHP As StatRec
Public AddMP As StatRec
Public AddSP As StatRec

Sub BattleMsg(ByVal Index As Long, _
   ByVal Msg As String, _
   ByVal Color As Byte, _
   ByVal Side As Byte)
    Call SendDataTo(Index, DAMAGEDISPLAY_CHAR & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Sub ClearChar(ByVal Index As Long, _
   ByVal CharNum As Long)
    Dim N As Long

    Player(Index).Char(CharNum).Name = vbNullString
    Player(Index).Char(CharNum).Class = 1
    Player(Index).Char(CharNum).Sprite = 0
    Player(Index).Char(CharNum).Level = 0
    Player(Index).Char(CharNum).Exp = 0
    Player(Index).Char(CharNum).Access = 0
    Player(Index).Char(CharNum).PK = NO
    Player(Index).Char(CharNum).POINTS = 0
    Player(Index).Char(CharNum).Guild = vbNullString
    Player(Index).Char(CharNum).HP = 0
    Player(Index).Char(CharNum).MP = 0
    Player(Index).Char(CharNum).SP = 0
    Player(Index).Char(CharNum).STR = 0
    Player(Index).Char(CharNum).DEF = 0
    Player(Index).Char(CharNum).Speed = 0
    Player(Index).Char(CharNum).Magi = 0

    For N = 1 To MAX_INV
        Player(Index).Char(CharNum).Inv(N).num = 0
        Player(Index).Char(CharNum).Inv(N).Value = 0
        Player(Index).Char(CharNum).Inv(N).Dur = 0
    Next

    For N = 1 To MAX_PLAYER_SPELLS
        Player(Index).Char(CharNum).Spell(N) = 0
    Next

    Player(Index).Char(CharNum).ArmorSlot = 0
    Player(Index).Char(CharNum).WeaponSlot = 0
    Player(Index).Char(CharNum).HelmetSlot = 0
    Player(Index).Char(CharNum).ShieldSlot = 0
    Player(Index).Char(CharNum).Map = 0
    Player(Index).Char(CharNum).X = 0
    Player(Index).Char(CharNum).Y = 0
    Player(Index).Char(CharNum).Dir = 0
End Sub

Sub ClearClasses()
    Dim i As Long

    For i = 1 To Max_Classes
        Class(i).Name = vbNullString
        Class(i).AdvanceFrom = 0
        Class(i).LevelReq = 0
        Class(i).Type = 1
        Class(i).STR = 0
        Class(i).DEF = 0
        Class(i).Speed = 0
        Class(i).Magi = 0
        Class(i).FemaleSprite = 0
        Class(i).MaleSprite = 0
        Class(i).Map = 0
        Class(i).X = 0
        Class(i).Y = 0
    Next

End Sub

Sub ClearGrid()
    Dim i As Long, Y As Long, X As Long

    For i = 1 To MAX_MAPS
        For X = 0 To MAX_MAPX
            For Y = 0 To MAX_MAPY
                Grid(i).Loc(X, Y).Blocked = False
            Next
        Next
    Next

End Sub
    
Sub ClearItem(ByVal Index As Long)
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
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
    Item(Index).AttackSpeed = 0
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

Sub ClearMap(ByVal MapNum As Long)
    Dim X As Long
    Dim Y As Long

    Map(MapNum).Name = vbNullString
    Map(MapNum).Revision = 0
    Map(MapNum).Moral = 0
    Map(MapNum).Up = 0
    Map(MapNum).Down = 0
    Map(MapNum).Left = 0
    Map(MapNum).Right = 0
    Map(MapNum).Indoors = 0

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map(MapNum).Tile(X, Y).Ground = 0
            Map(MapNum).Tile(X, Y).Mask = 0
            Map(MapNum).Tile(X, Y).Anim = 0
            Map(MapNum).Tile(X, Y).Mask2 = 0
            Map(MapNum).Tile(X, Y).M2Anim = 0
            Map(MapNum).Tile(X, Y).Fringe = 0
            Map(MapNum).Tile(X, Y).FAnim = 0
            Map(MapNum).Tile(X, Y).Fringe2 = 0
            Map(MapNum).Tile(X, Y).F2Anim = 0
            Map(MapNum).Tile(X, Y).Type = 0
            Map(MapNum).Tile(X, Y).Data1 = 0
            Map(MapNum).Tile(X, Y).Data2 = 0
            Map(MapNum).Tile(X, Y).Data3 = 0
            Map(MapNum).Tile(X, Y).String1 = vbNullString
            Map(MapNum).Tile(X, Y).String2 = vbNullString
            Map(MapNum).Tile(X, Y).String3 = vbNullString
            Map(MapNum).Tile(X, Y).Light = 0
            Map(MapNum).Tile(X, Y).GroundSet = -1
            Map(MapNum).Tile(X, Y).MaskSet = -1
            Map(MapNum).Tile(X, Y).AnimSet = -1
            Map(MapNum).Tile(X, Y).Mask2Set = -1
            Map(MapNum).Tile(X, Y).M2AnimSet = -1
            Map(MapNum).Tile(X, Y).FringeSet = -1
            Map(MapNum).Tile(X, Y).FAnimSet = -1
            Map(MapNum).Tile(X, Y).Fringe2Set = -1
            Map(MapNum).Tile(X, Y).F2AnimSet = -1
        Next
    Next

    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
End Sub

Sub ClearMapItem(ByVal Index As Long, _
   ByVal MapNum As Long)
    MapItem(MapNum, Index).num = 0
    MapItem(MapNum, Index).Value = 0
    MapItem(MapNum, Index).Dur = 0
    MapItem(MapNum, Index).X = 0
    MapItem(MapNum, Index).Y = 0
End Sub

Sub ClearMapItems()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, Y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal Index As Long, _
   ByVal MapNum As Long)
    MapNpc(MapNum, Index).num = 0
    MapNpc(MapNum, Index).TargetType = 0
    MapNpc(MapNum, Index).Target = 0
    MapNpc(MapNum, Index).HP = 0
    MapNpc(MapNum, Index).MP = 0
    MapNpc(MapNum, Index).SP = 0
    MapNpc(MapNum, Index).X = 0
    MapNpc(MapNum, Index).Y = 0
    MapNpc(MapNum, Index).Dir = 0

    ' Server use only
    MapNpc(MapNum, Index).SpawnWait = 0
    MapNpc(MapNum, Index).AttackTimer = 0
    MapNpc(MapNum, Index).LastAttack = 0
End Sub

Sub ClearMapNpcs()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, Y)
        Next
    Next

End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

End Sub

Sub ClearNpc(ByVal Index As Long)
    Dim i As Long

    Npc(Index).Name = vbNullString
    Npc(Index).AttackSay = vbNullString
    Npc(Index).Sprite = 0
    Npc(Index).SpawnSecs = 0
    Npc(Index).Behavior = 0
    Npc(Index).Range = 0
    Npc(Index).STR = 0
    Npc(Index).DEF = 0
    Npc(Index).Speed = 0
    Npc(Index).Magi = 0
    Npc(Index).Big = 0
    Npc(Index).MaxHp = 0
    Npc(Index).Exp = 0
    Npc(Index).SpawnTime = 0
    Npc(Index).Speech = 0

    For i = 1 To MAX_NPC_DROPS
        Npc(Index).ItemNPC(i).Chance = 0
        Npc(Index).ItemNPC(i).ItemNum = 0
        Npc(Index).ItemNPC(i).ItemValue = 0
    Next

End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long
    Dim N As Long

    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString

    For i = 1 To MAX_CHARS
        Player(Index).Char(i).Name = vbNullString
        Player(Index).Char(i).Class = 1
        Player(Index).Char(i).Level = 0
        Player(Index).Char(i).Sprite = 0
        Player(Index).Char(i).Exp = 0
        Player(Index).Char(i).Access = 0
        Player(Index).Char(i).PK = NO
        Player(Index).Char(i).POINTS = 0
        Player(Index).Char(i).Guild = vbNullString
        Player(Index).Char(i).HP = 0
        Player(Index).Char(i).MP = 0
        Player(Index).Char(i).SP = 0
        Player(Index).Char(i).STR = 0
        Player(Index).Char(i).DEF = 0
        Player(Index).Char(i).Speed = 0
        Player(Index).Char(i).Magi = 0

        For N = 1 To MAX_INV
            Player(Index).Char(i).Inv(N).num = 0
            Player(Index).Char(i).Inv(N).Value = 0
            Player(Index).Char(i).Inv(N).Dur = 0
        Next

        For N = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(i).Spell(N) = 0
        Next

        Player(Index).Char(i).ArmorSlot = 0
        Player(Index).Char(i).WeaponSlot = 0
        Player(Index).Char(i).HelmetSlot = 0
        Player(Index).Char(i).ShieldSlot = 0
        Player(Index).Char(i).Map = 0
        Player(Index).Char(i).X = 0
        Player(Index).Char(i).Y = 0
        Player(Index).Char(i).Dir = 0

        For N = 1 To MAX_FRIENDS
            Player(Index).Char(i).Friends(N) = vbNullString
        Next
    Next

    Player(Index).Pet.Alive = NO

    ' Temporary vars
    Player(Index).Buffer = vbNullString
    Player(Index).IncBuffer = vbNullString
    Player(Index).CharNum = 0
    Player(Index).InGame = "NO"
    Player(Index).AttackTimer = 0
    Player(Index).DataTimer = 0
    Player(Index).DataBytes = 0
    Player(Index).DataPackets = 0
    Player(Index).PartyID = 0
    Player(Index).InParty = 0
    Player(Index).Invited = 0
    Player(Index).Target = 0
    Player(Index).TargetType = 0
    Player(Index).CastedSpell = NO
    Player(Index).GettingMap = NO
    Player(Index).Emoticon = -1
    Player(Index).InTrade = 0
    Player(Index).TradePlayer = 0
    Player(Index).TradeOk = 0
    Player(Index).TradeItemMax = 0
    Player(Index).TradeItemMax2 = 0

    For N = 1 To MAX_PLAYER_TRADES
        Player(Index).Trading(N).InvName = vbNullString
        Player(Index).Trading(N).InvNum = 0
    Next

    Player(Index).ChatPlayer = 0
End Sub

Sub ClearShop(ByVal Index As Long)
    Dim i As Long
    Dim z As Long

    Shop(Index).Name = vbNullString
    Shop(Index).JoinSay = vbNullString
    Shop(Index).LeaveSay = vbNullString

    For z = 1 To 6
        For i = 1 To MAX_TRADES
            Shop(Index).TradeItem(z).Value(i).GiveItem = 0
            Shop(Index).TradeItem(z).Value(i).GiveValue = 0
            Shop(Index).TradeItem(z).Value(i).GetItem = 0
            Shop(Index).TradeItem(z).Value(i).GetValue = 0
        Next
    Next

End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

Sub ClearSpeech(ByVal Index As Long)
    Dim i As Long
    Dim o As Long

    Speech(Index).Name = vbNullString

    For o = 0 To MAX_SPEECH_OPTIONS
        Speech(Index).num(o).Exit = 0
        Speech(Index).num(o).Respond = 0
        Speech(Index).num(o).SaidBy = 0
        Speech(Index).num(o).text = "Write what you want to be said here."
        Speech(Index).num(o).Script = 0

        For i = 1 To 3
            Speech(Index).num(o).Responces(i).Exit = 0
            Speech(Index).num(o).Responces(i).GoTo = 0
            Speech(Index).num(o).Responces(i).text = "Write a responce here."
        Next
    Next

End Sub

Sub ClearSpeeches()
    Dim i As Long

    For i = 1 To MAX_SPEECH
        Call ClearSpeech(i)
    Next

End Sub

Sub ClearSpell(ByVal Index As Long)
    Spell(Index).Name = vbNullString
    Spell(Index).ClassReq = 0
    Spell(Index).LevelReq = 0
    Spell(Index).Type = 0
    Spell(Index).Data1 = 0
    Spell(Index).Data2 = 0
    Spell(Index).Data3 = 0
    Spell(Index).MPCost = 0
    Spell(Index).sound = 0
    Spell(Index).Range = 0
    Spell(Index).SpellAnim = 0
    Spell(Index).SpellTime = 40
    Spell(Index).SpellDone = 1
    Spell(Index).AE = 0
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

Sub ClearTempTile()
    Dim i As Long, Y As Long, X As Long

    For i = 1 To MAX_MAPS
        TempTile(i).DoorTimer = 0

        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                TempTile(i).DoorOpen(X, Y) = NO
            Next
        Next
    Next

End Sub

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(Class(ClassNum).STR / 2) + Class(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(Class(ClassNum).Magi / 2) + Class(ClassNum).Magi) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(Class(ClassNum).Speed / 2) + Class(ClassNum).Speed) * 2
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
End Function

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).Char(Player(Index).CharNum).ArmorSlot
End Function

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Function GetPlayerDEF(ByVal Index As Long) As Long
    Dim add As Long

    add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddDef
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddDef
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddDef
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddDef
    End If

    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF + add
End Function
Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Char(Player(Index).CharNum).Guildaccess
End Function

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).Char(Player(Index).CharNum).HelmetSlot
End Function

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
End Function

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num
End Function

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value
End Function

Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).Level
End Function

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Function GetPlayerMAGI(ByVal Index As Long) As Long
    Dim add As Long

    add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddMagi
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddMagi
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddMagi
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddMagi
    End If

    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).Magi + add
End Function

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
End Function

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim add As Long

    add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddHP
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddHP
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddHP
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddHP
    End If

    CharNum = Player(Index).CharNum

    'GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerstr(index) / 2) + Class(Player(index).Char(CharNum).Class).str) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(Index) * AddHP.Level) + (GetPlayerstr(Index) * AddHP.STR) + (GetPlayerDEF(Index) * AddHP.DEF) + (GetPlayerMAGI(Index) * AddHP.Magi) + (GetPlayerSPEED(Index) * AddHP.Speed) + add
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim add As Long

    add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddMP
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddMP
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddMP
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddMP
    End If

    CharNum = Player(Index).CharNum

    'GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + Class(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(Index) * AddMP.Level) + (GetPlayerstr(Index) * AddMP.STR) + (GetPlayerDEF(Index) * AddMP.DEF) + (GetPlayerMAGI(Index) * AddMP.Magi) + (GetPlayerSPEED(Index) * AddMP.Speed) + add
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim add As Long

    add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddSP
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddSP
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddSP
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddSP
    End If

    CharNum = Player(Index).CharNum

    'GetPlayerMaxSP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSPEED(index) / 2) + Class(Player(index).Char(CharNum).Class).SPEED) * 2) + add
    GetPlayerMaxSP = (GetPlayerLevel(Index) * AddSP.Level) + (GetPlayerstr(Index) * AddSP.STR) + (GetPlayerDEF(Index) * AddSP.DEF) + (GetPlayerMAGI(Index) * AddSP.Magi) + (GetPlayerSPEED(Index) * AddSP.Speed) + add
End Function

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
End Function

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
End Function

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
End Function

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
End Function

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).Char(Player(Index).CharNum).ShieldSlot
End Function

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Function GetPlayerSPEED(ByVal Index As Long) As Long
    Dim add As Long

    add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddSpeed
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddSpeed
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddSpeed
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddSpeed
    End If

    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).Speed + add
End Function

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot)
End Function

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).Sprite
End Function

Function GetPlayerstr(ByVal Index As Long) As Long
    Dim add As Long

    add = 0

    If GetPlayerWeaponSlot(Index) > 0 Then
        add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddStr
    End If

    If GetPlayerArmorSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddStr
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddStr
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddStr
    End If

    GetPlayerstr = Player(Index).Char(Player(Index).CharNum).STR + add
End Function

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).Char(Player(Index).CharNum).WeaponSlot
End Function

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).X
End Function

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).Y
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(Player(Index).CharNum).Access = Access
End Sub

Sub SetPlayerArmorSlot(ByVal Index As Long, _
   InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ArmorSlot = InvNum
End Sub

Sub SetPlayerClass(ByVal Index As Long, _
   ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

Sub SetPlayerDEF(ByVal Index As Long, _
   ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub
Sub SetPlayerDir(ByVal Index As Long, _
   ByVal Dir As Long)
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
End Sub

Sub SetPlayerExp(ByVal Index As Long, _
   ByVal Exp As Long)
    Player(Index).Char(Player(Index).CharNum).Exp = Exp
End Sub

Sub SetPlayerGuild(ByVal Index As Long, _
   ByVal Guild As String)
    Player(Index).Char(Player(Index).CharNum).Guild = Guild
End Sub

Sub SetPlayerGuildAccess(ByVal Index As Long, _
   ByVal Guildaccess As Long)
    Player(Index).Char(Player(Index).CharNum).Guildaccess = Guildaccess
End Sub

Sub SetPlayerHelmetSlot(ByVal Index As Long, _
   InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).HelmetSlot = InvNum
End Sub

Sub SetPlayerHP(ByVal Index As Long, _
   ByVal HP As Long)
    Player(Index).Char(Player(Index).CharNum).HP = HP

    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    End If

    If GetPlayerHP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).HP = 0
    End If

    Call SendStats(Index)
End Sub

Sub SetPlayerInvItemDur(ByVal Index As Long, _
   ByVal InvSlot As Long, _
   ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Sub SetPlayerInvItemNum(ByVal Index As Long, _
   ByVal InvSlot As Long, _
   ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num = ItemNum
End Sub

Sub SetPlayerInvItemValue(ByVal Index As Long, _
   ByVal InvSlot As Long, _
   ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Sub SetPlayerLevel(ByVal Index As Long, _
   ByVal Level As Long)
    Player(Index).Char(Player(Index).CharNum).Level = Level
End Sub

Sub SetPlayerMAGI(ByVal Index As Long, _
   ByVal Magi As Long)
    Player(Index).Char(Player(Index).CharNum).Magi = Magi
End Sub

Sub SetPlayerMap(ByVal Index As Long, _
   ByVal MapNum As Long)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(Player(Index).CharNum).Map = MapNum
    End If

End Sub

Sub SetPlayerMP(ByVal Index As Long, _
   ByVal MP As Long)
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    End If

    If GetPlayerMP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).MP = 0
    End If

End Sub

Sub SetPlayerPK(ByVal Index As Long, _
   ByVal PK As Long)
    Player(Index).Char(Player(Index).CharNum).PK = PK
End Sub

Sub SetPlayerPOINTS(ByVal Index As Long, _
   ByVal POINTS As Long)
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
End Sub

Sub SetPlayerShieldSlot(ByVal Index As Long, _
   InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ShieldSlot = InvNum
End Sub

Sub SetPlayerSP(ByVal Index As Long, _
   ByVal SP As Long)
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    End If

    If GetPlayerSP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).SP = 0
    End If

End Sub

Sub SetPlayerSPEED(ByVal Index As Long, _
   ByVal Speed As Long)
    Player(Index).Char(Player(Index).CharNum).Speed = Speed
End Sub

Sub SetPlayerSpell(ByVal Index As Long, _
   ByVal SpellSlot As Long, _
   ByVal SpellNum As Long)
    Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Sub SetPlayerSprite(ByVal Index As Long, _
   ByVal Sprite As Long)
    Player(Index).Char(Player(Index).CharNum).Sprite = Sprite
End Sub

Sub SetPlayerstr(ByVal Index As Long, _
   ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Sub SetPlayerWeaponSlot(ByVal Index As Long, _
   InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).WeaponSlot = InvNum
End Sub

Sub SetPlayerX(ByVal Index As Long, _
   ByVal X As Long)
    Player(Index).Char(Player(Index).CharNum).X = X
End Sub

Sub SetPlayerY(ByVal Index As Long, _
   ByVal Y As Long)
    Player(Index).Char(Player(Index).CharNum).Y = Y
End Sub
