Attribute VB_Name = "modTypes"

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

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
Public MAX_SPEECH As Long

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

' GetVar values
Public MusicOn As Byte
Public SoundOn As Byte
Public MapGridOn As Byte
Public NPCDamageOn As Byte
Public PlayerNameOn As Byte
Public PlayerDamageOn As Byte
Public NPCNameOn As Byte
Public SpeechBubblesOn As Byte
Public EmoticonSoundOn As Byte

' Controlling volume
Public MusicVolume As Byte

Public Const MAX_ARROWS As Byte = 100
Public Const MAX_PLAYER_ARROWS As Byte = 100

Public Const MAX_INV As Byte = 24
Public Const MAX_MAP_NPCS As Byte = 15
Public Const MAX_PLAYER_SPELLS As Byte = 20
Public Const MAX_TRADES As Byte = 66
Public Const MAX_PLAYER_TRADES As Byte = 8
Public Const MAX_NPC_DROPS As Byte = 10
Public Const MAX_SPEECH_OPTIONS As Byte = 20
Public Const MAX_FRIENDS As Byte = 20

Public Const NO As Byte = 0
Public Const YES As Byte = 1

' Version constants
Public Const CLIENT_MAJOR As Byte = 3
Public Const CLIENT_MINOR As Byte = 2
Public Const CLIENT_REVISION As Byte = 0

' Account constants
Public Const NAME_LENGTH As Byte = 20
Public Const MAX_CHARS As Byte = 3

' Security password
Public Const SEC_CODE As String = "pingumadethisgiakenedited"

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

' Size constants (of player sprites)
Public Const SIZE_X As Byte = 32
Public Const SIZE_Y As Byte = 32

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

' Speach bubble constants
Public Const DISPLAY_BUBBLE_TIME As Long = 2000 ' In milliseconds.
Public DISPLAY_BUBBLE_WIDTH As Byte
Public Const MAX_BUBBLE_WIDTH As Byte = 6 ' In tiles. Includes corners.
Public Const MAX_LINE_LENGTH As Byte = 23 ' In characters.
Public Const MAX_LINES As Byte = 3

' Spell constants
Public Const SPELL_TYPE_ADDHP As Byte = 0
Public Const SPELL_TYPE_ADDMP As Byte = 1
Public Const SPELL_TYPE_ADDSP As Byte = 2
Public Const SPELL_TYPE_SUBHP As Byte = 3
Public Const SPELL_TYPE_SUBMP As Byte = 4
Public Const SPELL_TYPE_SUBSP As Byte = 5
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

Type ChatBubble
    Text As String
    Created As Long
End Type

Type PlayerInvRec
    Num As Long
    Value As Long
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

Type MapArrowRec
    Arrow As Byte
    ArrowNum As Long
    ArrowAnim As Long
    ArrowTime As Long
    ArrowVarX As Long
    ArrowVarY As Long
    ArrowX As Long
    ArrowY As Long
    ArrowPosition As Byte
    ArrowOwner As Long
    ArrowMap As Long
End Type

Type PetRec
    Sprite As Long
    
    Alive As Byte
    
    HP As Long
    MaxHP As Long
    
    Map As Long
    x As Long
    y As Long
    Dir As Byte
    
    Moving As Byte
    XOffset As Long
    YOffset As Long
    
    AttackTimer As Long
    Attacking As Byte
    
    LastAttack As Long
End Type

Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
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
    Speed As Long
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
    
    ' Pet!
    Pet As PetRec
    
    ' Client use only
    MaxHP As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    LastAttack As Long
    MapGetTimer As Long
    CastedSpell As Byte
    
    SpellNum As Long
    SpellAnim() As SpellAnimRec

    EmoticonNum As Long
    EmoticonSound As String
    EmoticonType As Long
    EmoticonTime As Long
    EmoticonVar As Long
    EmoticonPlayed As Boolean
    
    LevelUp As Long
    LevelUpT As Long

    'Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec
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
    x As Long
    y As Long
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
    Arrow(1 To MAX_PLAYER_ARROWS) As MapArrowRec
End Type

Type ClassRec
    Name As String * NAME_LENGTH
    MaleSprite As Long
    FemaleSprite As Long
    
    Locked As Long
    
    STR As Long
    DEF As Long
    Speed As Long
    MAGI As Long
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    desc As String * 150
    
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

Type MapItemRec
    Num As Long
    Value As Long
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
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    
    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
        
    STR  As Long
    DEF As Long
    Speed As Long
    MAGI As Long
    Big As Long
    MaxHP As Long
    EXP As Long
    SpawnTime As Long
    
    Speech As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
End Type

Type MapNpcRec
    Num As Long
    
    Target As Long
    
    HP As Long
    MaxHP As Long
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

Type PlayerTradeRec
    InvNum As Long
    InvName As String
End Type
Public Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec
Public Trading2(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

Type EmoRec
    Pic As Long
    Sound As String
    Command As String
    Type As Byte
End Type

Type OptionRec
    Text As String
    GoTo As Long
    Exit As Byte
End Type

Type InvSpeechRec
    Exit As Byte
    Text As String
    SaidBy As Byte
    Respond As Byte
    Script As Long
    Responces(1 To 3) As OptionRec
End Type

Type SpeechRec
    Name As String
    Num(0 To MAX_SPEECH_OPTIONS) As InvSpeechRec
End Type

Type DropRainRec
    x As Long
    y As Long
    Randomized As Boolean
    Speed As Byte
End Type

' Bubble thing
Public Bubble() As ChatBubble

' Used for parsing
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1
Public NEXT_CHAR As String * 1

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
Public Speech() As SpeechRec

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
    Name As String
    Pic As Long
    Range As Byte
End Type
Public Arrows(1 To MAX_ARROWS) As ArrowRec

Type BattleMsgRec
    Msg As String
    Index As Byte
    Color As Byte
    Time As Long
    done As Byte
    y As Long
End Type
Public BattlePMsg() As BattleMsgRec
Public BattleMMsg() As BattleMsgRec

Type ItemDurRec
    Item As Long
    Dur As Long
    done As Byte
End Type
Public ItemDur(1 To 4) As ItemDurRec

Public TempNpcSpawn(1 To MAX_MAP_NPCS) As LocRec

Public Inventory As Long

Sub ClearTempTile()
Dim x As Long, y As Long

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            TempTile(x, y).DoorOpen = NO
            
            TempTile(x, y).Ground = 0
            TempTile(x, y).Mask = 0
            TempTile(x, y).Anim = 0
            TempTile(x, y).Mask2 = 0
            TempTile(x, y).M2Anim = 0
            TempTile(x, y).Fringe = 0
            TempTile(x, y).FAnim = 0
            TempTile(x, y).Fringe2 = 0
            TempTile(x, y).F2Anim = 0
            TempTile(x, y).Type = TILE_TYPE_NONE
            TempTile(x, y).Data1 = 0
            TempTile(x, y).Data2 = 0
            TempTile(x, y).Data3 = 0
            TempTile(x, y).String1 = vbNullString
            TempTile(x, y).String2 = vbNullString
            TempTile(x, y).String3 = vbNullString
            TempTile(x, y).Light = 0
            TempTile(x, y).GroundSet = 0
            TempTile(x, y).MaskSet = 0
            TempTile(x, y).AnimSet = 0
            TempTile(x, y).Mask2Set = 0
            TempTile(x, y).M2AnimSet = 0
            TempTile(x, y).FringeSet = 0
            TempTile(x, y).FAnimSet = 0
            TempTile(x, y).Fringe2Set = 0
            TempTile(x, y).F2AnimSet = 0
        Next x
    Next y
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim I As Long
Dim n As Long

    Player(Index).Name = vbNullString
    Player(Index).Guild = vbNullString
    Player(Index).Guildaccess = 0
    Player(Index).Class = 1
    Player(Index).Level = 0
    Player(Index).Sprite = 0
    Player(Index).EXP = 0
    Player(Index).Access = 0
    Player(Index).PK = NO
        
    Player(Index).HP = 0
    Player(Index).MP = 0
    Player(Index).SP = 0
        
    Player(Index).STR = 0
    Player(Index).DEF = 0
    Player(Index).Speed = 0
    Player(Index).MAGI = 0
        
    For n = 1 To MAX_INV
        Player(Index).Inv(n).Num = 0
        Player(Index).Inv(n).Value = 0
        Player(Index).Inv(n).Dur = 0
    Next n
        
    Player(Index).ArmorSlot = 0
    Player(Index).WeaponSlot = 0
    Player(Index).HelmetSlot = 0
    Player(Index).ShieldSlot = 0
        
    Player(Index).Map = 0
    Player(Index).x = 0
    Player(Index).y = 0
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
    Player(Index).EmoticonNum = -1
    Player(Index).EmoticonSound = vbNullString
    Player(Index).EmoticonType = 0
    Player(Index).EmoticonTime = 0
    Player(Index).EmoticonVar = 0
    Player(Index).EmoticonPlayed = True
    
    For I = 1 To MAX_SPELL_ANIM
        Player(Index).SpellAnim(I).CastedSpell = NO
        Player(Index).SpellAnim(I).SpellTime = 0
        Player(Index).SpellAnim(I).SpellVar = 0
        Player(Index).SpellAnim(I).SpellDone = 0
        
        Player(Index).SpellAnim(I).Target = 0
        Player(Index).SpellAnim(I).TargetType = TARGET_TYPE_PLAYER
    Next I
    
    Player(Index).SpellNum = 0
    
    For I = 1 To MAX_BLT_LINE
        BattlePMsg(I).Index = 1
        BattlePMsg(I).Time = I
        BattleMMsg(I).Index = 1
        BattleMMsg(I).Time = I
    Next I
    
    Inventory = 1
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index).Name = vbNullString
    Item(Index).desc = vbNullString
    
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
    Item(Index).AttackSpeed = 1000
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
    MapItem(Index).x = 0
    MapItem(Index).y = 0
End Sub

Sub ClearMaps()
Dim I As Long

For I = 1 To MAX_MAPS
    Call ClearMap(I)
Next I
End Sub

Sub ClearMap(ByVal MapNum As Long)
Dim I, x, y As Long

    I = MapNum
    Map(I).Name = vbNullString
    Map(I).Revision = 0
    Map(I).Moral = 0
    Map(I).Up = 0
    Map(I).Down = 0
    Map(I).Left = 0
    Map(I).Right = 0
    Map(I).Indoors = 0
        
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Map(I).Tile(x, y).Ground = 0
            Map(I).Tile(x, y).Mask = 0
            Map(I).Tile(x, y).Anim = 0
            Map(I).Tile(x, y).Mask2 = 0
            Map(I).Tile(x, y).M2Anim = 0
            Map(I).Tile(x, y).Fringe = 0
            Map(I).Tile(x, y).FAnim = 0
            Map(I).Tile(x, y).Fringe2 = 0
            Map(I).Tile(x, y).F2Anim = 0
            Map(I).Tile(x, y).Type = 0
            Map(I).Tile(x, y).Data1 = 0
            Map(I).Tile(x, y).Data2 = 0
            Map(I).Tile(x, y).Data3 = 0
            Map(I).Tile(x, y).String1 = vbNullString
            Map(I).Tile(x, y).String2 = vbNullString
            Map(I).Tile(x, y).String3 = vbNullString
            Map(I).Tile(x, y).Light = 0
            Map(I).Tile(x, y).GroundSet = -1
            Map(I).Tile(x, y).MaskSet = -1
            Map(I).Tile(x, y).AnimSet = -1
            Map(I).Tile(x, y).Mask2Set = -1
            Map(I).Tile(x, y).M2AnimSet = -1
            Map(I).Tile(x, y).FringeSet = -1
            Map(I).Tile(x, y).FAnimSet = -1
            Map(I).Tile(x, y).Fringe2Set = -1
            Map(I).Tile(x, y).F2AnimSet = -1
        Next x
    Next y
End Sub

Sub ClearMapItems()
Dim x As Long

    For x = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(x)
    Next x
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    MapNpc(Index).Num = 0
    MapNpc(Index).Target = 0
    MapNpc(Index).HP = 0
    MapNpc(Index).MP = 0
    MapNpc(Index).SP = 0
    MapNpc(Index).Map = 0
    MapNpc(Index).x = 0
    MapNpc(Index).y = 0
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

Sub ClearSpeech(ByVal Index As Long)
Dim I As Long
Dim O As Long

    Speech(Index).Name = vbNullString

    For O = 0 To MAX_SPEECH_OPTIONS
        Speech(Index).Num(O).Exit = 0
        Speech(Index).Num(O).Respond = 0
        Speech(Index).Num(O).SaidBy = 0
        Speech(Index).Num(O).Text = "Write what you want to be said here."
        Speech(Index).Num(O).Script = 0
    
        For I = 1 To 3
            Speech(Index).Num(O).Responces(I).Exit = 0
            Speech(Index).Num(O).Responces(I).GoTo = 0
            Speech(Index).Num(O).Responces(I).Text = "Write a responce here."
        Next I
    Next O
End Sub

Sub ClearSpeeches()
Dim I As Long

    For I = 1 To MAX_SPEECH
        Call ClearSpeech(I)
    Next I
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Guildaccess As Long)
    Player(Index).Guildaccess = Guildaccess
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
    GetPlayerExp = Player(Index).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    Player(Index).EXP = EXP
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
    GetPlayerSPEED = Player(Index).Speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Speed = Speed
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
If Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).y = y
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

