Attribute VB_Name = "Enumerations"
Option Explicit

' ******************************************
' **               rootSource               **
' ******************************************

' *************
' ** Packets **
' *************

' The order of the packets must match with the server's packet enumeration

' Packets sent by the server
Public Enum ServerPackets
    SAlertMsg = 1
    SAllChars
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    SPlayerSp
    SPlayerStats
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SSayMsg
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SEditItem
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SEditNpc
    SMapKey
    SEditMap
    SShopEditor
    SUpdateShop
    SEditShop
    SSpellEditor
    SUpdateSpell
    SEditSpell
    STrade
    SSpells
    SLeft
    SHighIndex
    SCastSpell
    SDoor
    SSendMaxes
    SSync
    SMapRevs
    'The following enum member automatically stores the number of messages,
    'since it is last. Any new messages must be placed above this entry.
    SMSG_COUNT
End Enum

' Packets recieved by the server
Public Enum ClientPackets
    CGetClasses = 1
    CNewAccount
    CDelAccount
    CLogin
    CAddChar
    CDelChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CGlobalMsg
    CAdminMsg
    CPlayerMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CPlayerInfoRequest
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CGetStats
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CEditItem
    CSaveItem
    CRequestEditNpc
    CEditNpc
    CSaveNpc
    CRequestEditShop
    CEditShop
    CSaveShop
    CRequestEditSpell
    CEditSpell
    CSaveSpell
    CDelete
    CSetAccess
    CWhosOnline
    CSetMotd
    CTrade
    CTradeRequest
    CFixItem
    CSearch
    CParty
    CJoinParty
    CLeaveParty
    CSpells
    CCast
    CQuit
    CSync
    CMapReqs
    CSleepInn
    CRemoveFromGuild
    CCreateGuild
    CInviteGuild
    CKickGuild
    CGuildPromote
    CLeaveGuild
    'The following enum member automatically stores the number of messages,
    'since it is last. Any new messages must be placed above this entry.
    CMSG_COUNT
End Enum

' Holds the memory address of the packet subs
Public HandleDataSub(SMSG_COUNT) As Long

' ****************
' ** Statistics **
' ****************

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Defense
    Speed
    Magic
    Stat_Count ' This must be at the end
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    SP
    Vital_Count ' This must be at the end
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    Helmet
    Shield
    Equipment_Count ' This must be at the end
End Enum


