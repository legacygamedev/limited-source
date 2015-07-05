Attribute VB_Name = "modEnumerations"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

' *************
' ** Packets **
' *************

' The order of the packets must match with the server's packet enumeration

' Packets sent by the client
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
    CMSG_COUNT
End Enum

' Packets recieved by the client
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
    SMSG_COUNT
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
    SPEED
    Magic
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    SP
    ' Mak sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    Helmet
    Shield
    ' Mak sure Equipment_Count is below everything else
    Equipment_Count
End Enum


