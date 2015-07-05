Attribute VB_Name = "modMessages"
Option Explicit
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/10/2005  Verrigan   Module created to enumerate the various
'*                        server messages. You can use this file
'*                        to store client messages, too.
'****************************************************************
Public Enum SMsgTypes
  SMsgGetClasses = 0
  SMsgNewAccount
  SMsgDelAccount
  SMsgLogin
  SMsgAddChar
  SMsgDelChar
  SMsgUseChar
  SMsgSay
  SMsgEmote
  SMsgBroadcast
  SMsgGlobal
  SMsgAdmin
  SMsgPlayer
  SMsgPlayerMove
  SMsgPlayerDir
  SMsgUseItem
  SMsgAttack
  SMsgUseStatPoint
  SMsgPlayerInfoRequest
  SMsgWarpMeTo
  SMsgWarpToMe
  SMsgWarpTo
  SMsgSetSprite
  SMsgGetStats
  SMsgRequestNewMap
  SMsgMapData
  SMsgNeedMap
  SMsgMapGetItem
  SMsgMapDropItem
  SMsgMapRespawn
  SMsgMapReport
  SMsgKickPlayer
  SMsgBanList
  SMsgBanDestroy
  SMsgBanPlayer
  SMsgRequestEditMap
  SMsgRequestEditItem
  SMsgEditItem
  SMsgSaveItem
  SMsgRequestEditNPC
  SMsgEditNPC
  SMsgSaveNPC
  SMsgRequestEditShop
  SMsgEditShop
  SMsgSaveShop
  SMsgRequestEditSpell
  SMsgEditSpell
  SMsgSaveSpell
  SMsgSetAccess
  SMsgWhosOnline
  SMsgSetMOTD
  SMsgTrade
  SMsgTradeRequest
  SMsgFixItem
  SMsgSearch
  SMsgParty
  SMsgJoinParty
  SMsgLeaveParty
  SMsgSpells
  SMsgCast
  SMsgRequestLocation
  'The following enum member automatically stores the number of messages,
  'since it is last. Any new messages must be placed above this entry.
  SMSG_COUNT
End Enum
Public HandleDataSub(SMSG_COUNT) As Long
