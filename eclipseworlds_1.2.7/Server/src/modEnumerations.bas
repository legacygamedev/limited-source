Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the client's Packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SLoginOk
    SNewCharClasses
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHP
    SPlayerMP '
    SPlayerStats
    SPlayerSkills
    SPlayerPoints
    SPlayerLevel
    SPlayerGuild
    SPlayerSprite
    SPlayerTitles
    SPlayerStatus
    SPlayerPK
    SPlayerData
    SPlayerMove
    SPlayerWarp
    SNPCMove
    SPlayerDir
    SNPCDir
    SAttack
    SNPCAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNPCData
    SMapNPCTarget
    SMapDone
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SREditor
    SSpawnNPC
    SNPCDead
    SNPCEditor
    SUpdateNPC
    SEditMap
    SEditEvent
    SShopEditor
    SUpdateShop
    SSpellEditor
    SUpdateSpell
    SSpells
    SSpell
    SLeft
    SResourceCache
    SResourceEditor
    SUpdateResource
    SSendPing
    SDoorAnimation
    SActionMsg
    SPlayerEXP
    SBlood
    SAnimationEditor
    SUpdateAnimation
    SAnimation
    SMapNPCVitals
    SSpellCooldown
    SClearSpellBuffer
    SSayMsg
    SOpenShop
    SResetShopAction
    SStunned
    SMapWornEq
    sbank
    STrade
    SCloseTrade
    STradeUpdate
    STradeStatus
    STradeRequest
    SPartyInvite
    SPartyUpdate
    SPartyVitals
    STarget
    SHotbar
    SGuildMembers
    SGuildInvite
    SMapReport
    SNPCSpellBuffer
    SCheckpoint
    SUpdateLogs
    SFriendsList
    SFoesList
    SHighIndex
    SEntitySound
    SGameData
    SSendNews
    SSound
    SBanEditor
    SUpdateBan
    STitleEditor
    SUpdateTitle
    SMoralEditor
    SUpdateMoral
    SClassEditor
    SUpdateClass
    SCloseClient
    SLeaveGame
    SEmoticonEditor
    SUpdateEmoticon
    SCheckEmoticon
    SSpawnEvent
    SEventMove
    SEventDir
    SEventChat
    SEventStart
    SEventEnd
    SPlayBGM
    SPlaySound
    SFadeoutBGM
    SStopSound
    SSwitchesAndVariables
    SMapEventData
    SChatBubble
    SSpecialEffect
    
    ' Character editor
    SPlayersOnline
    SAllCharacters
    SExtendedPlayerData
    
    SAccessVerificator
    
    SEditQuest
    SUpdateQuest
    SQuestRequest
    SQuestMsg
    
    SRefreshCharEditor
    
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CDelAccount
    CLogin
    CAddChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CGlobalMsg
    CAdminMsg
    CPrivateMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CSetPlayerSprite
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    COpenMaps
    CKickPlayer
    CMutePlayer
    CBanPlayer
    CRequestPlayerData
    CRequestPlayerStats
    CRequestSpellCooldown
    CRequestEditMap
    CRequestEditEvent
    CRequestEditItem
    CSaveItem
    CRequestEditNPC
    CSaveNPC
    CRequestEditShop
    CSaveShop
    CRequestEditSpell
    CSaveSpell
    CSetAccess
    CWhosOnline
    CSetMOTD
    CSetSMotd
    CSetGMotd
    CSearch
    CSpells
    CCastSpell
    CSwapInvSlots
    CSwapSpellSlots
    CSwapHotbarSlots
    CRequestEditResource
    CSaveResource
    CCheckPing
    CUnequip
    CRequestItems
    CRequestNPCs
    CRequestResources
    CSpawnItem
    CUseStatPoint
    CRequestEditAnimation
    CSaveAnimation
    CRequestAnimations
    CRequestSpells
    CRequestShops
    CRequestLevelUp
    CForgetSpell
    CCloseShop
    CBuyItem
    CSellItem
    CSwapBankSlots
    CDepositItem
    CWithdrawItem
    CCloseBank
    CAdminWarp
    CFixItem
    CTradeRequest
    CAcceptTradeRequest
    CDeclineTradeRequest
    CCanTrade
    CTradeItem
    CUntradeItem
    CHotbarChange
    CPartyMsg
    CGuildCreate
    CGuildChangeAccess
    CGuildInvite
    CAcceptGuild
    CDeclineGuild
    CAcceptTrade
    CDeclineTrade
    CGuildRemove
    CGuildDisband
    CGuildMsg
    CBreakSpell
    CFriendsList
    CAddFriend
    CRemoveFriend
    CFoesList
    CAddFoe
    CRemoveFoe
    CAcceptParty
    CDeclineParty
    CPartyRequest
    CPartyLeave
    CUpdateData
    CSaveBan
    CRequestEditBans
    CRequestBans
    CSetTitle
    CRequestEditTitles
    CSaveTitle
    CRequestTitles
    CChangeStatus
    CRequestEditMorals
    CSaveMoral
    CRequestMorals
    CRequestEditClasses
    CSaveClass
    CRequestClasses
    CDestoryItem
    CRequestEditEmoticons
    CSaveEmoticon
    CRequestEmoticons
    CCheckEmoticon
    CEventChatReply
    CEvent
    CSwitchesAndVariables
    CRequestSwitchesAndVariables
    
    ' Character editor
    CRequestAllCharacters
    CRequestPlayersOnline
    CRequestExtendedPlayerData
    CCharacterUpdate
    
    CTarget
    
    CRequestEditQuests
    CSaveQuest
    CQuitQuest
    CAcceptQuest
    CRequestQuests
    
    CChangeDataSize
    
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(CMSG_COUNT) As Long

' Stats used by Players, NPCs and Classes
Public Enum Stats
    Strength = 1
    Endurance
    Intelligence
    Agility
    Spirit
    
    ' Make sure Stat_Count is below everything else
    Stat_count
End Enum

' Vitals used by Players, NPCs and Classes
Public Enum Vitals
    HP = 1
    MP
    
    ' Make sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Body
    Head
    Shield
    Hands
    Feet
    Ring
    Neck
    
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum

' Layers in a map
Public Enum MapLayer
    Ground = 1
    Mask
    Cover
    Fringe
    Roof
    
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum

' Sound entities
Public Enum SoundEntity
    seAnimation = 1
    seItem
    seNPC
    seResource
    seSpell
    
    ' Make sure SoundEntity_Count is below everything else
    SoundEntity_Count
End Enum

Public Enum Skills
    Herbalism = 1
    Alchemy
    Farming
    Crafting
    Woodcutting
    Fletching
    Firemaking
    Mining
    Smithing
    Fishing
    Cooking
    Prayer
    
    ' Make sure Skill_Count is below everything else
    Skill_Count
End Enum

Public Enum Proficiency
    Light = 1
    Medium
    Heavy
    Sword
    Dagger
    Bow
    Crossbow
    Mace
    Axe
    Spear
    Staff
    
    ' Make sure Proficiency_Count is below everything else
    Proficiency_Count
End Enum

' Event Types
Public Enum EventType
    ' Message
    evAddText = 1
    evShowText
    evShowChoices
    ' Game Progression
    evPlayerVar
    evPlayerSwitch
    evSelfSwitch
    ' Flow Control
    evCondition
    evExitProcess
    ' Player
    evChangeItems
    evRestoreHP
    evRestoreMP
    evLevelUp
    evChangeLevel
    evChangeSkills
    evChangeClass
    evChangeSprite
    evChangeGender
    evChangePK
    ' Movement
    evWarpPlayer
    evSetMoveRoute
    ' Character
    evPlayAnimation
    ' Music and Sounds
    evPlayBGM
    evFadeoutBGM
    evPlaySound
    evStopSound
    ' Etc...
    evCustomScript
    evSetAccess
    ' Shop/Bank
    evOpenBank
    evOpenShop
    'New
    evGiveExp
    evShowChatBubble
    evLabel
    evGotoLabel
    evSpawnNPC
    evFadeIn
    evFadeOut
    evFlashWhite
    evSetFog
    evSetweather
    evSetTint
    evWait
    evAddTitle
    evRemoveTitle
End Enum


