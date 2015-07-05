Attribute VB_Name = "modGeneral"

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.
Option Explicit

Public Declare Function GetTickCount _
   Lib "kernel32" () As Long

' Version constants
Public Const CLIENT_MAJOR As Byte = 3
Public Const CLIENT_MINOR As Byte = 2
Public Const CLIENT_REVISION As Byte = 0

' Security password
Public Const SEC_CODE As String = "pingumadethisgiakenedited"

' Used for respawning items
Public SpawnSeconds As Long

' Used for weather effects
Public GameWeather As Long
Public WeatherSeconds As Long
Public GameTime As Long
Public TimeSeconds As Long
Public RainIntensity As Long

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back players and npcs hp
Public GiveHPTimer As Long
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean
Public CurrentLoad As Long

Sub Main()
On Error GoTo ErrHandler

    Dim i As Long
    Dim f As Long
    Dim stringy As String
    
    ServerOn = NO
    
    ' Packets!

    MAXINFO_CHAR = Chr$(0)
    INFO_CHAR = Chr$(1)
    NPCHP_CHAR = Chr$(2)
    ALERTMSG_CHAR = Chr$(3)
    PLAINMSG_CHAR = Chr$(4)
    ALLCHARS_CHAR = Chr$(5)
    LOGINOK_CHAR = Chr$(6)
    NEWCHARCLASSES_CHAR = Chr$(7)
    CLASSESDATA_CHAR = Chr$(8)
    INGAME_CHAR = Chr$(9)
    PLAYERINV_CHAR = Chr$(10)
    PLAYERINVUPDATE_CHAR = Chr$(11)
    PLAYERWORNEQ_CHAR = Chr$(12)
    PLAYERPOINTS_CHAR = Chr$(13)
    PLAYERHP_CHAR = Chr$(14)
    PETHP_CHAR = Chr$(15)
    PLAYERMP_CHAR = Chr$(16)
    MAPMSG2_CHAR = Chr$(17)
    PLAYERSP_CHAR = Chr$(18)
    PLAYERSTATSPACKET_CHAR = Chr$(19)
    PLAYERDATA_CHAR = Chr$(20)
    PETDATA_CHAR = Chr$(21)
    PLAYERMOVE_CHAR = Chr$(22)
    PETMOVE_CHAR = Chr$(23)
    NPCMOVE_CHAR = Chr$(24)
    PLAYERDIR_CHAR = Chr$(25)
    NPCDIR_CHAR = Chr$(26)
    PLAYERXY_CHAR = Chr$(27)
    ATTACKPLAYER_CHAR = Chr$(28)
    ATTACKNPC_CHAR = Chr$(29)
    PETATTACKNPC_CHAR = Chr$(30)
    NPCATTACK_CHAR = Chr$(31)
    NPCATTACKPET_CHAR = Chr$(32)
    CHECKFORMAP_CHAR = Chr$(33)
    MAPDATA_CHAR = Chr$(34)
    MAPITEMDATA_CHAR = Chr$(35)
    MAPNPCDATA_CHAR = Chr$(36)
    MAPDONE_CHAR = Chr$(37)
    SAYMSG_CHAR = Chr$(38)
    SPAWNITEM_CHAR = Chr$(39)
    ITEMEDITOR_CHAR = Chr$(40)
    UPDATEITEM_CHAR = Chr$(41)
    EDITITEM_CHAR = Chr$(42)
    SPAWNNPC_CHAR = Chr$(43)
    NPCDEAD_CHAR = Chr$(44)
    NPCEDITOR_CHAR = Chr$(45)
    UPDATENPC_CHAR = Chr$(46)
    EDITNPC_CHAR = Chr$(47)
    MAPKEY_CHAR = Chr$(48)
    EDITMAP_CHAR = Chr$(49)
    SHOPEDITOR_CHAR = Chr$(50)
    UPDATESHOP_CHAR = Chr$(51)
    EDITSHOP_CHAR = Chr$(52)
    MAINEDITOR_CHAR = Chr$(53)
    SPELLEDITOR_CHAR = Chr$(54)
    UPDATESPELL_CHAR = Chr$(55)
    EDITSPELL_CHAR = Chr$(56)
    TRADE_CHAR = Chr$(57)
    STARTSPEECH_CHAR = Chr$(58)
    SPELLS_CHAR = Chr$(59)
    WEATHER_CHAR = Chr$(60)
    TIME_CHAR = Chr$(61)
    ONLINELIST_CHAR = Chr$(62)
    BLITPLAYERDMG_CHAR = Chr$(63)
    BLITNPCDMG_CHAR = Chr$(64)
    PPTRADING_CHAR = Chr$(65)
    QTRADE_CHAR = Chr$(66)
    UPDATETRADEITEM_CHAR = Chr$(67)
    TRADING_CHAR = Chr$(68)
    PPCHATING_CHAR = Chr$(69)
    QCHAT_CHAR = Chr$(70)
    SENDCHAT_CHAR = Chr$(71)
    SOUND_CHAR = Chr$(72)
    SPRITECHANGE_CHAR = Chr$(73)
    CHANGEDIR_CHAR = Chr$(74)
    CHANGEPETDIR_CHAR = Chr$(75)
    FLASHEVENT_CHAR = Chr$(76)
    PROMPT_CHAR = Chr$(77)
    SPEECHEDITOR_CHAR = Chr$(78)
    SPEECH_CHAR = Chr$(79)
    EDITSPEECH_CHAR = Chr$(80)
    EMOTICONEDITOR_CHAR = Chr$(81)
    UPDATEEMOTICON_CHAR = Chr$(82)
    EDITEMOTICON_CHAR = Chr$(83)
    CLEARTEMPTILE_CHAR = Chr$(84)
    FRIENDLIST_CHAR = Chr$(85)
    ARROWEDITOR_CHAR = Chr$(86)
    UPDATEARROW_CHAR = Chr$(87)
    EDITARROW_CHAR = Chr$(89)
    CHECKARROWS_CHAR = Chr$(90)
    CHECKSPRITE_CHAR = Chr$(91)
    MAPREPORT_CHAR = Chr$(92)
    SPELLANIM_CHAR = Chr$(93)
    CHECKEMOTICONS_CHAR = Chr$(94)
    DAMAGEDISPLAY_CHAR = Chr$(95)
    ITEMBREAK_CHAR = Chr$(96)
    GETINFO_CHAR = Chr$(97)
    GATCLASSES_CHAR = Chr$(98)
    NEWFACCOUNTIED_CHAR = Chr$(99)
    DELIMACCOUNTED_CHAR = Chr$(100)
    LOGINATION_CHAR = Chr$(101)
    ADDACHARA_CHAR = Chr$(102)
    DELIMBOCHARU_CHAR = Chr$(103)
    USAGAKRIM_CHAR = Chr$(104)
    GUILDCHANGEACCESS_CHAR = Chr$(105)
    GUILDDISOWN_CHAR = Chr$(106)
    GUILDLEAVE_CHAR = Chr$(107)
    MAKEGUILD_CHAR = Chr$(108)
    GUILDMEMBER_CHAR = Chr$(109)
    GUILDTRAINEE_CHAR = Chr$(110)
    EMOTEMSG_CHAR = Chr$(111)
    BROADCASTMSG_CHAR = Chr$(112)
    GLOBALMSG_CHAR = Chr$(113)
    ADMINMSG_CHAR = Chr$(114)
    PLAYERMSG_CHAR = Chr$(115)
    USEITEM_CHAR = Chr$(116)
    ATTACK_CHAR = Chr$(117)
    USESTATPOINT_CHAR = Chr$(118)
    PLAYERINFOREQUEST_CHAR = Chr$(119)
    SETSPRITE_CHAR = Chr$(120)
    SETPLAYERSPRITE_CHAR = Chr$(121)
    GETSTATS_CHAR = Chr$(122)
    REQUESTNEWMAP_CHAR = Chr$(123)
    NEEDMAP_CHAR = Chr$(124)
    MAPGETITEM_CHAR = Chr$(125)
    MAPDROPITEM_CHAR = Chr$(126)
    MAPRESPAWN_CHAR = Chr$(127)
    KICKPLAYER_CHAR = Chr$(128)
    BANLIST_CHAR = Chr$(129)
    BANDESTROY_CHAR = Chr$(130)
    BANPLAYER_CHAR = Chr$(131)
    REQUESTEDITMAP_CHAR = Chr$(132)
    REQUESTEDITITEM_CHAR = Chr$(133)
    SAVEITEM_CHAR = Chr$(134)
    REQUESTEDITNPC_CHAR = Chr$(135)
    SAVENPC_CHAR = Chr$(136)
    REQUESTEDITQUEST_CHAR = Chr$(137)
    REQUESTEDITSHOP_CHAR = Chr$(138)
    ADDFRIEND_CHAR = Chr$(139)
    REMOVEFRIEND_CHAR = Chr$(140)
    SAVESHOP_CHAR = Chr$(141)
    REQUESTEDITMAIN_CHAR = Chr$(142)
    REQUESTEDITSPELL_CHAR = Chr$(143)
    SAVESPELL_CHAR = Chr$(144)
    SETACCESS_CHAR = Chr$(145)
    WHOSONLINE_CHAR = Chr$(146)
    SETMOTD_CHAR = Chr$(147)
    TRADEREQUEST_CHAR = Chr$(148)
    FIXITEM_CHAR = Chr$(149)
    SEARCH_CHAR = Chr$(150)
    PLAYERCHAT_CHAR = Chr$(151)
    ACHAT_CHAR = Chr$(152)
    DCHAT_CHAR = Chr$(153)
    ATRADE_CHAR = Chr$(154)
    DTRADE_CHAR = Chr$(155)
    UPDATETRADEINV_CHAR = Chr$(156)
    SWAPITEMS_CHAR = Chr$(157)
    PARTY_CHAR = Chr$(158)
    JOINPARTY_CHAR = Chr$(159)
    LEAVEPARTY_CHAR = Chr$(160)
    PARTYCHAT_CHAR = Chr$(161)
    GUILDCHAT_CHAR = Chr$(162)
    NEWMAIN_CHAR = Chr$(163)
    REQUESTBACKUPMAIN_CHAR = Chr$(164)
    CAST_CHAR = Chr$(165)
    REQUESTLOCATION_CHAR = Chr$(167)
    KILLPET_CHAR = Chr$(168)
    REFRESH_CHAR = Chr$(169)
    PETMOVESELECT_CHAR = Chr$(170)
    BUYSPRITE_CHAR = Chr$(171)
    CHECKCOMMANDS_CHAR = Chr$(172)
    REQUESTEDITARROW_CHAR = Chr$(173)
    SAVEARROW_CHAR = Chr$(174)
    SPEECHSCRIPT_CHAR = Chr$(175)
    REQUESTEDITSPEECH_CHAR = Chr$(176)
    SAVESPEECH_CHAR = Chr$(177)
    NEEDSPEECH_CHAR = Chr$(178)
    REQUESTEDITEMOTICON_CHAR = Chr$(179)
    SAVEEMOTICON_CHAR = Chr$(180)
    GMTIME_CHAR = Chr$(181)
    WARPTO_CHAR = Chr$(182)
    WARPTOME_CHAR = Chr$(183)
    WARPPLAYER_CHAR = Chr$(184)
    ARROWHIT_CHAR = Chr$(185)
    PPCHATTING_CHAR = Chr$(186)
    TEMPTILE_CHAR = Chr$(187)
    TEMPATTRIBUTE_CHAR = Chr$(188)
    LEVELUP_CHAR = Chr$(189)
    GATGLASSES_CHAR = Chr$(190)
    USAGAKARIM_CHAR = Chr$(191)
    MAPMSG_CHAR = Chr$(192)
    PPTRADE_CHAR = Chr$(193)
    NEWPARTY_CHAR = Chr$(194)
    FORGETSPELL_CHAR = Chr$(195)
    RETURNSCRIPT_CHAR = Chr$(196)
    'CLOSINGDOWN_CHAR = Chr$(197)
    
    SEP_CHAR = Chr$(253)
    END_CHAR = Chr$(254)
    NEXT_CHAR = Chr$(255)

    'Call InitServer

    CurrentLoad = 0
    Randomize Timer
    'nid.cbSize = Len(nid)
    'nid.hwnd = frmServer.hwnd
    'nid.uId = vbNull
    'nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    'nid.uCallBackMessage = WM_MOUSEMOVE
    'nid.hIcon = frmServer.Icon
    'nid.szTip = GAME_NAME & " Server" & vbNullChar

    ' Add to the sys tray
    'Call Shell_NotifyIcon(NIM_ADD, nid)
    'TrayAdd frmServer, Server_BuildToolTipString, MouseMove

    ' Init atmosphere
    GameWeather = WEATHER_NONE
    WeatherSeconds = 0
    GameTime = TIME_DAY
    TimeSeconds = 0
    RainIntensity = 25

    ' Check if the maps directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\maps", vbDirectory)) <> "maps" Then
        Call MkDir$(App.Path & "\Maps")
    End If

    If LCase$(Dir$(App.Path & "\logs", vbDirectory)) <> "logs" Then
        Call MkDir$(App.Path & "\Logs")
    End If

    ' Check if the accounts directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\accounts", vbDirectory)) <> "accounts" Then
        Call MkDir$(App.Path & "\Accounts")
    End If

    If LCase$(Dir$(App.Path & "\npcs", vbDirectory)) <> "npcs" Then
        Call MkDir$(App.Path & "\Npcs")
    End If

    If LCase$(Dir$(App.Path & "\items", vbDirectory)) <> "items" Then
        Call MkDir$(App.Path & "\Items")
    End If

    If LCase$(Dir$(App.Path & "\spells", vbDirectory)) <> "spells" Then
        Call MkDir$(App.Path & "\Spells")
    End If

    If LCase$(Dir$(App.Path & "\shops", vbDirectory)) <> "shops" Then
        Call MkDir$(App.Path & "\Shops")
    End If

    If LCase$(Dir$(App.Path & "\speech", vbDirectory)) <> "speech" Then
        Call MkDir$(App.Path & "\Speech")
    End If

    ServerLog = True

    If Not FileExist("Data.ini") Then
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "GameName", "Elysium"
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "WebSite", "http://www.elysiumsource.net"
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "Port", 4000
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "HPRegen", 1
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "MPRegen", 1
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "SPRegen", 1
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "Scrolling", 1

        'SpecialPutVar App.Path & "\Data.ini", "CONFIG", "AutoTurn", 0
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "SCRIPTING", 1
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_PLAYERS", 25
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_ITEMS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_NPCS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_SHOPS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_SPELLS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_MAPS", 200
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS", 20
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_GUILDS", 20
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS", 10
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS", 10
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_LEVEL", 500
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_PARTIES", 20
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS", 4
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_SPEECH", 25
    End If

    If Not FileExist("Stats.ini") Then
        SpecialPutVar App.Path & "\Stats.ini", "HP", "AddPerLevel", 10
        SpecialPutVar App.Path & "\Stats.ini", "HP", "AddPerstr", 10
        SpecialPutVar App.Path & "\Stats.ini", "HP", "AddPerDef", 0
        SpecialPutVar App.Path & "\Stats.ini", "HP", "AddPerMagi", 0
        SpecialPutVar App.Path & "\Stats.ini", "HP", "AddPerSpeed", 0
        SpecialPutVar App.Path & "\Stats.ini", "MP", "AddPerLevel", 10
        SpecialPutVar App.Path & "\Stats.ini", "MP", "AddPerstr", 0
        SpecialPutVar App.Path & "\Stats.ini", "MP", "AddPerDef", 0
        SpecialPutVar App.Path & "\Stats.ini", "MP", "AddPerMagi", 10
        SpecialPutVar App.Path & "\Stats.ini", "MP", "AddPerSpeed", 0
        SpecialPutVar App.Path & "\Stats.ini", "SP", "AddPerLevel", 10
        SpecialPutVar App.Path & "\Stats.ini", "SP", "AddPerstr", 0
        SpecialPutVar App.Path & "\Stats.ini", "SP", "AddPerDef", 0
        SpecialPutVar App.Path & "\Stats.ini", "SP", "AddPerMagi", 0
        SpecialPutVar App.Path & "\Stats.ini", "SP", "AddPerSpeed", 20
    End If
    
    Load frmServer
    frmServer.Show

    Call SetStatus("Loading settings...")
    AddHP.Level = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerLevel"))
    AddHP.STR = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerstr"))
    AddHP.DEF = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerDef"))
    AddHP.Magi = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerMagi"))
    AddHP.Speed = Val(GetVar(App.Path & "\Stats.ini", "HP", "AddPerSpeed"))
    AddMP.Level = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerLevel"))
    AddMP.STR = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerstr"))
    AddMP.DEF = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerDef"))
    AddMP.Magi = Val(GetVar(App.Path & "\Stats.ini", "MP", "AddPerMagi"))
    AddMP.Speed = Val(GetVar(App.Path & "\Stats.ini", "MP", vbNullString))
    AddSP.Level = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerLevel"))
    AddSP.STR = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerstr"))
    AddSP.DEF = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerDef"))
    AddSP.Magi = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerMagi"))
    AddSP.Speed = Val(GetVar(App.Path & "\Stats.ini", "SP", "AddPerSpeed"))
    GAME_NAME = Trim$(GetVar(App.Path & "\Data.ini", "CONFIG", "GameName"))
    MAX_PLAYERS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_PLAYERS"))
    MAX_ITEMS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_ITEMS"))
    MAX_NPCS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_NPCS"))
    MAX_SHOPS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_SHOPS"))
    MAX_SPELLS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_SPELLS"))
    MAX_MAPS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAPS"))
    MAX_MAP_ITEMS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS"))
    MAX_GUILDS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILDS"))
    MAX_GUILD_MEMBERS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS"))
    MAX_EMOTICONS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS"))
    MAX_LEVEL = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_LEVEL"))
    MAX_PARTIES = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_PARTIES"))
    MAX_PARTY_MEMBERS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS"))
    MAX_SPEECH = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_SPEECH"))
    SCRIPTING = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "SCRIPTING"))
    HPRegenOn = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "HPRegen"))
    MPRegenOn = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "MPRegen"))
    SPRegenOn = Val(GetVar(App.Path & "\Data.ini", "CONFIG", "SPRegen"))
    MAX_MAPX = 30
    MAX_MAPY = 30

    If GetVar(App.Path & "\Data.ini", "CONFIG", "Scrolling") = NO Then
        MAX_MAPX = 19
        MAX_MAPY = 13
    Else
        MAX_MAPX = 30
        MAX_MAPY = 30
    End If

    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim TempTile(1 To MAX_MAPS) As TempTileRec
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim Item(0 To MAX_ITEMS) As ItemRec
    ReDim Npc(0 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
    ReDim Grid(1 To MAX_MAPS) As GridRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Guild(1 To MAX_GUILDS) As GuildRec
    ReDim Party(1 To MAX_PARTIES) As PartyRec
    ReDim Speech(1 To MAX_SPEECH) As SpeechRec
    ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec

    For i = 1 To MAX_GUILDS
        ReDim Guild(i).Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
    Next

    For i = 1 To MAX_PARTIES
        ReDim Party(i).Member(1 To MAX_PARTY_MEMBERS) As Long
    Next

    For i = 1 To MAX_MAPS
        ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim TempTile(i).DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY) As Byte
        ReDim Grid(i).Loc(0 To MAX_MAPX, 0 To MAX_MAPY) As MapGridRec
    Next

    ReDim Experience(1 To MAX_LEVEL) As Long
    START_MAP = 1
    START_X = MAX_MAPX / 2
    START_Y = MAX_MAPY / 2
    GAME_PORT = GetVar(App.Path & "\Data.ini", "CONFIG", "Port")

    'SCRIPTING
    If SCRIPTING = 1 Then
        Call SetStatus("Loading scripts...")
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    End If

    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_PORT

    ' Init all the player sockets
    For i = 1 To MAX_PLAYERS
        Call SetStatus("Initializing player array...")
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    'For i = 1 To MAX_PLAYERS
    '    Call ShowPLR(i)
    'Next

    'If Not FileExist("CMessages.ini") Then

    '    For i = 1 To 6
    '        PutVar App.Path & "\CMessages.ini", "MESSAGES", "Title" & i, "Custom Msg"
    '        PutVar App.Path & "\CMessages.ini", "MESSAGES", "Message" & i, vbNullString
    '    Next

    'End If

    'For i = 1 To 6
    '    CMessages(i).Title = GetVar(App.Path & "\CMessages.ini", "MESSAGES", "Title" & i)
    '    CMessages(i).Message = GetVar(App.Path & "\CMessages.ini", "MESSAGES", "Message" & i)
        'frmServer.CustomMsg(i - 1).Caption = CMessages(i).Title
    'Next

    'frmServer.lstTopics.Clear
    'i = 1

    'Do While FileExist("Guide\" & i & ".txt")
    '    f = FreeFile
    '    Open App.Path & "\Guide\" & i & ".txt" For Input As #f
    '    Line Input #f, stringy
    '    frmServer.lstTopics.AddItem (stringy)
    '    Close #f
    '    i = i + 1
    'Loop

    'frmServer.lstTopics.Selected(0) = True
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTile
    'Call SetStatus("Clearing maps...")
    'Call ClearMaps
    'Call SetStatus("Clearing map items...")
    'Call ClearMapItems
    'Call SetStatus("Clearing map npcs...")
    'Call ClearMapNpcs
    'Call SetStatus("Clearing npcs...")
    'Call ClearNpcs
    'Call SetStatus("Clearing items...")
    'Call ClearItems
    'Call SetStatus("Clearing shops...")
    'Call ClearShops
    'Call SetStatus("Clearing spells...")
    'Call ClearSpells
    'Call SetStatus("Clearing exp...")
    'Call ClearExps
    'Call SetStatus("Clearing emoticons...")
    'Call ClearEmos
    'Call SetStatus("Clearing parties...")
    'Call ClearParties
    'Call SetStatus("Clearing speech...")
    'Call ClearSpeeches
    Call SetStatus("Loading emoticons...")
    Call LoadEmos
    'Call SetStatus("Clearing arrows...")
    'Call ClearArrows
    Call SetStatus("Loading arrows...")
    Call LoadArrows
    Call SetStatus("Loading exp...")
    Call LoadExps
    Call SetStatus("Loading classes...")
    Call LoadClasses

    'Call SetStatus("Loading first class advandcement...")
    'Call LoadClasses2
    'Call SetStatus("Loading second class advandcement...")
    'Call Loadclasses3
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading speeches...")
    Call LoadSpeeches
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Setting up the grid...")
    Call SetUpGrid
    'frmServer.MapList.Clear

    'For i = 1 To MAX_MAPS
    '    frmServer.MapList.AddItem i & ": " & Map(i).Name
    'Next

    'frmServer.MapList.Selected(0) = True

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("accounts\charlist.txt") Then
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #f
        Close #f
    End If

    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    
    ServerStartTime = GetTickCount
    
    ' Add to the tray
    TrayAdd frmServer, Server_BuildToolTipString, MouseMove

    frmServer.Hide
    DoEvents
    SpawnSeconds = 0
    'frmServer.tmrMasterTimer.Enabled = True
    
    'Lets start the server!
    ServerLoop
    
    Exit Sub
    
ErrHandler:
    Call AddLog("Error in frmServer on start-up.", "errorlist.txt")
    MsgBox "There was an error on start-up! Sorry, gotta close...", vbOKOnly, "Error"
    End
End Sub

'Sub CheckGiveHP()
'    Dim i As Long

'    If GetTickCount > GiveHPTimer + 10000 Then

'        For i = 1 To MAX_PLAYERS

'            If IsPlaying(i) Then
'                If GetPlayerHP(i) < GetPlayerMaxHP(i) And GetPlayerHP(i) >= 0 Then
'                    Call SetPlayerHP(i, GetPlayerHP(i) + GetPlayerHPRegen(i))
'                    Call SendHP(i)
'                End If

'                If GetPlayerMP(i) < GetPlayerMaxMP(i) And GetPlayerMP(i) >= 0 Then
'                    Call SetPlayerMP(i, GetPlayerMP(i) + GetPlayerMPRegen(i))
'                    Call SendMP(i)
'                End If

'                If GetPlayerSP(i) < GetPlayerMaxSP(i) And GetPlayerSP(i) >= 0 Then
'                    Call SetPlayerSP(i, GetPlayerSP(i) + GetPlayerSPRegen(i))
'                    Call SendSP(i)
'                End If
'            End If

'            DoEvents
'        Next

'        GiveHPTimer = GetTickCount
'    End If

'End Sub

Sub CheckSpawnMapItems()
    Dim X As Long, Y As Long

    ' Used for map item respawning
    SpawnSeconds = SpawnSeconds + 1

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    If SpawnSeconds >= 120 Then

        ' 2 minutes have passed
        For Y = 1 To MAX_MAPS

            ' Make sure no one is on the map when it respawns
            If PlayersOnMap(Y) = False Then

                ' Clear out unnecessary junk
                For X = 1 To MAX_MAP_ITEMS
                    Call ClearMapItem(X, Y)
                Next

                ' Spawn the items
                Call SpawnMapItems(Y)
                Call SendMapItemsToAll(Y)
            End If

            DoEvents
        Next

        SpawnSeconds = 0
    End If

End Sub

Sub DestroyServer()
On Error Resume Next

    Dim i As Long

    ' Say bye to the system tray
    'Call Shell_NotifyIcon(NIM_DELETE, nid)
    TrayDelete
    Call SetStatus("Shutting down...")
    frmServer.Visible = True
    'frmServer.Visible = False

    DoEvents
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Unloading sockets and timers...")

    For i = 1 To MAX_PLAYERS
        Call SetStatus("Unloading sockets and timers... " & i & "/" & MAX_PLAYERS)

        DoEvents
        Unload frmServer.Socket(i)
    Next
    
    frmServer.Hide
    DoEvents
    Unload frmServer

    'If frmServer.chkChat.value = Checked Then
    '    Call SetStatus("Saving chat logs...")
    '    Call SaveLogs
    'End If
    End
End Sub

Sub GameAI()
On Error GoTo ErrHandler

    Dim i As Long, X As Long, Y As Long, N As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, TickCount As Long
    Dim Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long, Target As Long
    Dim DidWalk As Boolean

    'WeatherSeconds = WeatherSeconds + 1
    'TimeSeconds = TimeSeconds + 1
    ' Lets change the weather if its time to
    If WeatherSeconds >= 60 Then
        i = Int(Rnd * 3)

        If i <> GameWeather Then
            GameWeather = i
            Call SendWeatherToAll
        End If

        WeatherSeconds = 0
    End If

    ' Check if we need to switch from day to night or night to day
    If TimeSeconds >= 60 Then
        If GameTime = TIME_DAY Then
            GameTime = TIME_NIGHT
        Else
            GameTime = TIME_DAY
        End If

        Call SendTimeToAll
        TimeSeconds = 0
    End If

    For Y = 1 To MAX_MAPS

        If PlayersOnMap(Y) = YES Then
            TickCount = GetTickCount

            ' ////////////////////////////////////
            ' // This is used for closing doors //
            ' ////////////////////////////////////
            If TickCount > TempTile(Y).DoorTimer + 5000 Then

                For y1 = 0 To MAX_MAPY
                    For x1 = 0 To MAX_MAPX

                        If Map(Y).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(Y).DoorOpen(x1, y1) = YES Then
                            TempTile(Y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(Y, MAPKEY_CHAR & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & END_CHAR)
                        End If

                        If Map(Y).Tile(x1, y1).Type = TILE_TYPE_DOOR And TempTile(Y).DoorOpen(x1, y1) = YES Then
                            TempTile(Y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(Y, MAPKEY_CHAR & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & END_CHAR)
                        End If

                    Next
                Next

            End If

            For X = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(Y, X).num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).Npc(X) > 0 And MapNpc(Y, X).num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then

                        For i = 1 To MAX_PLAYERS

                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = Y And MapNpc(Y, X).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITER Then
                                    N = Npc(NpcNum).Range
                                    DistanceX = MapNpc(Y, X).X - GetPlayerX(i)
                                    DistanceY = MapNpc(Y, X).Y - GetPlayerY(i)

                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1

                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= N And DistanceY <= N Then
                                        If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                            If Trim$(Npc(NpcNum).AttackSay) <> vbNullString Then
                                                Call PlayerMsg(i, "A " & Trim$(Npc(NpcNum).Name) & ": " & Trim$(Npc(NpcNum).AttackSay) & vbNullString, SayColor)
                                            End If

                                            MapNpc(Y, X).TargetType = TARGET_TYPE_PLAYER
                                            MapNpc(Y, X).Target = i
                                        End If
                                    End If
                                End If
                            End If

                        Next

                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).Npc(X) > 0 And MapNpc(Y, X).num > 0 Then
                    Target = MapNpc(Y, X).Target

                    ' Check to see if its time for the npc to walk
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then

                        ' Check to see if we are following a player or not
                        If Target > 0 Then
                            If MapNpc(Y, X).TargetType = TARGET_TYPE_PLAYER Then

                                ' Check if the player is even playing, if so follow'm
                                If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                    DidWalk = False
                                    i = Int(Rnd * 5)

                                    ' Lets move the npc
                                    Select Case i

                                        Case 0

                                            ' Up
                                            If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 1

                                            ' Right
                                            If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 2

                                            ' Down
                                            If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 3

                                            ' Left
                                            If MapNpc(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNpc(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNpc(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNpc(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                    End Select

                                    ' Check if we can't move and if player is behind something and if we can just switch dirs
                                    If Not DidWalk Then
                                        If MapNpc(Y, X).X - 1 = GetPlayerX(Target) And MapNpc(Y, X).Y = GetPlayerY(Target) Then
                                            If MapNpc(Y, X).Dir <> DIR_LEFT Then
                                                Call NpcDir(Y, X, DIR_LEFT)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapNpc(Y, X).X + 1 = GetPlayerX(Target) And MapNpc(Y, X).Y = GetPlayerY(Target) Then
                                            If MapNpc(Y, X).Dir <> DIR_RIGHT Then
                                                Call NpcDir(Y, X, DIR_RIGHT)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapNpc(Y, X).X = GetPlayerX(Target) And MapNpc(Y, X).Y - 1 = GetPlayerY(Target) Then
                                            If MapNpc(Y, X).Dir <> DIR_UP Then
                                                Call NpcDir(Y, X, DIR_UP)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapNpc(Y, X).X = GetPlayerX(Target) And MapNpc(Y, X).Y + 1 = GetPlayerY(Target) Then
                                            If MapNpc(Y, X).Dir <> DIR_DOWN Then
                                                Call NpcDir(Y, X, DIR_DOWN)
                                            End If

                                            DidWalk = True
                                        End If

                                        ' We could not move so player must be behind something, walk randomly.
                                        If Not DidWalk Then
                                            i = Int(Rnd * 2)

                                            If i = 1 Then
                                                i = Int(Rnd * 4)

                                                If CanNpcMove(Y, X, i) Then
                                                    Call NpcMove(Y, X, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If

                                Else
                                    MapNpc(Y, X).Target = 0
                                End If

                            Else

                                ' Check if the pet is even playing, if so follow'm
                                If IsPlaying(Target) And Player(Target).Pet.Map = Y Then
                                    DidWalk = False
                                    i = Int(Rnd * 5)

                                    ' Lets move the npc
                                    Select Case i

                                        Case 0

                                            ' Up
                                            If MapNpc(Y, X).Y > Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNpc(Y, X).Y < Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNpc(Y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNpc(Y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 1

                                            ' Right
                                            If MapNpc(Y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNpc(Y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNpc(Y, X).Y < Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNpc(Y, X).Y > Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 2

                                            ' Down
                                            If MapNpc(Y, X).Y < Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNpc(Y, X).Y > Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNpc(Y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNpc(Y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 3

                                            ' Left
                                            If MapNpc(Y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNpc(Y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNpc(Y, X).Y > Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNpc(Y, X).Y < Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                    End Select

                                    ' Check if we can't move and if pet is behind something and if we can just switch dirs
                                    If Not DidWalk Then
                                        If MapNpc(Y, X).X - 1 = Player(Target).Pet.X And MapNpc(Y, X).Y = Player(Target).Pet.Y Then
                                            If MapNpc(Y, X).Dir <> DIR_LEFT Then
                                                Call NpcDir(Y, X, DIR_LEFT)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapNpc(Y, X).X + 1 = Player(Target).Pet.X And MapNpc(Y, X).Y = Player(Target).Pet.Y Then
                                            If MapNpc(Y, X).Dir <> DIR_RIGHT Then
                                                Call NpcDir(Y, X, DIR_RIGHT)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapNpc(Y, X).X = Player(Target).Pet.X And MapNpc(Y, X).Y - 1 = Player(Target).Pet.Y Then
                                            If MapNpc(Y, X).Dir <> DIR_UP Then
                                                Call NpcDir(Y, X, DIR_UP)
                                            End If

                                            DidWalk = True
                                        End If

                                        If MapNpc(Y, X).X = Player(Target).Pet.X And MapNpc(Y, X).Y + 1 = Player(Target).Pet.Y Then
                                            If MapNpc(Y, X).Dir <> DIR_DOWN Then
                                                Call NpcDir(Y, X, DIR_DOWN)
                                            End If

                                            DidWalk = True
                                        End If

                                        ' We could not move so pet must be behind something, walk randomly.
                                        If Not DidWalk Then
                                            i = Int(Rnd * 2)

                                            If i = 1 Then
                                                i = Int(Rnd * 4)

                                                If CanNpcMove(Y, X, i) Then
                                                    Call NpcMove(Y, X, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If

                                Else
                                    MapNpc(Y, X).Target = 0
                                End If
                            End If

                        Else
                            i = Int(Rnd * 4)

                            If i = 1 Then
                                i = Int(Rnd * 4)

                                If CanNpcMove(Y, X, i) Then
                                    Call NpcMove(Y, X, i, MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
                End If

                ' //////////////////////////////////////////////////////
                ' // This is used for npcs to attack players and pets //
                ' //////////////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).Npc(X) > 0 And MapNpc(Y, X).num > 0 Then
                    Target = MapNpc(Y, X).Target

                    If MapNpc(Y, X).TargetType <> TARGET_TYPE_LOCATION And MapNpc(Y, X).TargetType <> TARGET_TYPE_NPC Then

                        ' Check if the npc can attack the targeted player player
                        If Target > 0 Then
                            If MapNpc(Y, X).TargetType = TARGET_TYPE_PLAYER Then

                                ' Is the target playing and on the same map?
                                If IsPlaying(Target) And GetPlayerMap(Target) = Y Then

                                    ' Can the npc attack the player?
                                    If CanNpcAttackPlayer(X, Target) Then
                                        If Not CanPlayerBlockHit(Target) Then
                                            Damage = Npc(NpcNum).STR - GetPlayerProtection(Target) + (Rnd * 5) - 2

                                            If Damage > 0 Then
                                                Call NpcAttackPlayer(X, Target, Damage)
                                            Else
                                                Call BattleMsg(Target, "The " & Trim$(Npc(NpcNum).Name) & " couldn't hurt you!", BrightBlue, 1)

                                                'Call PlayerMsg(Target, "The " & Trim$(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                            End If

                                        Else
                                            Call BattleMsg(Target, "You blocked the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan, 1)

                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                        End If
                                    End If

                                Else

                                    ' Player left map or game, set target to 0
                                    MapNpc(Y, X).Target = 0
                                End If

                            Else

                                ' Is the target playing and on the same map?
                                If IsPlaying(Target) And Player(Target).Pet.Map = Y Then

                                    ' Can the npc attack the pet?
                                    If CanNpcAttackPet(X, Target) Then
                                        Damage = Npc(NpcNum).STR - Player(Target).Pet.Level + (Rnd * 5) - 2

                                        If Damage > 0 Then
                                            Call NpcAttackPet(X, Target, Damage)
                                        End If
                                    End If

                                Else

                                    ' Pet left map or game, set target to 0
                                    MapNpc(Y, X).Target = 0
                                End If
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNpc(Y, X).num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNpc(Y, X).HP > 0 Then
                        MapNpc(Y, X).HP = MapNpc(Y, X).HP + GetNpcHPRegen(NpcNum)

                        ' Check if they have more then they should and if so just set it to max
                        If MapNpc(Y, X).HP > GetNpcMaxHP(NpcNum) Then
                            MapNpc(Y, X).HP = GetNpcMaxHP(NpcNum)
                        End If

                        Call SendDataToMap(Y, NPCHP_CHAR & SEP_CHAR & X & SEP_CHAR & MapNpc(Y, X).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(Y, X).num) & END_CHAR)
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).str > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(Y, X).num = 0 And Map(Y).Npc(X) > 0 Then
                    If TickCount > MapNpc(Y, X).SpawnWait + (Npc(Map(Y).Npc(X)).SpawnSecs * 1000) Then
                        Call SpawnNpc(X, Y)
                    End If
                End If

                If MapNpc(Y, X).num > 0 Then

                    ' If the NPC hasn't been fighting, why send it's HP?
                    If GetTickCount < MapNpc(Y, X).LastAttack + 6000 Then
                        Call SendDataToMap(Y, NPCHP_CHAR & SEP_CHAR & X & SEP_CHAR & MapNpc(Y, X).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(Y, X).num) & END_CHAR)
                    End If
                End If

            Next

        End If

        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

    ' //////////////////////////////////////////////////////////
    ' // Used for moving pets (it took a while it get going!) //
    ' //////////////////////////////////////////////////////////
    For X = 1 To MAX_PLAYERS

        If Player(X).Pet.Alive = YES Then
            x1 = Player(X).Pet.X
            y1 = Player(X).Pet.Y
            x2 = Player(X).Pet.XToGo
            y2 = Player(X).Pet.YToGo

            If Player(X).Pet.Target > 0 Then
                If Player(X).Pet.TargetType = TARGET_TYPE_PLAYER Then
                    x2 = GetPlayerX(Player(X).Pet.Target)
                    y2 = GetPlayerY(Player(X).Pet.Target)
                End If

                If Player(X).Pet.TargetType = TARGET_TYPE_NPC Then
                    If CanPetAttackNpc(X, Player(X).Pet.Target) Then
                        Damage = Player(X).Pet.Level - Npc(Player(X).Pet.Target).STR + (Rnd * 5) - 2

                        If Damage > 0 Then
                            Call PetAttackNpc(X, Player(X).Pet.Target, Damage)
                            x2 = x1
                            y2 = y1
                        End If

                    Else
                        x2 = MapNpc(Player(X).Pet.Map, Player(X).Pet.Target).X
                        y2 = MapNpc(Player(X).Pet.Map, Player(X).Pet.Target).Y
                    End If
                End If

            Else

                If Player(X).Pet.Map = GetPlayerMap(X) Or Player(X).Pet.MapToGo = 0 Then
                    If Player(X).Pet.XToGo = -1 Or Player(X).Pet.YToGo = -1 Then
                        i = Int(Rnd * 4)

                        If i = 1 Then
                            i = Int(Rnd * 4)

                            If i = DIR_UP Then
                                y2 = y1 - 1
                                x2 = Player(X).Pet.X
                            End If

                            If i = DIR_DOWN Then
                                y2 = y1 + 1
                                x2 = Player(X).Pet.X
                            End If

                            If i = DIR_RIGHT Then
                                x2 = x1 + 1
                                y2 = Player(X).Pet.Y
                            End If

                            If i = DIR_LEFT Then
                                x2 = x1 - 1
                                y2 = Player(X).Pet.Y
                            End If

                            If Not IsValid(x2, y2) Then
                                x2 = x1
                                y2 = y1
                            End If

                            If Grid(Player(X).Pet.Map).Loc(x2, y2).Blocked = True Then
                                x2 = x1
                                y2 = y1
                            End If

                        Else
                            x2 = x1
                            y2 = y1
                        End If
                    End If

                Else

                    If Map(Player(X).Pet.Map).Up = Player(X).Pet.MapToGo Then
                        y2 = y1 - 1
                    Else

                        If Map(Player(X).Pet.Map).Down = Player(X).Pet.MapToGo Then
                            y2 = y1 + 1
                        Else

                            If Map(Player(X).Pet.Map).Left = Player(X).Pet.MapToGo Then
                                x2 = x1 - 1
                            Else

                                If Map(Player(X).Pet.Map).Right = Player(X).Pet.MapToGo Then
                                    x2 = x1 + 1
                                Else
                                    i = Int(Rnd * 4)

                                    If i = 1 Then
                                        i = Int(Rnd * 4)

                                        If i = DIR_UP Then y2 = y1 - 1
                                        If i = DIR_DOWN Then y2 = y1 + 1
                                        If i = DIR_RIGHT Then x2 = x1 + 1
                                        If i = DIR_LEFT Then x2 = x1 - 1
                                        If Not IsValid(x2, y2) Then
                                            x2 = x1
                                            y2 = y1
                                        End If

                                        If Grid(Player(X).Pet.Map).Loc(x2, y2).Blocked = True Then
                                            x2 = x1
                                            y2 = y1
                                        End If

                                    Else
                                        x2 = x1
                                        y2 = y1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If x1 < x2 Then

                ' RIGHT not left
                If y1 < y2 Then

                    ' DOWN not up
                    If x2 - x1 > y2 - y1 Then

                        ' RIGHT not down
                        If CanPetMove(X, DIR_RIGHT) Then

                            ' RIGHT works
                            Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                        Else

                            If CanPetMove(X, DIR_DOWN) Then

                                ' DOWN works and right doesn't
                                Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                            Else

                                ' Nothing works, random time
                                i = Int(Rnd * 4)

                                If CanPetMove(X, i) Then
                                    Call PetMove(X, i, MOVING_WALKING)
                                End If
                            End If
                        End If

                    Else

                        If x2 - x1 <> y2 - y1 Then

                            ' DOWN not right
                            If CanPetMove(X, DIR_DOWN) Then

                                ' DOWN works
                                Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                            Else

                                If CanPetMove(X, DIR_RIGHT) Then

                                    ' RIGHT works and down doesn't
                                    Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    i = Int(Rnd * 4)

                                    If CanPetMove(X, i) Then
                                        Call PetMove(X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If

                        Else

                            ' Both are equal
                            If CanPetMove(X, DIR_RIGHT) Then

                                ' RIGHT works
                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN and RIGHT work
                                    i = (Int(Rnd * 2) * 2) + 1

                                    If CanPetMove(X, i) Then
                                        Call PetMove(X, i, MOVING_WALKING)
                                    End If

                                Else

                                    ' RIGHT works only
                                    Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                End If

                            Else

                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN works only
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    i = Int(Rnd * 4)

                                    If CanPetMove(X, i) Then
                                        Call PetMove(X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        End If
                    End If

                Else

                    If y1 <> y2 Then

                        ' UP not down
                        If x2 - x1 > y1 - y2 Then

                            ' RIGHT not up
                            If CanPetMove(X, DIR_RIGHT) Then

                                ' RIGHT works
                                Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                            Else

                                If CanPetMove(X, DIR_UP) Then

                                    ' UP works and right doesn't
                                    Call PetMove(X, DIR_UP, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    i = Int(Rnd * 4)

                                    If CanPetMove(X, i) Then
                                        Call PetMove(X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If

                        Else

                            If x2 - x1 <> y1 - y2 Then

                                ' UP not right
                                If CanPetMove(X, DIR_UP) Then

                                    ' UP works
                                    Call PetMove(X, DIR_UP, MOVING_WALKING)
                                Else

                                    If CanPetMove(X, DIR_RIGHT) Then

                                        ' RIGHT works and up doesn't
                                        Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        i = Int(Rnd * 4)

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If

                            Else

                                ' Both are equal
                                If CanPetMove(X, DIR_RIGHT) Then

                                    ' RIGHT works
                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP and RIGHT work
                                        i = Int(Rnd * 2) * 3

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If

                                    Else

                                        ' RIGHT works only
                                        Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                    End If

                                Else

                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP works only
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        i = Int(Rnd * 4)

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            End If
                        End If

                    Else

                        ' Target is horizontal
                        If CanPetMove(X, DIR_RIGHT) Then

                            ' RIGHT works
                            Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                        Else

                            ' Right doesn't work
                            If CanPetMove(X, DIR_UP) Then
                                If CanPetMove(X, DIR_DOWN) Then

                                    ' UP and DOWN work
                                    i = Int(Rnd * 2)
                                    Call PetMove(X, i, MOVING_WALKING)
                                Else

                                    ' Only UP works
                                    Call PetMove(X, DIR_UP, MOVING_WALKING)
                                End If

                            Else

                                If CanPetMove(X, DIR_DOWN) Then

                                    ' Only DOWN works
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    ' Nothing works, only left is left (heh)
                                    If CanPetMove(X, DIR_LEFT) Then
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    Else

                                        ' Nothing works at all, let it die
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

            Else

                If x1 <> x2 Then

                    ' LEFT not right
                    If y1 < y2 Then

                        ' DOWN not up
                        If x1 - x2 > y2 - y1 Then

                            ' LEFT not down
                            If CanPetMove(X, DIR_LEFT) Then

                                ' LEFT works
                                Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                            Else

                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN works and left doesn't
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    i = Int(Rnd * 4)

                                    If CanPetMove(X, i) Then
                                        Call PetMove(X, i, MOVING_WALKING)
                                    End If
                                End If
                            End If

                        Else

                            If x1 - x2 <> y2 - y1 Then

                                ' DOWN not left
                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN works
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' LEFT works and down doesn't
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        i = Int(Rnd * 4)

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If

                            Else

                                ' Both are equal
                                If CanPetMove(X, DIR_LEFT) Then

                                    ' LEFT works
                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' DOWN and LEFT work
                                        i = Int(Rnd * 2) + 1
                                        Call PetMove(X, i, MOVING_WALKING)
                                    Else

                                        ' LEFT works only
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    End If

                                Else

                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' DOWN works only
                                        Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        i = Int(Rnd * 4)

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            End If
                        End If

                    Else

                        If y1 <> y2 Then

                            ' UP not down
                            If x1 - x2 > y1 - y2 Then

                                ' LEFT not up
                                If CanPetMove(X, DIR_LEFT) Then

                                    ' LEFT works
                                    Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                Else

                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP works and left doesn't
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        i = Int(Rnd * 4)

                                        If CanPetMove(X, i) Then
                                            Call PetMove(X, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If

                            Else

                                If x1 - x2 <> y1 - y2 Then

                                    ' UP not LEFT
                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP works
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        If CanPetMove(X, DIR_LEFT) Then

                                            ' LEFT works and up doesn't
                                            Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                        Else

                                            ' Nothing works, random time
                                            i = Int(Rnd * 4)

                                            If CanPetMove(X, i) Then
                                                Call PetMove(X, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If

                                Else

                                    ' Both are equal
                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' LEFT works
                                        If CanPetMove(X, DIR_UP) Then

                                            ' UP and LEFT work
                                            i = Int(Rnd * 2) * 2
                                            Call PetMove(X, i, MOVING_WALKING)
                                        Else

                                            ' LEFT works only
                                            Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                        End If

                                    Else

                                        If CanPetMove(X, DIR_UP) Then

                                            ' UP works only
                                            Call PetMove(X, DIR_UP, MOVING_WALKING)
                                        Else

                                            ' Nothing works, random time
                                            i = Int(Rnd * 4)

                                            If CanPetMove(X, i) Then
                                                Call PetMove(X, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                        Else

                            ' Target is horizontal
                            If CanPetMove(X, DIR_LEFT) Then

                                ' LEFT works
                                Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                            Else

                                ' LEFT doesn't work
                                If CanPetMove(X, DIR_UP) Then
                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' UP and DOWN work
                                        i = Int(Rnd * 2)
                                        Call PetMove(X, i, MOVING_WALKING)
                                    Else

                                        ' Only UP works
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    End If

                                Else

                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' Only DOWN works
                                        Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                    Else

                                        ' Nothing works, only right is left (heh)
                                        If CanPetMove(X, DIR_RIGHT) Then
                                            Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                        Else

                                            ' Nothing works at all, let it die
                                            Player(X).Pet.MapToGo = Player(X).Pet.Map
                                            Player(X).Pet.XToGo = -1
                                            Player(X).Pet.YToGo = -1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                Else

                    ' Target is vertical
                    If y1 < y2 Then

                        ' DOWN not up
                        If CanPetMove(X, DIR_DOWN) Then
                            Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                        Else

                            ' Down doesn't work
                            If CanPetMove(X, DIR_RIGHT) Then
                                If CanPetMove(X, DIR_LEFT) Then

                                    ' RIGHT and LEFT work
                                    i = Int((Rnd * 2) + 2)
                                    Call PetMove(X, i, MOVING_WALKING)
                                Else

                                    ' RIGHT works only
                                    Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                End If

                            Else

                                If CanPetMove(X, DIR_LEFT) Then

                                    ' LEFT works only
                                    Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                Else

                                    ' Nothing works, lets try up
                                    If CanPetMove(X, DIR_UP) Then
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        ' Nothing at all works, let it die
                                        Player(X).Pet.MapToGo = Player(X).Pet.Map
                                        Player(X).Pet.XToGo = -1
                                        Player(X).Pet.YToGo = -1
                                    End If
                                End If
                            End If
                        End If

                    Else

                        If y1 <> y2 Then

                            ' UP not down
                            If CanPetMove(X, DIR_UP) Then
                                Call PetMove(X, DIR_UP, MOVING_WALKING)
                            Else

                                ' UP doesn't work
                                If CanPetMove(X, DIR_RIGHT) Then
                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' RIGHT and LEFT work
                                        i = Int((Rnd * 2) + 2)
                                        Call PetMove(X, i, MOVING_WALKING)
                                    Else

                                        ' RIGHT works only
                                        Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                    End If

                                Else

                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' LEFT works only
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    Else

                                        ' Nothing works, lets try down
                                        If CanPetMove(X, DIR_DOWN) Then
                                            Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                        Else

                                            ' Nothing at all works, let it die
                                            Player(X).Pet.MapToGo = Player(X).Pet.Map
                                            Player(X).Pet.XToGo = -1
                                            Player(X).Pet.YToGo = -1
                                        End If
                                    End If
                                End If
                            End If

                        Else

                            ' Question:
                            '   What do we do now?
                            ' Answer:
                            Player(X).Pet.MapToGo = Player(X).Pet.Map
                            Player(X).Pet.XToGo = -1
                            Player(X).Pet.YToGo = -1

                            ' Explaination:
                            '   If y1 - y2 = 0 and x1 - x2 = 0...
                            '   We must be at the location we want to move to!
                            '   Cancel the movement for the future
                        End If
                    End If
                End If
            End If
        End If

    Next
    Exit Sub
    
ErrHandler:
    Call AddLog("Error avoided in Sub GameAI()!", "errorlist.txt")
End Sub

Sub PlayerSaveTimer()
    Static MinPassed As Long
    MinPassed = MinPassed + 1

    If MinPassed >= 60 Then
        If TotalOnlinePlayers > 0 Then

            'Call TextAdd(frmServer.txtText(0), "Saving all online players...", True)
            'Call GlobalMsg("Saving all online players...", Pink)
            'For i = 1 To MAX_PLAYERS
            ' If IsPlaying(i) Then
            ' Call SavePlayer(i)
            ' End If
            ' DoEvents
            'Next
            PlayerI = 1
            PlayerTimer = YES
            tmrPlayerSave = NO
        End If

        MinPassed = 0
    End If

End Sub

Sub PlayerSaveTimer2()

    If PlayerI <= MAX_PLAYERS Then
        If IsPlaying(PlayerI) Then
            Call SavePlayer(PlayerI)
            Call PlayerMsg(PlayerI, GetPlayerName(PlayerI) & " is now saved.", Yellow)
        End If
        PlayerI = PlayerI + 1
    End If

    If PlayerI >= MAX_PLAYERS Then
        PlayerI = 1
        PlayerTimer = YES
        tmrPlayerSave = YES
    End If

End Sub

'Sub ServerLogic()
'    Dim i As Long

    ' Check for disconnections
    'For i = 1 To MAX_PLAYERS

    '    If frmServer.Socket(i).State > 7 Then
    '        Call CloseSocket(i)
    '    End If

    'Next

'    Call CheckGiveHP
    'Call GameAI
'End Sub

Sub SetStatus(ByVal Status As String)
    frmServer.lblStatus.Caption = Status
End Sub

'Function GetIP() As String
'On Error Resume Next
'    Dim ip As String
'    ip = Split(frmServer.Inet1.OpenURL("http://whatismyip.com/"), "<h1>")(1)
'    ip = Split(ip, " ")(3)
'    ip = Trim$(Split(ip, "</h1>")(0))
'
'    If Not ip <> vbNullString Then ip = "Localhost(127.0.0.1)"

'    GetIP = Trim$(ip)
'    ip = ""

'End Function

Public Function Server_BuildToolTipString() As String

'*****************************************************************
'Builds the tooltip string
'*****************************************************************
Dim kBpsIn As Single
Dim kBpsOut As Single

    'Get the number of connections
    'Server_UpdateConnections

    'Display statistics (Kilobytes)
    On Error Resume Next
        kBpsIn = Round((DataKBIn * 0.0009765625) / ((GetTickCount - ServerStartTime) * 0.001), 5)
        kBpsOut = Round((DataKBOut * 0.0009765625) / ((GetTickCount - ServerStartTime) * 0.001), 5)
    On Error GoTo 0

    'Display statistics (Bytes)
    'kBpsIn = Round(((DataKBIn * 1024) + DataIn) / ((timeGetTime - ServerStartTime) / 1000), 5)
    'kBpsOut = Round(((DataKBOut * 1024) + DataOut) / ((timeGetTime - ServerStartTime) / 1000), 5)
    
    'Build the string
    'Server_BuildToolTipString = "Connections: " & TotalOnlinePlayers & vbNewLine & _
    '                            "kBps in: " & kBpsIn & vbNewLine & _
    '                            "kBps out: " & kBpsOut

    Server_BuildToolTipString = "Players on: " & TotalOnlinePlayers & " / " & MAX_PLAYERS & vbNewLine & _
                                "KB/s in: " & kBpsIn & vbNewLine & _
                                "KB/s out: " & kBpsOut

End Function

Public Sub LogChats()
    Static ChatSecs As Long
    Dim SaveTime As Long

    SaveTime = 3600

    'If frmServer.chkChat.Value = Unchecked Then
    '    ChatSecs = SaveTime
    '    Label6.Caption = "Chat Log Save Disabled!"
    '    Exit Sub
    'End If

    If ChatSecs <= 0 Then ChatSecs = SaveTime
    'If ChatSecs > 60 Then
    '    Label6.Caption = "Chat Log Save In " & Int(ChatSecs / 60) & " Minute(s)"
    'Else
    '    Label6.Caption = "Chat Log Save In " & Int(ChatSecs) & " Second(s)"
    'End If
    ChatSecs = ChatSecs - 1

    If ChatSecs <= 0 Then
        Call AddLog("Chat logs have been saved!", "serverlog.txt")
        Call SaveLogs
        ChatSecs = 0
    End If
End Sub

Public Sub SendSound(ByVal Index As Long, ByVal Sound_ID As Byte, ByVal SEND As Byte, Optional ByVal Spell As Long)

    If Sound_ID = MAGIC_SOUND Then
        If Spell > 0 Then
            Call SendDataTo(Index, SOUND_CHAR & SEP_CHAR & MAGIC_SOUND & SEP_CHAR & Spell & END_CHAR)
            Exit Sub
        End If
    End If

    Select Case SEND
    
        Case SDT 'SendDataTo
            Call SendDataTo(Index, SOUND_CHAR & SEP_CHAR & Sound_ID & END_CHAR)
            Exit Sub
        
        Case SDTAB 'SendDataToAllBut
            Call SendDataToAllBut(Index, SOUND_CHAR & SEP_CHAR & Sound_ID & END_CHAR)
            Exit Sub
        
        Case SDTM 'SendDataToMap
            Call SendDataToMap(GetPlayerMap(Index), SOUND_CHAR & SEP_CHAR & Sound_ID & END_CHAR)
            Exit Sub
    
    End Select

End Sub
