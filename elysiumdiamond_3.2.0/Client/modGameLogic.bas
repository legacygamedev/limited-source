Attribute VB_Name = "modGameLogic"

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Private Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const SRCAND As Long = &H8800C6
Public Const SRCCOPY As Long = &HCC0020
Public Const SRCPAINT As Long = &HEE0086

' Translucecy stuff
Public Const GWL_EXSTYLE As Long = (-20)
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const LWA_ALPHA As Long = &H2&

Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

' Menu states
Public Const MENU_STATE_NEWACCOUNT As Byte = 0
Public Const MENU_STATE_DELACCOUNT As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_NEWCHAR As Byte = 4
Public Const MENU_STATE_ADDCHAR As Byte = 5
Public Const MENU_STATE_DELCHAR As Byte = 6
Public Const MENU_STATE_USECHAR As Byte = 7
Public Const MENU_STATE_INIT As Byte = 8

' Speed moving vars
Public Const WALK_SPEED As Byte = 4
Public Const RUN_SPEED As Byte = 8
Public Const GM_WALK_SPEED As Byte = 4
Public Const GM_RUN_SPEED As Byte = 8
'Set the variable to your desire,
'32 is a safe and recommended setting

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Game text buffer
Public MyText As String

' Index of actual player
Public MyIndex As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public EditorSet As Byte

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key opene ditor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long
Public KeyOpenEditorMsg As String

' Map for local use
Public SaveMapItem() As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public InSpeechEditor As Boolean
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InEmoticonEditor As Boolean
Public InArrowEditor As Boolean
Public InSpawnEditor As Boolean
Public EditorIndex As Long

' Used to know what npc we are choosing the spawn for
Public SpawnLocator As Long

' Game fps
Public GameFPS As Long

' Used for atmosphere
Public GameWeather As Long
Public GameTime As Long
Public RainIntensity As Long

' Scrolling Variables
Public NewPlayerX As Long
Public NewPlayerY As Long
Public NewXOffset As Long
Public NewYOffset As Long
Public NewX As Long
Public NewY As Long

' Damage Variables
Public DmgDamage As Long
Public DmgTime As Long
Public NPCDmgDamage As Long
Public NPCDmgTime As Long
Public NPCWho As Long

Public EditorItemX As Long
Public EditorItemY As Long

Public EditorShopNum As Long

Public EditorItemNum1 As Byte
Public EditorItemNum2 As Byte
Public EditorItemNum3 As Byte

Public Arena1 As Byte
Public Arena2 As Byte
Public Arena3 As Byte

Public ii As Long, iii As Long
Public sx As Long

Public MouseDownX As Long
Public MouseDownY As Long

Public SpritePic As Long
Public SpriteItem As Long
Public SpritePrice As Long

Public SoundFileName As String

Public ScreenMode As Byte

Public SignLine1 As String
Public SignLine2 As String
Public SignLine3 As String

Public ClassChange As Long
Public ClassChangeReq As Long

Public NoticeTitle As String
Public NoticeText As String
Public NoticeSound As String

Public ScriptNum As Long

Public Connucted As Boolean

Public SpeechEditorCurrentNumber As Long
Public SpeechConvo1 As Long
Public SpeechConvo2 As Long
Public SpeechConvo3 As Long

Public ShopNum As Long

Public GoDebug As Long

Public MouseX As Long
Public MouseY As Long
                    
Sub Main()
Dim I As Long
Dim Ending As String

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
    
    ScreenMode = 0
    
    File(1).Name = "cmcs21.ocx"
    File(2).Name = "MSBIND.dll"
    File(3).Name = "msdxm.ocx"
    File(4).Name = "msscript.ocx"
    File(5).Name = "MSSTDFMN.dll"
    File(6).Name = "msstdfmt.dll"
    File(7).Name = "PaintX.dll"
    File(8).Name = "SHDOCVW.dll"
    File(9).Name = "dx7vb.dll"
    File(10).Name = "MSCOMCTL.ocx"
    File(11).Name = "TABCTL32.ocx"
    File(12).Name = "RICHTX32.ocx"
    File(13).Name = "RICHED32.dll"

    frmSendGetData.Visible = True
    Call SetStatus("Checking folders...")
    DoEvents
    
    ' Check if the maps directory is there, if its not make it
    If LCase$(Dir(App.Path & "\Maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\Maps")
    End If
    If UCase(Dir(App.Path & "\GFX", vbDirectory)) <> "GFX" Then
        Call MkDir(App.Path & "\GFX")
    End If
    If UCase(Dir(App.Path & "\GUI", vbDirectory)) <> "GUI" Then
        Call MkDir(App.Path & "\GUI")
    End If
    If UCase(Dir(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then
        Call MkDir(App.Path & "\Music")
    End If
    If UCase(Dir(App.Path & "\SFX", vbDirectory)) <> "SFX" Then
        Call MkDir(App.Path & "\SFX")
    End If
    If UCase(Dir(App.Path & "\Flashs", vbDirectory)) <> "FLASHS" Then
        Call MkDir(App.Path & "\Flashs")
    End If
    If UCase$(Dir(App.Path & "\DLL", vbDirectory)) <> "DLL" Then
        Call MkDir(App.Path & "\DLL")
    End If
    
    Dim Filename As String
    Filename = App.Path & "\config.ini"
    If FileExist("config.ini") Then
        frmMirage.chkbubblebar.Value = Val(GetVar(Filename, "CONFIG", "SpeechBubbles"))
        frmMirage.chkEmoSound.Value = Val(GetVar(Filename, "CONFIG", "EmoticonSound"))
        frmMirage.chknpcname.Value = Val(GetVar(Filename, "CONFIG", "NPCName"))
        frmMirage.chkplayername.Value = Val(GetVar(Filename, "CONFIG", "PlayerName"))
        frmMirage.chkplayerdamage.Value = Val(GetVar(Filename, "CONFIG", "NPCDamage"))
        frmMirage.chknpcdamage.Value = Val(GetVar(Filename, "CONFIG", "PlayerDamage"))
        frmMirage.chkmusic.Value = Val(GetVar(Filename, "CONFIG", "Music"))
        frmMirage.chkSound.Value = Val(GetVar(Filename, "CONFIG", "Sound"))
        frmMirage.chkAutoScroll.Value = Val(GetVar(Filename, "CONFIG", "AutoScroll"))

        If Val(GetVar(Filename, "CONFIG", "MapGrid")) = 0 Then
            frmMapEditor.chkGrid.Value = 0
        Else
            frmMapEditor.chkGrid.Value = 1
        End If
    Else
        Call PutVar(App.Path & "\config.ini", "UPDATER", "FileName", "Diamond.exe")
        Call PutVar(App.Path & "\config.ini", "UPDATER", "WebSite", vbNullString)
        Call PutVar(App.Path & "\config.ini", "IPCONFIG", "IP", "127.0.0.1")
        Call PutVar(App.Path & "\config.ini", "IPCONFIG", "PORT", 4000)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Account", vbNullString)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Password", vbNullString)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "WebSite", vbNullString)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "SpeechBubbles", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "EmoticonSound", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "NPCName", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "NPCDamage", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "PlayerName", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "PlayerDamage", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "MapGrid", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Music", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Sound", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "AutoScroll", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Registered", 1)
        Call PutVar(App.Path * "\config.ini", "CONFIG", "MusicVol", 100)
    End If
    
    MusicOn = Val(GetVar(App.Path & "\config.ini", "CONFIG", "Music"))
    SoundOn = Val(GetVar(App.Path & "\config.ini", "CONFIG", "Sound"))
    MapGridOn = Val(GetVar(App.Path & "\config.ini", "CONFIG", "MapGrid"))
    NPCDamageOn = Val(GetVar(App.Path & "\config.ini", "CONFIG", "NPCDamage"))
    PlayerNameOn = Val(GetVar(App.Path & "\config.ini", "CONFIG", "PlayerName"))
    PlayerDamageOn = Val(GetVar(App.Path & "\config.ini", "CONFIG", "PlayerDamage"))
    NPCNameOn = Val(GetVar(App.Path & "\config.ini", "CONFIG", "NPCName"))
    SpeechBubblesOn = Val(GetVar(App.Path & "\config.ini", "CONFIG", "SpeechBubbles"))
    EmoticonSoundOn = Val(GetVar(App.Path & "\config.ini", "CONFIG", "EmoticonSound"))
    MusicVolume = Val(GetVar(App.Path & "\config.ini", "CONFIG", "MusicVol"))
    
    If MusicVolume > 100 Then
        Call PutVar(App.Path & "\config.ini", "CONFIG", "MusicVol", 100)
        MusicVolume = 100
    End If
    
    Call InitDirectSM
    
    frmMirage.scrlMusicVol.Value = MusicVolume
    frmMirage.lblMusicVol.Caption = MusicVolume
    
    If FileExist("GUI\Colors.txt") = False Then
        Call PutVar(App.Path & "\GUI\Colors.txt", "CHATBOX", "R", 152)
        Call PutVar(App.Path & "\GUI\Colors.txt", "CHATBOX", "G", 146)
        Call PutVar(App.Path & "\GUI\Colors.txt", "CHATBOX", "B", 120)
        
        Call PutVar(App.Path & "\GUI\Colors.txt", "CHATTEXTBOX", "R", 152)
        Call PutVar(App.Path & "\GUI\Colors.txt", "CHATTEXTBOX", "G", 146)
        Call PutVar(App.Path & "\GUI\Colors.txt", "CHATTEXTBOX", "B", 120)
        
        Call PutVar(App.Path & "\GUI\Colors.txt", "BACKGROUND", "R", 152)
        Call PutVar(App.Path & "\GUI\Colors.txt", "BACKGROUND", "G", 146)
        Call PutVar(App.Path & "\GUI\Colors.txt", "BACKGROUND", "B", 120)
        
        Call PutVar(App.Path & "\GUI\Colors.txt", "SPELLLIST", "R", 152)
        Call PutVar(App.Path & "\GUI\Colors.txt", "SPELLLIST", "G", 146)
        Call PutVar(App.Path & "\GUI\Colors.txt", "SPELLLIST", "B", 120)

        Call PutVar(App.Path & "\GUI\Colors.txt", "WHOLIST", "R", 152)
        Call PutVar(App.Path & "\GUI\Colors.txt", "WHOLIST", "G", 146)
        Call PutVar(App.Path & "\GUI\Colors.txt", "WHOLIST", "B", 120)
        
        Call PutVar(App.Path & "\GUI\Colors.txt", "NEWCHAR", "R", 152)
        Call PutVar(App.Path & "\GUI\Colors.txt", "NEWCHAR", "G", 146)
        Call PutVar(App.Path & "\GUI\Colors.txt", "NEWCHAR", "B", 120)
    End If
    
    Dim R1 As Long, G1 As Long, B1 As Long
    R1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "CHATBOX", "R"))
    G1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "CHATBOX", "G"))
    B1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "CHATBOX", "B"))
    frmMirage.txtChat.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "CHATTEXTBOX", "R"))
    G1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "CHATTEXTBOX", "G"))
    B1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "CHATTEXTBOX", "B"))
    frmMirage.txtMyTextBox.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "BACKGROUND", "R"))
    G1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "BACKGROUND", "G"))
    B1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "BACKGROUND", "B"))
    frmMirage.Picture9.BackColor = RGB(R1, G1, B1)
    frmMirage.Picture8.BackColor = RGB(R1, G1, B1)
    frmMirage.picInv3.BackColor = RGB(R1, G1, B1)
    frmMirage.itmDesc.BackColor = RGB(R1, G1, B1)
    frmMirage.picWhosOnline.BackColor = RGB(R1, G1, B1)
    frmMirage.picGuildAdmin.BackColor = RGB(R1, G1, B1)
    frmMirage.picGuild.BackColor = RGB(R1, G1, B1)
    frmMirage.picEquip.BackColor = RGB(R1, G1, B1)
    frmMirage.picPlayerSpells.BackColor = RGB(R1, G1, B1)
    frmMirage.picOptions.BackColor = RGB(R1, G1, B1)
    
    frmMirage.chkbubblebar.BackColor = RGB(R1, G1, B1)
    frmMirage.chkEmoSound.BackColor = RGB(R1, G1, B1)
    frmMirage.chknpcname.BackColor = RGB(R1, G1, B1)
    frmMirage.chkplayername.BackColor = RGB(R1, G1, B1)
    frmMirage.chkplayerdamage.BackColor = RGB(R1, G1, B1)
    frmMirage.chknpcdamage.BackColor = RGB(R1, G1, B1)
    frmMirage.chkmusic.BackColor = RGB(R1, G1, B1)
    frmMirage.chkSound.BackColor = RGB(R1, G1, B1)
    frmMirage.chkAutoScroll.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "SPELLLIST", "R"))
    G1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "SPELLLIST", "G"))
    B1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "SPELLLIST", "B"))
    frmMirage.lstSpells.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "WHOLIST", "R"))
    G1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "WHOLIST", "G"))
    B1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "WHOLIST", "B"))
    frmMirage.lstOnline.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "FRIENDLIST", "R"))
    G1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "FRIENDLIST", "G"))
    B1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "FRIENDLIST", "B"))
    frmMirage.lstFriend.BackColor = RGB(R1, G1, B1)

    R1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "NEWCHAR", "R"))
    G1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "NEWCHAR", "G"))
    B1 = Val(GetVar(App.Path & "\GUI\Colors.txt", "NEWCHAR", "B"))
    frmNewChar.optMale.BackColor = RGB(R1, G1, B1)
    frmNewChar.optFemale.BackColor = RGB(R1, G1, B1)
    
    Call SetStatus("Checking file registration...")
    DoEvents

    If GetVar(App.Path & "\config.ini", "CONFIG", "Registered") = 1 Then
        For I = 1 To MAX_REG
            Call Register(File(I).Name)
        Next I
       
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Registered", 2)
    End If
    
    Call SetStatus("Checking status...")
    DoEvents
    
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InSpeechEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InEmoticonEditor = False
    InArrowEditor = False
    InSpawnEditor = False
    
    frmMirage.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
    frmSpriteChange.Picsprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
    
    Call SetStatus("Initializing TCP Settings...")
    DoEvents
    
    Call TcpInit
    frmMainMenu.Visible = True
    frmSendGetData.Visible = False
End Sub

Sub SetStatus(ByVal Caption As String)
    frmSendGetData.lblStatus.Caption = Caption
End Sub

Sub MenuState(ByVal State As Long)
    Connucted = True
    frmSendGetData.Visible = True
    Call SetStatus("Connecting to server...")
    Select Case State
        Case MENU_STATE_NEWACCOUNT
            frmNewAccount.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text)
            End If
            
        Case MENU_STATE_DELACCOUNT
            frmDeleteAccount.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending account deletion request ...")
                Call SendDelAccount(frmDeleteAccount.txtName.Text, frmDeleteAccount.txtPassword.Text)
            End If
        
        Case MENU_STATE_LOGIN
            frmLogin.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
            End If
        
        Case MENU_STATE_NEWCHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, getting available classes...")
                Call SendGetClasses
            End If
            
        Case MENU_STATE_ADDCHAR
            frmNewChar.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character addition data...")
                If frmNewChar.optMale.Value = True Then
                    Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbClass.ListIndex + 1, frmChars.lstChars.ListIndex + 1)
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbClass.ListIndex + 1, frmChars.lstChars.ListIndex + 1)
                End If
            End If
        
        Case MENU_STATE_DELCHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character deletion request...")
                Call SendDelChar(frmChars.lstChars.ListIndex + 1)
            End If
            
        Case MENU_STATE_USECHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending char data...")
                Call SendUseChar(frmChars.lstChars.ListIndex + 1)
            End If
    End Select

    If Not IsConnected And Connucted = True Then
        frmMainMenu.Visible = True
        frmSendGetData.Visible = False
        Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & WEBSITE, vbOKOnly, GAME_NAME)
    End If
End Sub
Sub GameInit()
    frmMirage.Visible = True
    frmSendGetData.Visible = False
    Call InitDirectX
End Sub

Sub GameLoop()
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim x As Long
Dim y As Long
Dim I As Long
Dim rec_back As RECT
    
    ' Set the focus
    frmMirage.picScreen.SetFocus
    
    ' Set font
    Call SetFont("Fixedsys", 18, 0, 0, 0, 0)
    ' Fixedsys's size can't be changed and bold doesn't seem to work
                
    ' Used for calculating fps
    TickFPS = GetTickCount
    FPS = 0
    
    Do While InGame
        Tick = GetTickCount
        
        ' Check to make sure they aren't trying to auto do anything
        If GetAsyncKeyState(VK_UP) >= 0 And DirUp = True Then DirUp = False
        If GetAsyncKeyState(VK_DOWN) >= 0 And DirDown = True Then DirDown = False
        If GetAsyncKeyState(VK_LEFT) >= 0 And DirLeft = True Then DirLeft = False
        If GetAsyncKeyState(VK_RIGHT) >= 0 And DirRight = True Then DirRight = False
        If GetAsyncKeyState(VK_CONTROL) >= 0 And ControlDown = True Then ControlDown = False
        If GetAsyncKeyState(VK_SHIFT) >= 0 And ShiftDown = True Then ShiftDown = False
        
        ' Check to make sure we are still connected
        If Not IsConnected Then InGame = False
        
        ' Check if we need to restore surfaces
        If NeedToRestoreSurfaces Then
            DD.RestoreAllSurfaces
            Call InitSurfaces
        End If
                
    If GettingMap = False Then
        If GetPlayerPOINTS(MyIndex) > 0 Then
            frmMirage.AddStr.Visible = True
            frmMirage.AddDef.Visible = True
            frmMirage.AddSpeed.Visible = True
            frmMirage.AddMagi.Visible = True
        Else
            frmMirage.AddStr.Visible = False
            frmMirage.AddDef.Visible = False
            frmMirage.AddSpeed.Visible = False
            frmMirage.AddMagi.Visible = False
        End If
        ' Visual Inventory
        Dim Q As Long
        Dim Qq As Long
        Dim IT As Long
               
        If GetTickCount > IT + 500 And frmMirage.picInv3.Visible = True Then
            For Q = 0 To MAX_INV - 1
                Qq = Player(MyIndex).Inv(Q + 1).Num
               
                If frmMirage.picInv(Q).Picture <> LoadPicture() Then
                    frmMirage.picInv(Q).Picture = LoadPicture()
                Else
                    If Qq = 0 Then
                        frmMirage.picInv(Q).Picture = LoadPicture()
                    Else
                        Call BitBlt(frmMirage.picInv(Q).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qq).Pic - Int(Item(Qq).Pic / 6) * 6) * PIC_X, Int(Item(Qq).Pic / 6) * PIC_Y, SRCCOPY)
                    End If
                End If
            Next Q
        End If
                        
        NewX = 10
        NewY = 7
                          
        NewPlayerY = Player(MyIndex).y - NewY
        NewPlayerX = Player(MyIndex).x - NewX
        
        NewX = NewX * PIC_X
        NewY = NewY * PIC_Y
        
        NewXOffset = Player(MyIndex).XOffset
        NewYOffset = Player(MyIndex).YOffset
        
        If Player(MyIndex).y - 7 < 1 Then
            NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = 0
            If Player(MyIndex).y = 7 And Player(MyIndex).Dir = DIR_UP Then
                NewPlayerY = Player(MyIndex).y - 7
                NewY = 7 * PIC_Y
                NewYOffset = Player(MyIndex).YOffset
            End If
        ElseIf Player(MyIndex).y + 8 > MAX_MAPY + 1 Then
            NewY = (Player(MyIndex).y - 16) * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = MAX_MAPY - 13
            If Player(MyIndex).y = 24 And Player(MyIndex).Dir = DIR_DOWN Then
                NewPlayerY = Player(MyIndex).y - 7
                NewY = 7 * PIC_Y
                NewYOffset = Player(MyIndex).YOffset
            End If
        End If
        
        If Player(MyIndex).x - 10 < 1 Then
            NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = 0
            If Player(MyIndex).x = 10 And Player(MyIndex).Dir = DIR_LEFT Then
                NewPlayerX = Player(MyIndex).x - 10
                NewX = 10 * PIC_X
                NewXOffset = Player(MyIndex).XOffset
            End If
        ElseIf Player(MyIndex).x + 11 > MAX_MAPX + 1 Then
            NewX = (Player(MyIndex).x - 11) * PIC_X + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = MAX_MAPX - 19
            If Player(MyIndex).x = 21 And Player(MyIndex).Dir = DIR_RIGHT Then
                NewPlayerX = Player(MyIndex).x - 10
                NewX = 10 * PIC_X
                NewXOffset = Player(MyIndex).XOffset
            End If
        End If
        
        sx = 32
        If MAX_MAPX = 19 Then
            NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = 0
            NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = 0
            sx = 0
        End If
        
        ' Blit out tiles layers ground/anim1/anim2
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltTile(x, y)
            Next x
        Next y
       
    If ScreenMode = 0 Then
        ' Blit out the items
        For I = 1 To MAX_MAP_ITEMS
            If MapItem(I).Num > 0 Then
                Call BltItem(I)
            End If
        Next I
        
        ' Blit out NPC hp bars
        For I = 1 To MAX_MAP_NPCS
            If GetTickCount < MapNpc(I).LastAttack + 5000 Then
                Call BltNpcBars(I)
            End If
        Next I
              
        ' Blit players bar
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) Then
                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                    If GetTickCount < Player(MyIndex).LastAttack + 5000 Then
                        Call BltPlayerBars(I)
                    End If
                End If
                If Player(I).Pet.Map = GetPlayerMap(MyIndex) And Player(I).Pet.Alive = YES Then
                    If GetTickCount < Player(MyIndex).Pet.LastAttack + 5000 Then
                        Call BltPetBars(I)
                    End If
                End If
            End If
        Next I
        
        ' Blit out the sprite change attribute
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltSpriteChange(x, y)
            Next x
        Next y
        
        ' Blit out arrows
        Call BltArrow(GetPlayerMap(MyIndex))
        
        ' Blit out the npcs
        For I = 1 To MAX_MAP_NPCS
            Call BltNpc(I)
        Next I
        
        ' Blit out players
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) Then
                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                    If Player(I).Pet.Alive = YES Then
                        Call BltPet(I)
                    End If
                    Call BltPlayer(I)
                End If
            End If
        Next I
        
        If SIZE_Y > PIC_Y Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                        If Player(I).Pet.Alive = YES Then
                            Call BltPetTop(I)
                        End If
                        Call BltPlayerTop(I)
                    End If
                End If
            Next I
        End If
        
        ' Blit the spells
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                Call BltSpell(I)
            End If
        Next I
        
        ' Blit out the sprite change attribute
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltSpriteChange2(x, y)
            Next x
        Next y
        
        If SIZE_Y > PIC_Y Then
            ' Blit out the npc's top
            For I = 1 To MAX_MAP_NPCS
                Call BltNpcTop(I)
            Next I
        End If
    End If
    
    If ScreenMode = 0 Then
        ' Blit out the npcs
        For I = 1 To MAX_MAP_NPCS
            Call BltNpcTop(I)
        Next I
    End If
                
    ' Blit out tile layer fringe
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Call BltFringeTile(x, y)
        Next x
    Next y
      
    'If ScreenMode = 0 Then
        ' Blit out the npcs
    '    For i = 1 To MAX_MAP_NPCS
    '        If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).y).Fringe < 1 Then
    '            If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).y).FAnim < 1 Then
    '                If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).y).Fringe2 < 1 Then
    '                    If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).y).F2Anim < 1 Then
    '                        Call BltNpcTop(i)
    '                    End If
    '                End If
    '            End If
    '        End If
    '    Next i
    'End If
    
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) = True Then
            If Player(I).LevelUpT + 3000 > GetTickCount Then
                rec.Top = Int(32 / TilesInSheets) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = (32 - Int(32 / TilesInSheets) * TilesInSheets) * PIC_X
                rec.Right = rec.Left + 96
                
                If I = MyIndex Then
                    x = NewX + sx
                    y = NewY + sx
                    Call DD_BackBuffer.BltFast(x - 32, y - 10 - Player(I).LevelUp, DD_TileSurf(6), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    x = GetPlayerX(I) * PIC_X + sx + Player(I).XOffset
                    y = GetPlayerY(I) * PIC_Y + sx + Player(I).YOffset
                    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - 32 - NewXOffset, y - (NewPlayerY * PIC_Y) - 10 - Player(I).LevelUp - NewYOffset, DD_TileSurf(6), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                If Player(I).LevelUp >= 3 Then
                    Player(I).LevelUp = Player(I).LevelUp - 1
                ElseIf Player(I).LevelUp >= 1 Then
                    Player(I).LevelUp = Player(I).LevelUp + 1
                End If
            Else
                Player(I).LevelUpT = 0
            End If
        End If
    Next I
    
    If GettingMap = False Then
        If GameTime = TIME_NIGHT And Map(GetPlayerMap(MyIndex)).Indoors = 0 And InEditor = False Then
            Call Night
        End If
        If frmMapEditor.chkDayNight.Value = 1 And InEditor = True Then
            Call Night
        End If
        If Map(GetPlayerMap(MyIndex)).Indoors = 0 Then Call BltWeather
    End If

    If InEditor = True And MapGridOn = 1 Then
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltTile2(x * 32, y * 32, 0)
            Next x
        Next y
    End If
End If
    
    ' Lock the backbuffer so we can draw text and names
    TexthDC = DD_BackBuffer.GetDC
    If GettingMap = False Then
        If ScreenMode = 0 Then
            If NPCDamageOn = YES Then
                If PlayerNameOn = NO Then
                    If GetTickCount < NPCDmgTime + 2000 Then
                        Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 22 - ii + sx, NPCDmgDamage, QBColor(BrightRed))
                    End If
                Else
                    If GetPlayerGuild(MyIndex) <> vbNullString Then
                        If GetTickCount < NPCDmgTime + 2000 Then
                            Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 42 - ii + sx, NPCDmgDamage, QBColor(BrightRed))
                        End If
                    Else
                        If GetTickCount < NPCDmgTime + 2000 Then
                            Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 22 - ii + sx, NPCDmgDamage, QBColor(BrightRed))
                        End If
                    End If
                End If
                ii = ii + 1
            End If
            
            If PlayerDamageOn = YES Then
                If NPCWho > 0 Then
                    If MapNpc(NPCWho).Num > 0 Then
                        If NPCNameOn = NO Then
                            If Npc(MapNpc(NPCWho).Num).Big = 0 Then
                                If GetTickCount < DmgTime + 2000 Then
                                    Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 20 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                                End If
                            Else
                                If GetTickCount < DmgTime + 2000 Then
                                    Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 47 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                                End If
                            End If
                        Else
                            If Npc(MapNpc(NPCWho).Num).Big = 0 Then
                                If GetTickCount < DmgTime + 2000 Then
                                    Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 30 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                                End If
                            Else
                                If GetTickCount < DmgTime + 2000 Then
                                    Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 57 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                                End If
                            End If
                        End If
                        iii = iii + 1
                    End If
                End If
            End If
            
            'Draw NPC Names
            If NPCNameOn = YES Then
                For I = LBound(MapNpc) To UBound(MapNpc)
                    If MapNpc(I).Num > 0 Then
                        Call BltMapNPCName(I)
                    End If
                Next I
            End If
            
            ' Draw Player Names
            If PlayerNameOn = YES Then
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                        Call BltPlayerGuildName(I)
                        Call BltPlayerName(I)
                        If Player(I).Pet.Alive = YES And Player(I).Pet.Map = GetPlayerMap(MyIndex) Then
                            Call BltPetName(I)
                        End If
                    End If
                Next I
            End If
     
            ' speech bubble stuffs
            If SpeechBubblesOn = YES Then
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                        If Bubble(I).Text <> vbNullString Then
                            Call BltPlayerText(I)
                        End If
        
                        If GetTickCount() > Bubble(I).Created + DISPLAY_BUBBLE_TIME Then
                            Bubble(I).Text = vbNullString
                        End If
                    End If
                Next I
            End If
    
            ''Draw NPC Names
            'If NPCNameOn = YES Then
            '    For i = LBound(MapNpc) To UBound(MapNpc)
            '        If MapNpc(i).Num > 0 Then
            '            Call BltMapNPCName(i)
            '        End If
            '    Next i
            'End If
            
            ' Blit out attribs if in editor
            If InEditor Then
                For y = 0 To MAX_MAPY
                    For x = 0 To MAX_MAPX
                        With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                            If .Type = TILE_TYPE_BLOCKED Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "B", QBColor(BrightRed))
                            If .Type = TILE_TYPE_WARP Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "W", QBColor(BrightBlue))
                            If .Type = TILE_TYPE_ITEM Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "I", QBColor(White))
                            If .Type = TILE_TYPE_NPCAVOID Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "N", QBColor(White))
                            If .Type = TILE_TYPE_KEY Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "K", QBColor(White))
                            If .Type = TILE_TYPE_KEYOPEN Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "O", QBColor(White))
                            If .Type = TILE_TYPE_HEAL Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "H", QBColor(BrightGreen))
                            If .Type = TILE_TYPE_KILL Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "K", QBColor(BrightRed))
                            If .Type = TILE_TYPE_SHOP Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "S", QBColor(Yellow))
                            If .Type = TILE_TYPE_CBLOCK Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "CB", QBColor(Black))
                            If .Type = TILE_TYPE_ARENA Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "A", QBColor(BrightGreen))
                            If .Type = TILE_TYPE_SOUND Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "PS", QBColor(Yellow))
                            If .Type = TILE_TYPE_SPRITE_CHANGE Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SC", QBColor(Grey))
                            If .Type = TILE_TYPE_SIGN Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SI", QBColor(Yellow))
                            If .Type = TILE_TYPE_DOOR Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "D", QBColor(Black))
                            If .Type = TILE_TYPE_NOTICE Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "N", QBColor(BrightGreen))
                            If .Type = TILE_TYPE_CHEST Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "C", QBColor(Brown))
                            If .Type = TILE_TYPE_CLASS_CHANGE Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "CG", QBColor(White))
                            If .Type = TILE_TYPE_SCRIPTED Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SC", QBColor(Yellow))
                            If .Light > 0 Then Call DrawText(TexthDC, x * PIC_X + sx + 18 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 14 - (NewPlayerY * PIC_Y) - NewYOffset, "L", QBColor(Yellow))
                        End With
                        
                        If InSpawnEditor Then
                            For I = 1 To MAX_MAP_NPCS
                                If TempNpcSpawn(I).Used = YES Then
                                    If x = TempNpcSpawn(I).x And y = TempNpcSpawn(I).y Then
                                        Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, I, QBColor(White))
                                    End If
                                End If
                            Next I
                        End If
                    Next x
                Next y
            End If
            
            ' Blit the text they are putting in
            'MyText = frmMirage.txtMyTextBox.Text
            'frmMirage.txtMyTextBox.Text = MyText
            
            'If Len(MyText) > 4 Then
                'frmMirage.txtMyTextBox.SelStart = Len(frmMirage.txtMyTextBox.Text) + 1
            'End If
                    
            ' Draw map name
            If Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NONE Then
                Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).Name), QBColor(BrightRed))
            ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_SAFE Then
                Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).Name), QBColor(White))
            ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NO_PENALTY Then
                Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).Name), QBColor(Black))
            End If
            
            ' Battle messages
            For I = 1 To MAX_BLT_LINE
                If BattlePMsg(I).Index > 0 Then
                    If BattlePMsg(I).Time + 7000 > GetTickCount Then
                        Call DrawText(TexthDC, 1 + sx, BattlePMsg(I).y + frmMirage.picScreen.Height - 15 + sx, Trim$(BattlePMsg(I).Msg), QBColor(BattlePMsg(I).Color))
                    Else
                        BattlePMsg(I).done = 0
                    End If
                End If
                
                If BattleMMsg(I).Index > 0 Then
                    If BattleMMsg(I).Time + 7000 > GetTickCount Then
                        Call DrawText(TexthDC, (frmMirage.picScreen.Width - (Len(BattleMMsg(I).Msg) * 8)) + sx, BattleMMsg(I).y + frmMirage.picScreen.Height - 15 + sx, Trim$(BattleMMsg(I).Msg), QBColor(BattleMMsg(I).Color))
                    Else
                        BattleMMsg(I).done = 0
                    End If
                End If
            Next I
        End If
    End If

        ' Check if we are getting a map, and if we are tell them so
        If GettingMap = True Then
            Call DrawText(TexthDC, 36, 36, "Receiving map...", QBColor(BrightCyan))
        End If
                        
        ' Release DC
        Call DD_BackBuffer.ReleaseDC(TexthDC)
        
        ' Blit out emoticons
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                Call BltEmoticons(I)
            End If
        Next I
        
        ' Get the rect for the back buffer to blit from
        rec.Top = 0
        rec.Bottom = (MAX_MAPY + 1) * PIC_Y
        rec.Left = 0
        rec.Right = (MAX_MAPX + 1) * PIC_X
        
        ' Get the rect to blit to
        Call DX.GetWindowRect(frmMirage.picScreen.hWnd, rec_pos)
        rec_pos.Bottom = rec_pos.Top - sx + ((MAX_MAPY + 1) * PIC_Y)
        rec_pos.Right = rec_pos.Left - sx + ((MAX_MAPX + 1) * PIC_X)
        rec_pos.Top = rec_pos.Bottom - ((MAX_MAPY + 1) * PIC_Y)
        rec_pos.Left = rec_pos.Right - ((MAX_MAPX + 1) * PIC_X)
        
        ' Blit the backbuffer
        Call DD_PrimarySurf.Blt(rec_pos, DD_BackBuffer, rec, DDBLT_WAIT)
        
        ' Check if player is trying to move
        Call CheckMovement
        
        ' Check to see if player is trying to attack
        Call CheckAttack
        
        ' Process player and pet movements (actually move them)
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) Then
                Call ProcessMovement(I)
                If Player(I).Pet.Alive = YES Then
                    Call ProcessPetMovement(I)
                End If
            End If
        Next I
        
        If InEditor Then
            frmMirage.shpSelect.Left = MouseX
            frmMirage.shpSelect.Top = MouseY
        End If
        
        ' Process npc movements (actually move them)
        For I = 1 To MAX_MAP_NPCS
            If Map(GetPlayerMap(MyIndex)).Npc(I) > 0 Then
                Call ProcessNpcMovement(I)
            End If
        Next I
  
        ' Change map animation every 250 milliseconds
        If GetTickCount > MapAnimTimer + 250 Then
            If MapAnim = 0 Then
                MapAnim = 1
            Else
                MapAnim = 0
            End If
            MapAnimTimer = GetTickCount
        End If
                
        ' Lock fps
        Do While GetTickCount < Tick + 35
            DoEvents
            Sleep 1
        Loop
        
        ' Calculate fps
        If GetTickCount > TickFPS + 1000 Then
            GameFPS = FPS
            TickFPS = GetTickCount
            FPS = 0
        Else
            FPS = FPS + 1
        End If
        
        Call MakeMidiLoop
        
        DoEvents
    Loop
    
    frmMirage.Visible = False
    frmSendGetData.Visible = True
    Call SetStatus("Destroying game data...")
    
    ' Shutdown the game
    Call GameDestroy
    
    ' Report disconnection if server disconnects
    If IsConnected = False Then
        Call MsgBox("Thank you for playing " & GAME_NAME & "!", vbOKOnly, GAME_NAME)
    End If
End Sub

Sub GameDestroy()

    'Tell the server, so it doesn't have to check so often
    'XX This causes errors I believe XX
    'Call SendData(CLOSINGDOWN_CHAR & SEP_CHAR & MyIndex & END_CHAR)

    Call DestroyDirectX
    Call StopMidi
    Call closemedown(True)
    End
End Sub

Public Sub closemedown(Optional ByVal ForceClose As Boolean = False)
Dim I As Long
On Error Resume Next

    For I = Forms.Count - 1 To 0 Step -1
        Unload Forms(I)
        Set Forms(I) = Nothing
            
        If Not ForceClose Then
            If Forms.Count > I Then
                Exit Sub
            End If
        End If
    Next I

    If ForceClose Or (Forms.Count = 0) Then Close
    If ForceClose Or (Forms.Count > 0) Then End
End Sub

Sub BltTile(ByVal x As Long, ByVal y As Long)
Dim Ground As Long
Dim Mask As Long
Dim Anim As Long
Dim Mask2 As Long
Dim M2Anim As Long
Dim GroundTileSet As Long
Dim MaskTileSet As Long
Dim AnimTileSet As Long
Dim Mask2TileSet As Long
Dim M2AnimTileSet As Long

    Ground = Map(GetPlayerMap(MyIndex)).Tile(x, y).Ground
    Mask = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask
    Anim = Map(GetPlayerMap(MyIndex)).Tile(x, y).Anim
    Mask2 = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2
    M2Anim = Map(GetPlayerMap(MyIndex)).Tile(x, y).M2Anim
    
    If TempTile(x, y).Ground <> 0 Then Ground = TempTile(x, y).Ground
    If TempTile(x, y).Mask <> 0 Then Mask = TempTile(x, y).Mask
    If TempTile(x, y).Anim <> 0 Then Anim = TempTile(x, y).Anim
    If TempTile(x, y).Mask2 <> 0 Then Mask2 = TempTile(x, y).Mask2
    If TempTile(x, y).M2Anim <> 0 Then M2Anim = TempTile(x, y).M2Anim
    
    GroundTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).GroundSet
    MaskTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).MaskSet
    AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).AnimSet
    Mask2TileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2Set
    M2AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).M2AnimSet
    
    ' If (GroundTileSet >= 0 And TileFile(GroundTileSet) = 0) Or (MaskTileSet >= 0 And TileFile(MaskTileSet) = 0) Or (AnimTileSet >= 0 And TileFile(AnimTileSet) = 0) Or (Mask2TileSet >= 0 And TileFile(Mask2TileSet) = 0) Or (M2AnimTileSet >= 0 And TileFile(M2AnimTileSet) = 0) Then Exit Sub
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = (y - NewPlayerY) * PIC_Y + sx - NewYOffset
        .Bottom = .Top + PIC_Y
        .Left = (x - NewPlayerX) * PIC_X + sx - NewXOffset
        .Right = .Left + PIC_X
    End With
    
    If GroundTileSet < 0 Then
        GroundTileSet = 0
        Ground = 0
    End If
    
    rec.Top = Int(Ground / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Ground - Int(Ground / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(GroundTileSet), rec, DDBLT_WAIT)
    'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT)
    
    ' Is there an animation tile to plot?
    If (MapAnim = 0) Or (Anim <= 0) Then
        If Mask > 0 And MaskTileSet >= 0 And TempTile(x, y).DoorOpen = NO Then
            rec.Top = Int(Mask / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask - Int(Mask / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(MaskTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If Anim > 0 And AnimTileSet >= 0 Then
            rec.Top = Int(Anim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim - Int(Anim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    ' Is there a second animation tile to plot?
    If (MapAnim = 0) Or (M2Anim <= 0) Then
        If Mask2 > 0 And Mask2TileSet >= 0 Then
            rec.Top = Int(Mask2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask2 - Int(Mask2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(Mask2TileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If M2Anim > 0 And M2AnimTileSet >= 0 Then
            rec.Top = Int(M2Anim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (M2Anim - Int(M2Anim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(M2AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltItem(ByVal ItemNum As Long)
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapItem(ItemNum).y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = MapItem(ItemNum).x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    rec.Top = Int(Item(MapItem(ItemNum).Num).Pic / 6) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Item(MapItem(ItemNum).Num).Pic - Int(Item(MapItem(ItemNum).Num).Pic / 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_ItemSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast((MapItem(ItemNum).x - NewPlayerX) * PIC_X + sx - NewXOffset, (MapItem(ItemNum).y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFringeTile(ByVal x As Long, ByVal y As Long)
Dim Fringe As Long
Dim FAnim As Long
Dim Fringe2 As Long
Dim F2Anim As Long
Dim FringeTileSet As Long
Dim FAnimTileSet As Long
Dim Fringe2TileSet As Long
Dim F2AnimTileSet As Long

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    Fringe = Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe
    FAnim = Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnim
    Fringe2 = Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2
    F2Anim = Map(GetPlayerMap(MyIndex)).Tile(x, y).F2Anim
    
    If TempTile(x, y).Fringe <> 0 Then Fringe = TempTile(x, y).Fringe
    If TempTile(x, y).FAnim <> 0 Then FAnim = TempTile(x, y).FAnim
    If TempTile(x, y).Fringe2 <> 0 Then Fringe2 = TempTile(x, y).Fringe2
    If TempTile(x, y).F2Anim <> 0 Then F2Anim = TempTile(x, y).F2Anim
    
    FringeTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).FringeSet
    FAnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnimSet
    Fringe2TileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2Set
    F2AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).F2AnimSet
        
    ' If (FringeTileSet >= 0 And TileFile(FringeTileSet) = 0) Or (FAnimTileSet >= 0 And TileFile(FAnimTileSet) = 0) Or (Fringe2TileSet >= 0 And TileFile(Fringe2TileSet) = 0) Or (F2AnimTileSet >= 0 And TileFile(F2AnimTileSet) = 0) Then Exit Sub
        
    ' Is there an animation tile to plot?
    If (MapAnim = 0) Or (FAnim <= 0) Then
        If Fringe > 0 And FringeTileSet >= 0 Then
            rec.Top = Int(Fringe / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe - Int(Fringe / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FringeTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    
    Else
        If FAnim > 0 And FAnimTileSet >= 0 Then
            rec.Top = Int(FAnim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (FAnim - Int(FAnim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FAnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
    End If

    ' Is there a second animation tile to plot?
    If (MapAnim = 0) Or (F2Anim <= 0) Then
        If Fringe2 > 0 And Fringe2TileSet >= 0 Then
            rec.Top = Int(Fringe2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe2 - Int(Fringe2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(Fringe2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F2Anim And F2AnimTileSet >= 0 > 0 Then
            rec.Top = Int(F2Anim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (F2Anim - Int(F2Anim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(F2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(Index) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (SIZE_Y - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + ((SIZE_X - PIC_X) / 2)
        .Right = .Left + PIC_X + ((SIZE_X - PIC_X) / 2)
    End With
    
    ' Check for animation
    Anim = 0
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).YOffset > PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).XOffset > PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
        Player(Index).Attacking = 0
        Player(Index).AttackTimer = 0
    End If
    
    rec.Top = GetPlayerSprite(Index) * SIZE_Y + (SIZE_Y - PIC_Y)
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (GetPlayerDir(Index) * (3 * (SIZE_X / PIC_X)) + (Anim * (SIZE_X / PIC_X))) * PIC_X
    rec.Right = rec.Left + SIZE_X

    x = GetPlayerX(Index) * PIC_X - (SIZE_X - PIC_X) / 2 + sx + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset + (SIZE_Y - PIC_Y)
    
    If SIZE_X > PIC_X Then
        If x < 0 Then
            x = Player(Index).XOffset + sx + ((SIZE_X - PIC_X) / 2)
            If GetPlayerDir(Index) = DIR_RIGHT And Player(Index).Moving > 0 Then
                rec.Left = rec.Left - Player(Index).XOffset
            Else
                rec.Left = rec.Left - Player(Index).XOffset + ((SIZE_X - PIC_X) / 2)
            End If
        End If
        
        If x > MAX_MAPX * 32 Then
            x = MAX_MAPX * 32 + sx - ((SIZE_X - PIC_X) / 2) + Player(Index).XOffset
            If GetPlayerDir(Index) = DIR_LEFT And Player(Index).Moving > 0 Then
                rec.Right = rec.Right + Player(Index).XOffset
            Else
                rec.Right = rec.Right + Player(Index).XOffset - ((SIZE_X - PIC_X) / 2)
            End If
        End If
    End If
    
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerTop(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(Index) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (SIZE_Y - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + ((SIZE_X - PIC_X) / 2)
        .Right = .Left + PIC_X + ((SIZE_X - PIC_X) / 2)
    End With
    
    ' Check for animation
    Anim = 0
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
        Player(Index).Attacking = 0
        Player(Index).AttackTimer = 0
    End If
    
    rec.Top = GetPlayerSprite(Index) * SIZE_Y
    rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
    rec.Left = (GetPlayerDir(Index) * (3 * (SIZE_X / PIC_X)) + (Anim * (SIZE_X / PIC_X))) * PIC_X
    rec.Right = rec.Left + SIZE_X

    x = GetPlayerX(Index) * PIC_X - (SIZE_X - PIC_X) / 2 + sx + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset
    
    
    If y < 0 Then
        y = 0
        If GetPlayerDir(Index) = DIR_DOWN And Player(Index).Moving > 0 Then
            rec.Top = rec.Top - Player(Index).YOffset
        Else
            rec.Top = rec.Top - Player(Index).YOffset + (SIZE_Y - PIC_Y)
        End If
    End If
    
    If SIZE_X > PIC_X Then
        If x < 0 Then
            x = Player(Index).XOffset + sx + ((SIZE_X - PIC_X) / 2)
            If GetPlayerDir(Index) = DIR_RIGHT And Player(Index).Moving > 0 Then
                rec.Left = rec.Left - Player(Index).XOffset
            Else
                rec.Left = rec.Left - Player(Index).XOffset + ((SIZE_X - PIC_X) / 2)
            End If
        End If
        
        If x > MAX_MAPX * 32 Then
            x = MAX_MAPX * 32 + sx - ((SIZE_X - PIC_X) / 2) + Player(Index).XOffset
            If GetPlayerDir(Index) = DIR_LEFT And Player(Index).Moving > 0 Then
                rec.Right = rec.Right + Player(Index).XOffset
            Else
                rec.Right = rec.Right + Player(Index).XOffset - ((SIZE_X - PIC_X) / 2)
            End If
        End If
    End If
    
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPet(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = Player(Index).Pet.y * PIC_Y + Player(Index).Pet.YOffset - (SIZE_Y - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = Player(Index).Pet.x * PIC_X + Player(Index).Pet.XOffset + ((SIZE_X - PIC_X) / 2)
        .Right = .Left + PIC_X + ((SIZE_X - PIC_X) / 2)
    End With
    
    ' Check for animation
    Anim = 0
    If Player(Index).Pet.Attacking = 0 Then
        Select Case Player(Index).Pet.Dir
            Case DIR_UP
                If (Player(Index).Pet.YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).Pet.YOffset > PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).Pet.XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).Pet.XOffset > PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).Pet.AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If Player(Index).Pet.AttackTimer + 1000 < GetTickCount Then
        Player(Index).Pet.Attacking = 0
        Player(Index).Pet.AttackTimer = 0
    End If
    
    rec.Top = Player(Index).Pet.Sprite * SIZE_Y + (SIZE_Y - PIC_Y)
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Player(Index).Pet.Dir * (3 * (SIZE_X / PIC_X)) + (Anim * (SIZE_X / PIC_X))) * PIC_X
    rec.Right = rec.Left + SIZE_X

    x = Player(Index).Pet.x * PIC_X - (SIZE_X - PIC_X) / 2 + sx + Player(Index).Pet.XOffset
    y = Player(Index).Pet.y * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).Pet.YOffset + (SIZE_Y - PIC_Y)
    
    If SIZE_X > PIC_X Then
        If x < 0 Then
            x = Player(Index).Pet.XOffset + sx + ((SIZE_X - PIC_X) / 2)
            If Player(Index).Pet.Dir = DIR_RIGHT And Player(Index).Pet.Moving > 0 Then
                rec.Left = rec.Left - Player(Index).Pet.XOffset
            Else
                rec.Left = rec.Left - Player(Index).Pet.XOffset + ((SIZE_X - PIC_X) / 2)
            End If
        End If
        
        If x > MAX_MAPX * 32 Then
            x = MAX_MAPX * 32 + sx - ((SIZE_X - PIC_X) / 2) + Player(Index).Pet.XOffset
            If Player(Index).Pet.Dir = DIR_LEFT And Player(Index).Pet.Moving > 0 Then
                rec.Right = rec.Right + Player(Index).Pet.XOffset
            Else
                rec.Right = rec.Right + Player(Index).Pet.XOffset - ((SIZE_X - PIC_X) / 2)
            End If
        End If
    End If
    
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPetTop(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = Player(Index).Pet.y * PIC_Y + Player(Index).Pet.YOffset - (SIZE_Y - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = Player(Index).Pet.x * PIC_X + Player(Index).Pet.XOffset + ((SIZE_X - PIC_X) / 2)
        .Right = .Left + PIC_X + ((SIZE_X - PIC_X) / 2)
    End With
    
    ' Check for animation
    Anim = 0
    If Player(Index).Pet.Attacking = 0 Then
        Select Case Player(Index).Pet.Dir
            Case DIR_UP
                If (Player(Index).Pet.YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).Pet.YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).Pet.XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).Pet.XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).Pet.AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If Player(Index).Pet.AttackTimer + 1000 < GetTickCount Then
        Player(Index).Pet.Attacking = 0
        Player(Index).Pet.AttackTimer = 0
    End If
    
    rec.Top = Player(Index).Pet.Sprite * SIZE_Y
    rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
    rec.Left = (Player(Index).Pet.Dir * (3 * (SIZE_X / PIC_X)) + (Anim * (SIZE_X / PIC_X))) * PIC_X
    rec.Right = rec.Left + SIZE_X

    x = Player(Index).Pet.x * PIC_X - (SIZE_X - PIC_X) / 2 + sx + Player(Index).Pet.XOffset
    y = Player(Index).Pet.y * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).Pet.YOffset
    
    
    If y < 0 Then
        y = 0
        If Player(Index).Pet.Dir = DIR_DOWN And Player(Index).Pet.Moving > 0 Then
            rec.Top = rec.Top - Player(Index).Pet.YOffset
        Else
            rec.Top = rec.Top - Player(Index).Pet.YOffset + (SIZE_Y - PIC_Y)
        End If
    End If
    
    If SIZE_X > PIC_X Then
        If x < 0 Then
            x = Player(Index).Pet.XOffset + sx + ((SIZE_X - PIC_X) / 2)
            If Player(Index).Pet.Dir = DIR_RIGHT And Player(Index).Pet.Moving > 0 Then
                rec.Left = rec.Left - Player(Index).Pet.XOffset
            Else
                rec.Left = rec.Left - Player(Index).Pet.XOffset + ((SIZE_X - PIC_X) / 2)
            End If
        End If
        
        If x > MAX_MAPX * 32 Then
            x = MAX_MAPX * 32 + sx - ((SIZE_X - PIC_X) / 2) + Player(Index).Pet.XOffset
            If Player(Index).Pet.Dir = DIR_LEFT And Player(Index).Pet.Moving > 0 Then
                rec.Right = rec.Right + Player(Index).Pet.XOffset
            Else
                rec.Right = rec.Right + Player(Index).Pet.XOffset - ((SIZE_X - PIC_X) / 2)
            End If
        End If
    End If
    
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltMapNPCName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long

If Npc(MapNpc(Index).Num).Big = 0 Then
    With Npc(MapNpc(Index).Num)
    'Draw name
        TextX = MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
        TextY = MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset - CLng(PIC_Y / 2) - 4 - (SIZE_Y - PIC_Y)
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
    End With
Else
    With Npc(MapNpc(Index).Num)
    'Draw name
        TextX = MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
        TextY = MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset - CLng(PIC_Y / 2) - 32
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
    End With
End If
End Sub

Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long, y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - (SIZE_Y - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).XOffset + ((SIZE_X - PIC_X) / 2)
        .Right = .Left + PIC_X + ((SIZE_X - PIC_X) / 2)
    End With
    
    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset > PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset > PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then
        MapNpc(MapNpcNum).Attacking = 0
        MapNpc(MapNpcNum).AttackTimer = 0
    End If
    
    If Npc(MapNpc(MapNpcNum).Num).Big = 0 Then
    
    rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * SIZE_Y + (SIZE_Y - PIC_Y)
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * (3 * (SIZE_X / PIC_X)) + (Anim * (SIZE_X / PIC_X))) * PIC_X
    rec.Right = rec.Left + SIZE_X

    x = MapNpc(MapNpcNum).x * PIC_X - (SIZE_X - PIC_X) / 2 + sx + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y - (SIZE_Y - PIC_Y) + sx + MapNpc(MapNpcNum).YOffset + (SIZE_Y - PIC_Y)
    
    If SIZE_X > PIC_X Then
        If x < 0 Then
            x = MapNpc(MapNpcNum).XOffset + sx + ((SIZE_X - PIC_X) / 2)
            If MapNpc(MapNpcNum).Dir = DIR_RIGHT And MapNpc(MapNpcNum).Moving > 0 Then
                rec.Left = rec.Left - MapNpc(MapNpcNum).XOffset
            Else
                rec.Left = rec.Left - MapNpc(MapNpcNum).XOffset + ((SIZE_X - PIC_X) / 2)
            End If
        End If
        
        If x > MAX_MAPX * 32 Then
            x = MAX_MAPX * 32 + sx - ((SIZE_X - PIC_X) / 2) + MapNpc(MapNpcNum).XOffset
            If MapNpc(MapNpcNum).Dir = DIR_LEFT And MapNpc(MapNpcNum).Moving > 0 Then
                rec.Right = rec.Right + MapNpc(MapNpcNum).XOffset
            Else
                rec.Right = rec.Right + MapNpc(MapNpcNum).XOffset - ((SIZE_X - PIC_X) / 2)
            End If
        End If
    End If
    
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    Else
        rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + (64 - PIC_Y)
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (MapNpc(MapNpcNum).Dir * (3 * (64 / PIC_X)) + (Anim * (64 / PIC_X))) * PIC_X
        rec.Right = rec.Left + 64
       
        x = MapNpc(MapNpcNum).x * PIC_X - (64 - PIC_X) / 2 + sx + MapNpc(MapNpcNum).XOffset
        y = MapNpc(MapNpcNum).y * PIC_Y - (64 - PIC_Y) + sx + MapNpc(MapNpcNum).YOffset + (64 - PIC_Y)
       
        If 64 > PIC_X Then
            If x < 0 Then
                x = MapNpc(MapNpcNum).XOffset + sx + ((64 - PIC_X) / 2)
                If MapNpc(MapNpcNum).Dir = DIR_RIGHT And MapNpc(MapNpcNum).Moving > 0 Then
                    rec.Left = rec.Left - MapNpc(MapNpcNum).XOffset
                Else
                    rec.Left = rec.Left - MapNpc(MapNpcNum).XOffset + ((64 - PIC_X) / 2)
                End If
            End If
           
            If x > MAX_MAPX * 42 Then
                x = MAX_MAPX * 42 + sx - ((64 - PIC_X) / 2) + MapNpc(MapNpcNum).XOffset
                If MapNpc(MapNpcNum).Dir = DIR_LEFT And MapNpc(MapNpcNum).Moving > 0 Then
                    rec.Right = rec.Right + MapNpc(MapNpcNum).XOffset
                Else
                    rec.Right = rec.Right + MapNpc(MapNpcNum).XOffset - ((64 - PIC_X) / 2)
                End If
            End If
        End If
       
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltNpcTop(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long, y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - (SIZE_Y - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).XOffset + ((SIZE_X - PIC_X) / 2)
        .Right = .Left + PIC_X + ((SIZE_X - PIC_X) / 2)
    End With
    
    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then
        MapNpc(MapNpcNum).Attacking = 0
        MapNpc(MapNpcNum).AttackTimer = 0
    End If
    
    If Npc(MapNpc(MapNpcNum).Num).Big = 0 Then
    
    rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * SIZE_Y
    rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
    rec.Left = (MapNpc(MapNpcNum).Dir * (3 * (SIZE_X / PIC_X)) + (Anim * (SIZE_X / PIC_X))) * PIC_X
    rec.Right = rec.Left + SIZE_X

    x = MapNpc(MapNpcNum).x * PIC_X - (SIZE_X - PIC_X) / 2 + sx + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y - (SIZE_Y - PIC_Y) + sx + MapNpc(MapNpcNum).YOffset
    
    
    If y < 0 Then
        y = 0
        If MapNpc(MapNpcNum).Dir = DIR_DOWN And MapNpc(MapNpcNum).Moving > 0 Then
            rec.Top = rec.Top - MapNpc(MapNpcNum).YOffset
        Else
            rec.Top = rec.Top - MapNpc(MapNpcNum).YOffset + (SIZE_Y - PIC_Y)
        End If
    End If
    
    If SIZE_X > PIC_X Then
        If x < 0 Then
            x = MapNpc(MapNpcNum).XOffset + sx + ((SIZE_X - PIC_X) / 2)
            If MapNpc(MapNpcNum).Dir = DIR_RIGHT And MapNpc(MapNpcNum).Moving > 0 Then
                rec.Left = rec.Left - MapNpc(MapNpcNum).XOffset
            Else
                rec.Left = rec.Left - MapNpc(MapNpcNum).XOffset + ((SIZE_X - PIC_X) / 2)
            End If
        End If
        
        If x > MAX_MAPX * 32 Then
            x = MAX_MAPX * 32 + sx - ((SIZE_X - PIC_X) / 2) + MapNpc(MapNpcNum).XOffset
            If MapNpc(MapNpcNum).Dir = DIR_LEFT And MapNpc(MapNpcNum).Moving > 0 Then
                rec.Right = rec.Right + MapNpc(MapNpcNum).XOffset
            Else
                rec.Right = rec.Right + MapNpc(MapNpcNum).XOffset - ((SIZE_X - PIC_X) / 2)
            End If
        End If
    End If
    
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    Else
        rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64
        rec.Bottom = rec.Top + (64 - PIC_Y)
        rec.Left = (MapNpc(MapNpcNum).Dir * (3 * (64 / PIC_X)) + (Anim * (64 / PIC_X))) * PIC_X
        'rec.Left = (MapNpc(MapNpcNum).Dir * (3 + Anim)) * 64
        rec.Right = rec.Left + 64
       
        x = MapNpc(MapNpcNum).x * PIC_X - (64 - PIC_X) / 2 + sx + MapNpc(MapNpcNum).XOffset
        y = MapNpc(MapNpcNum).y * PIC_Y - (64 - PIC_Y) + sx + MapNpc(MapNpcNum).YOffset
       
        If y < 0 Then
            y = 0
            If MapNpc(MapNpcNum).Dir = DIR_DOWN And MapNpc(MapNpcNum).Moving > 0 Then
                rec.Top = rec.Top - MapNpc(MapNpcNum).YOffset
            Else
                rec.Top = rec.Top - MapNpc(MapNpcNum).YOffset + (64 - PIC_Y)
            End If
        End If
       
        If 64 > PIC_X Then
            If x < 0 Then
                x = MapNpc(MapNpcNum).XOffset + sx + ((64 - PIC_X) / 2)
                If MapNpc(MapNpcNum).Dir = DIR_RIGHT And MapNpc(MapNpcNum).Moving > 0 Then
                    rec.Left = rec.Left - MapNpc(MapNpcNum).XOffset
                Else
                    rec.Left = rec.Left - MapNpc(MapNpcNum).XOffset + ((64 - PIC_X) / 2)
                End If
            End If
       
            If x > MAX_MAPX * 42 Then
                x = MAX_MAPX * 42 + sx - ((64 - PIC_X) / 2) + MapNpc(MapNpcNum).XOffset
                If MapNpc(MapNpcNum).Dir = DIR_LEFT And MapNpc(MapNpcNum).Moving > 0 Then
                    rec.Right = rec.Right + MapNpc(MapNpcNum).XOffset
                Else
                    rec.Right = rec.Right + MapNpc(MapNpcNum).XOffset - ((64 - PIC_X) / 2)
                End If
            End If
        End If
       
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltPetName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
    
    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                Color = QBColor(Brown)
            Case 1
                Color = QBColor(DarkGrey)
            Case 2
                Color = QBColor(Cyan)
            Case 3
                Color = QBColor(Blue)
            Case 4
                Color = QBColor(Pink)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If
        
    ' Draw name
    TextX = Player(Index).Pet.x * PIC_X + sx + Player(Index).Pet.XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & "'s Pet") / 2) * 8)
    TextY = Player(Index).Pet.y * PIC_Y + sx + Player(Index).Pet.YOffset - Int(PIC_Y / 2) - (SIZE_Y - PIC_Y)
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index) & "'s Pet", Color)
End Sub

Sub BltPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
    
    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                Color = QBColor(Brown)
            Case 1
                Color = QBColor(DarkGrey)
            Case 2
                Color = QBColor(Cyan)
            Case 3
                Color = QBColor(Blue)
            Case 4
                Color = QBColor(Pink)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If
        
    ' Draw name
    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
    TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - Int(PIC_Y / 2) - (SIZE_Y - PIC_Y)
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index), Color)
End Sub

Sub BltPlayerGuildName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long

    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerGuildAccess(Index)
            Case 0
                If GetPlayerSTR(Index) > 0 Then
                    Color = QBColor(Red)
                Else
                    Color = QBColor(Red)
                End If
            Case 1
                Color = QBColor(BrightCyan)
            Case 2
                Color = QBColor(Yellow)
            Case 3
                Color = QBColor(BrightGreen)
            Case 4
                Color = QBColor(Yellow)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If

If Index = MyIndex Then
    ' Draw name
    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + Int(PIC_X * 0.5) - ((Len(GetPlayerGuild(Index)) * 0.5) * 8)
    TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - Int(PIC_Y * 0.5) - 12
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerGuild(MyIndex), Color)
Else
    ' Draw name
    TextX = GetPlayerX(MyIndex) * PIC_X + sx + Player(MyIndex).XOffset + Int(PIC_X * 0.5) - ((Len(GetPlayerGuild(MyIndex)) * 0.5) * 8)
    TextY = GetPlayerY(MyIndex) * PIC_Y + sx + Player(MyIndex).YOffset - Int(PIC_Y * 0.5) - 12
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerGuild(MyIndex), Color)
End If
End Sub

Sub ProcessMovement(ByVal Index As Long)
    ' Check if player is walking, and if so process moving them over
    If Player(Index).Moving = MOVING_WALKING Then
        If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - GM_WALK_SPEED
                    If Player(Index).YOffset <= 0 Then Player(Index).Moving = 0
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + GM_WALK_SPEED
                    If Player(Index).YOffset >= 0 Then Player(Index).Moving = 0
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - GM_WALK_SPEED
                    If Player(Index).XOffset <= 0 Then Player(Index).Moving = 0
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + GM_WALK_SPEED
                    If Player(Index).XOffset >= 0 Then Player(Index).Moving = 0
            End Select
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - WALK_SPEED
                    If Player(Index).YOffset <= 0 Then Player(Index).Moving = 0
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + WALK_SPEED
                    If Player(Index).YOffset >= 0 Then Player(Index).Moving = 0
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - WALK_SPEED
                    If Player(Index).XOffset <= 0 Then Player(Index).Moving = 0
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + WALK_SPEED
                    If Player(Index).XOffset >= 0 Then Player(Index).Moving = 0
            End Select
        End If
    End If

    ' Check if player is running, and if so process moving them over
    If Player(Index).Moving = MOVING_RUNNING Then
        If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - GM_RUN_SPEED
                    If Player(Index).YOffset <= 0 Then Player(Index).Moving = 0
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + GM_RUN_SPEED
                    If Player(Index).YOffset >= 0 Then Player(Index).Moving = 0
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - GM_RUN_SPEED
                    If Player(Index).XOffset <= 0 Then Player(Index).Moving = 0
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + GM_RUN_SPEED
                    If Player(Index).XOffset >= 0 Then Player(Index).Moving = 0
            End Select
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - RUN_SPEED
                    If Player(Index).YOffset <= 0 Then Player(Index).Moving = 0
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + RUN_SPEED
                    If Player(Index).YOffset >= 0 Then Player(Index).Moving = 0
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - RUN_SPEED
                    If Player(Index).XOffset <= 0 Then Player(Index).Moving = 0
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + RUN_SPEED
                    If Player(Index).XOffset >= 0 Then Player(Index).Moving = 0
            End Select
        End If
    End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if npc is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - WALK_SPEED
            Case DIR_DOWN
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + WALK_SPEED
            Case DIR_LEFT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - WALK_SPEED
            Case DIR_RIGHT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (MapNpc(MapNpcNum).XOffset = 0) And (MapNpc(MapNpcNum).YOffset = 0) Then
            MapNpc(MapNpcNum).Moving = 0
        End If
    End If
End Sub

Sub ProcessPetMovement(ByVal PetNum As Long)
    ' Check if pet is walking, and if so process moving them over
    If Player(PetNum).Pet.Moving = MOVING_WALKING Then
        Select Case Player(PetNum).Pet.Dir
            Case DIR_UP
                Player(PetNum).Pet.YOffset = Player(PetNum).Pet.YOffset - WALK_SPEED
            Case DIR_DOWN
                Player(PetNum).Pet.YOffset = Player(PetNum).Pet.YOffset + WALK_SPEED
            Case DIR_LEFT
                Player(PetNum).Pet.XOffset = Player(PetNum).Pet.XOffset - WALK_SPEED
            Case DIR_RIGHT
                Player(PetNum).Pet.XOffset = Player(PetNum).Pet.XOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (Player(PetNum).Pet.XOffset = 0) And (Player(PetNum).Pet.YOffset = 0) Then
            Player(PetNum).Pet.Moving = 0
        End If
    End If
End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim Name As String
Dim I As Long
Dim n As Long

MyText = frmMirage.txtMyTextBox.Text

    ' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
    
        ' Tell the server they hit enter, and execute the sadscript
        If frmMirage.txtMyTextBox.Text = vbNullString Then
            Call SendData(RETURNSCRIPT_CHAR & END_CHAR)
            Exit Sub
        End If
    
        frmMirage.txtMyTextBox.Text = vbNullString
        
        If Player(MyIndex).y - 1 > -1 Then
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN And Player(MyIndex).Dir = DIR_UP Then
                Call AddText("The Sign Reads:", Black)
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1) <> vbNullString Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1), Grey)
                End If
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2) <> vbNullString Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2), Grey)
                End If
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3) <> vbNullString Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3), Grey)
                End If
            Exit Sub
            End If
        End If
        ' Broadcast message
        If Mid(MyText, 1, 1) = "," Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call BroadcastMsg(ChatText)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Emote message
        If Mid(MyText, 1, 1) = "-" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Party message
        If Mid(MyText, 1, 1) = "+" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call PartyMsg(ChatText)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Guild message
        If Mid(MyText, 1, 1) = "=" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call GuildMsg(ChatText)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Commands
        If Mid(MyText, 1, 9) = "/commands" Then
            Call AddText(":::::::: Commands ::::::::", White)
            Call AddText(", (broadcast), - (emote), + (party), = (guild), ! (private)", White)
            Call AddText("/info, /who, /fps, /inv, /stats, /chat, /chatdecline, /trade, /accept, /decline, /party, /join, /leave, /killpet, /refresh", White)
            Exit Sub
        End If
        
        ' Admin Commands
        If Mid(MyText, 1, 14) = "/admincommands" Then
            Call AddText(":::::::: Admin Commands ::::::::", White)
            Call AddText("; (global), @ (admin)", White)
            Call AddText("/daynight, /weather, /kick", White)
            Call AddText("/loc, /warp, /warptome, /mapeditor, /mapreport, /setsprite, /setplayersprite, /respawn, /motd, /banlist, /ban", White)
            Call AddText("/edititem, /editarrow, /editemoticon, /editnpc, /editshop, /editspell, /editspeech", White)
            Call AddText("/setaccess, /nullbanlist, /debug, /editmain, /mainbackup", White)
            Exit Sub
        End If
        
        ' Player message
        If Mid(MyText, 1, 1) = "!" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            Name = vbNullString
                    
            ' Get the desired player from the user text
            For I = 1 To Len(ChatText)
                If Mid(ChatText, I, 1) <> " " Then
                    Name = Name & Mid(ChatText, I, 1)
                Else
                    Exit For
                End If
            Next I
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - I > 0 Then
                ChatText = Mid(ChatText, I + 1, Len(ChatText) - I)
                    
                ' Send the message to the player
                Call PlayerMsg(ChatText, Name)
            Else
                Call AddText("Usage: !playername msghere", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
            
        ' // Commands //
        ' Verification User
        If LCase$(Mid(MyText, 1, 5)) = "/info" Then
            ChatText = Mid(MyText, 6, Len(MyText) - 5)
            Call SendData(PLAYERINFOREQUEST_CHAR & SEP_CHAR & ChatText & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Whos Online
        If LCase$(Mid(MyText, 1, 4)) = "/who" Then
            Call SendWhosOnline
            MyText = vbNullString
            Exit Sub
        End If
                        
        ' Checking fps
        If LCase$(Mid(MyText, 1, 4)) = "/fps" Then
            Call AddText("FPS: " & GameFPS, Pink)
            MyText = vbNullString
            Exit Sub
        End If
                
        ' Show inventory
        If LCase$(Mid(MyText, 1, 4)) = "/inv" Then
            frmMirage.picInv3.Visible = True
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Request stats
        If LCase$(Mid(MyText, 1, 6)) = "/stats" Then
            Call SendData(GETSTATS_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
         
        ' Refresh Player
        If LCase$(Mid(MyText, 1, 8)) = "/refresh" Then
            Call SendData(REFRESH_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Decline Chat
        If LCase$(Mid(MyText, 1, 12)) = "/chatdecline" Then
            Call SendData(DCHAT_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Accept Chat
        If LCase$(Mid(MyText, 1, 5)) = "/chat" Then
            Call SendData(ACHAT_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        If LCase$(Mid(MyText, 1, 6)) = "/trade" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid(MyText, 8, Len(MyText) - 7)
                Call SendTradeRequest(ChatText)
            Else
                Call AddText("Usage: /trade playernamehere", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Accept Trade
        If LCase$(Mid(MyText, 1, 7)) = "/accept" Then
            Call SendAcceptTrade
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Decline Trade
        If LCase$(Mid(MyText, 1, 8)) = "/decline" Then
            Call SendDeclineTrade
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Party request
        If LCase$(Mid(MyText, 1, 6)) = "/party" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid(MyText, 8, Len(MyText) - 7)
                Call SendPartyRequest(ChatText)
            Else
                Call AddText("Usage: /party playernamehere", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Kill pet
        If LCase$(Mid(MyText, 1, 8)) = "/killpet" Then
            Call SendData(KILLPET_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Join party
        If LCase$(Mid(MyText, 1, 5)) = "/join" Then
            Call SendJoinParty
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Leave party
        If LCase$(Mid(MyText, 1, 6)) = "/leave" Then
            Call SendLeaveParty
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Change music volume
        If LCase$(Mid(MyText, 1, 6)) = "/music" Then
            If Len(MyText) > 6 Then
                MyText = Mid(MyText, 7, Len(MyText) - 6)
                If IsNumeric(MyText) = True Then
                    If MyText > -1 And MyText < 101 Then
                        perf.SetMasterVolume (MyText * 42 - 3000)
                        frmMirage.scrlMusicVol.Value = MyText
                        frmMirage.lblMusicVol.Caption = MyText
                        Call PutVar(App.Path & "\config.ini", "CONFIG", "MusicVol", MyText)
                        MusicVolume = MyText
                        Call AddText("Music volume: " & MyText, 4)
                    Else
                        Call AddText("Number must be between 0 and 100!", BrightRed)
                    End If
                Else
                    Call AddText("You must enter a number! Usage: /music 100", BrightRed)
                End If
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' day night command
            If LCase$(Mid(MyText, 1, 9)) = "/daynight" Then
                If GameTime = TIME_DAY Then
                    GameTime = TIME_NIGHT
                Else
                    GameTime = TIME_DAY
                End If
                Call SendGameTime
                MyText = vbNullString
                Exit Sub
            End If
            
            ' weather command
            If LCase$(Mid(MyText, 1, 8)) = "/weather" Then
                If Len(MyText) > 8 Then
                    MyText = Mid(MyText, 9, Len(MyText) - 8)
                    If IsNumeric(MyText) = True Then
                        Call SendData(WEATHER_CHAR & SEP_CHAR & Val(MyText) & END_CHAR)
                    Else
                        If Trim$(LCase$(MyText)) = "none" Then I = 0
                        If Trim$(LCase$(MyText)) = "rain" Then I = 1
                        If Trim$(LCase$(MyText)) = "snow" Then I = 2
                        If Trim$(LCase$(MyText)) = "thunder" Then I = 3
                        Call SendData(WEATHER_CHAR & SEP_CHAR & I & END_CHAR)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Kicking a player
            If LCase$(Mid(MyText, 1, 5)) = "/kick" Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    Call SendKick(MyText)
                End If
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Global Message
            If Mid(MyText, 1, 1) = ";" Then
                ChatText = Mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Admin Message
            If Mid(MyText, 1, 1) = "@" Then
                ChatText = Mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call AdminMsg(ChatText)
                End If
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' // Mapper Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
            ' Location
            If LCase$(Mid(MyText, 1, 4)) = "/loc" Then
                Call SendRequestLocation
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Warpe
            If LCase$(Mid(MyText, 1, 6)) = "/warp " Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    Call SendWarp(MyText)
                End If
                Exit Sub
            End If
            
            ' Map Editor
            If LCase$(Mid(MyText, 1, 8)) = "/editmap" Or LCase$(Mid(MyText, 1, 10)) = "/mapeditor" Then
                Call SendRequestEditMap
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Map report
            If LCase$(Mid(MyText, 1, 10)) = "/mapreport" Then
                Call SendData(MAPREPORT_CHAR & END_CHAR)
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Setting sprite
            If LCase$(Mid(MyText, 1, 10)) = "/setsprite" Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = Mid(MyText, 12, Len(MyText) - 11)
                
                    Call SendSetSprite(Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Setting player sprite
            If LCase$(Mid(MyText, 1, 16)) = "/setplayersprite" Then
                If Len(MyText) > 19 Then
                    I = Val(Mid(MyText, 17, 1))
                
                    MyText = Mid(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerSprite(I, Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Respawn request
            If Mid(MyText, 1, 8) = "/respawn" Then
                Call SendMapRespawn
                MyText = vbNullString
                Exit Sub
            End If
        
            ' MOTD change
            If Mid(MyText, 1, 5) = "/motd" Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    If Trim$(MyText) <> vbNullString Then
                        Call SendMOTDChange(MyText)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Check the ban list
            If Mid(MyText, 1, 3) = "/banlist" Then
                Call SendBanList
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Banning a player
            If LCase$(Mid(MyText, 1, 4)) = "/ban" Then
                If Len(MyText) > 5 Then
                    MyText = Mid(MyText, 6, Len(MyText) - 5)
                    Call SendBan(MyText)
                    MyText = vbNullString
                End If
                Exit Sub
            End If
        End If
            
        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Editing item request
            If Mid(MyText, 1, 9) = "/edititem" Or Mid(MyText, 1, 11) = "/itemeditor" Then
                Call SendRequestEditItem
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing emoticon request
            If Mid(MyText, 1, 13) = "/editemoticon" Or Mid(MyText, 1, 15) = "/emoticoneditor" Then
                Call SendRequestEditEmoticon
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing arrow request
            If Mid(MyText, 1, 10) = "/editarrow" Or Mid(MyText, 1, 12) = "/arroweditor" Then
                Call SendRequestEditArrow
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing speech request
            If Mid(MyText, 1, 11) = "/editspeech" Or Mid(MyText, 1, 13) = "/speecheditor" Then
                Call SendRequestEditSpeech
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing npc request
            If Mid(MyText, 1, 8) = "/editnpc" Or Mid(MyText, 1, 10) = "/npceditor" Then
                Call SendRequestEditNpc
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing shop request
            If Mid(MyText, 1, 9) = "/editshop" Or Mid(MyText, 1, 11) = "/shopeditor" Then
                Call SendRequestEditShop
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Editing spell request
            If Mid(MyText, 1, 10) = "/editspell" Or Mid(MyText, 1, 12) = "/spelleditor" Then
                Call SendRequestEditSpell
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            ' Giving another player access
            If LCase$(Mid(MyText, 1, 10)) = "/setaccess" Then
                ' Get access #
                I = Val(Mid(MyText, 12, 1))
                
                MyText = Mid(MyText, 14, Len(MyText) - 13)
                
                Call SendSetAccess(MyText, I)
                MyText = vbNullString
                Exit Sub
            End If
                    
            ' Edit main.txt
            If Mid(MyText, 1, 9) = "/editmain" Or Mid(MyText, 1, 11) = "/maineditor" Then
                Call SendRequestEditMain
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Reload the backup
            If Mid(MyText, 1, 12) = "/mainbackup" Then
                Call SendRequestBackupMain
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Debugging
            If LCase$(Mid(MyText, 1, 6)) = "/debug" Then
                If GoDebug = YES Then
                    GoDebug = NO
                Else
                    GoDebug = YES
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Ban destroy
            If LCase$(Mid(MyText, 1, 15)) = "/destroybanlist" Or LCase$(Mid(MyText, 1, 12)) = "/nullbanlist" Then
                Call SendBanDestroy
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' Tell them its not a valid command
        If Left$(Trim$(MyText), 1) = "/" Then
            For I = 0 To MAX_EMOTICONS
                If Trim$(Emoticons(I).Command) = Trim$(MyText) And Trim$(Emoticons(I).Command) <> "/" Then
                    Call SendData(CHECKEMOTICONS_CHAR & SEP_CHAR & I & END_CHAR)
                    MyText = vbNullString
                Exit Sub
                End If
            Next I
            Call SendData(CHECKCOMMANDS_CHAR & SEP_CHAR & MyText & END_CHAR)
            MyText = vbNullString
        Exit Sub
        End If
            
        ' Say message
        If Len(Trim$(MyText)) > 0 Then
            Call SayMsg(MyText)
        End If
        MyText = vbNullString
        Exit Sub
    End If
    
    'frmMirage.txtMyTextBox.SetFocus
    
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If Len(MyText) > 0 Then
            'MyText = Mid(MyText, 1, Len(MyText) - 1)
        End If
    End If
    
    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) Then
        ' Make sure they just use standard keys, no gay shitty macro keys
        If KeyAscii >= 32 And KeyAscii <= 126 Then
            'frmMirage.txtMyTextBox.Text = frmMirage.txtMyTextBox.Text & Chr(KeyAscii)
            'MyText = MyText & Chr(KeyAscii)
        End If
    End If
End Sub

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim$(MyText) = vbNullString Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData(MAPGETITEM_CHAR & END_CHAR)
    End If
End Sub

Sub CheckAttack()
Dim AttackSpeed As Long
Dim I As Long
    If GetPlayerWeaponSlot(MyIndex) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If
    
    If ControlDown = True And Player(MyIndex).AttackTimer + AttackSpeed < GetTickCount And Player(MyIndex).Attacking = 0 Then
        Player(MyIndex).Attacking = 1
        Player(MyIndex).AttackTimer = GetTickCount
        Call SendData(ATTACK_CHAR & END_CHAR)
    End If
End Sub

Sub CheckInput2()
    If GettingMap = False Then
        If GetKeyState(VK_RETURN) < 0 Then
            Call CheckMapGetItem
        End If
        If GetKeyState(VK_CONTROL) < 0 Then
            ControlDown = True
        Else
            ControlDown = False
        End If
        If GetKeyState(VK_UP) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
        Else
            DirUp = False
        End If
        If GetKeyState(VK_DOWN) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
        Else
            DirDown = False
        End If
        If GetKeyState(VK_LEFT) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
        Else
            DirLeft = False
        End If
        If GetKeyState(VK_RIGHT) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
        Else
            DirRight = False
        End If
        If GetKeyState(VK_SHIFT) < 0 Then
            ShiftDown = True
        Else
            ShiftDown = False
        End If
    End If
End Sub

Sub CheckInput(ByVal KeyState As Byte, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If GettingMap = False Then
        If KeyState = 1 Then
            If KeyCode = vbKeyReturn Then
                Call CheckMapGetItem
            End If
            If KeyCode = vbKeyControl Then
                ControlDown = True
            End If
            If KeyCode = vbKeyUp Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = vbKeyDown Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = vbKeyLeft Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
            End If
            If KeyCode = vbKeyRight Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
            End If
            If KeyCode = vbKeyShift Then
                ShiftDown = True
            End If
        Else
            If KeyCode = vbKeyUp Then DirUp = False
            If KeyCode = vbKeyDown Then DirDown = False
            If KeyCode = vbKeyLeft Then DirLeft = False
            If KeyCode = vbKeyRight Then DirRight = False
            If KeyCode = vbKeyShift Then ShiftDown = False
            If KeyCode = vbKeyControl Then ControlDown = False
        End If
    End If
End Sub

Function IsTryingToMove() As Boolean
    If (DirUp = True) Or (DirDown = True) Or (DirLeft = True) Or (DirRight = True) Then
        IsTryingToMove = True
    Else
        IsTryingToMove = False
    End If
End Function

Function CanMove() As Boolean
Dim I As Long, d As Long
Dim x As Long, y As Long

    CanMove = True
    
    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' Make sure they haven't just casted a spell
    If Player(MyIndex).CastedSpell = YES Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Player(MyIndex).CastedSpell = NO
        Else
            CanMove = False
            Exit Function
        End If
    End If
    
    d = GetPlayerDir(MyIndex)
    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            ' Check to see if the map tile is blocked or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN Then
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_BLOCKED Or TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_NONE Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_UP Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            Else
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_BLOCKED Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_UP Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_CBLOCK Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data1 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data2 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
            End If
                                                    
            ' Check to see if the key door is open or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_KEY Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).DoorOpen = NO Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_UP Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For I = 1 To MAX_PLAYERS
                If I <> MyIndex Then
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                            If (GetPlayerX(I) = GetPlayerX(MyIndex)) And (GetPlayerY(I) = GetPlayerY(MyIndex) - 1) Then
                                CanMove = False
                            
                                ' Set the new direction if they weren't facing that direction
                                If d <> DIR_UP Then
                                    Call SendPlayerDir
                                End If
                                Exit Function
                            End If
                        End If
                        
                        ' Might as well check for pets too
                        If Player(I).Pet.Alive = YES Then
                            If Player(I).Pet.Map = GetPlayerMap(MyIndex) Then
                                If Player(I).Pet.x = GetPlayerX(MyIndex) And Player(I).Pet.y = GetPlayerY(MyIndex) - 1 Then
                                    CanMove = False
                            
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> DIR_UP Then
                                        Call SendPlayerDir
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Else
                    If Player(I).Pet.Alive = YES Then
                        If Player(I).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(I).Pet.x = GetPlayerX(MyIndex) And Player(I).Pet.y = GetPlayerY(MyIndex) - 1 Then
                                If IsValid(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 2) Then
                                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 2).Type = TILE_TYPE_BLOCKED Then
                                        CanMove = False
                                
                                        ' Set the new direction if they weren't facing that direction
                                        If d <> DIR_UP Then
                                            Call SendPlayerDir
                                        End If
                                        Exit Function
                                    End If
                                Else
                                    CanMove = False
                                
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> DIR_UP Then
                                        Call SendPlayerDir
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next I
        
            ' Check to see if a npc is already on that tile
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(I).Num > 0 Then
                    If (MapNpc(I).x = GetPlayerX(MyIndex)) And (MapNpc(I).y = GetPlayerY(MyIndex) - 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_UP Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next I
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Up > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
            
    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < MAX_MAPY Then
            ' Check to see if the map tile is blocked or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_SIGN Then
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_BLOCKED Or TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_NONE Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_DOWN Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            Else
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_BLOCKED Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_DOWN Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_CBLOCK Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data1 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data2 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_KEY Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).DoorOpen = NO Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_DOWN Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For I = 1 To MAX_PLAYERS
                If I <> MyIndex Then
                    If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                        If (GetPlayerX(I) = GetPlayerX(MyIndex)) And (GetPlayerY(I) = GetPlayerY(MyIndex) + 1) Then
                            CanMove = False
                            
                            ' Set the new direction if they weren't facing that direction
                            If d <> DIR_DOWN Then
                                Call SendPlayerDir
                            End If
                            Exit Function
                        End If
                    End If
                    
                    ' Might as well check for pets too
                    If Player(I).Pet.Alive = YES Then
                        If Player(I).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(I).Pet.x = GetPlayerX(MyIndex) And Player(I).Pet.y = GetPlayerY(MyIndex) + 1 Then
                                CanMove = False
                        
                                ' Set the new direction if they weren't facing that direction
                                If d <> DIR_DOWN Then
                                    Call SendPlayerDir
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    If Player(I).Pet.Alive = YES Then
                        If Player(I).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(I).Pet.x = GetPlayerX(MyIndex) And Player(I).Pet.y = GetPlayerY(MyIndex) + 1 Then
                                If IsValid(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 2) Then
                                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 2).Type = TILE_TYPE_BLOCKED Then
                                        CanMove = False
                                
                                        ' Set the new direction if they weren't facing that direction
                                        If d <> DIR_DOWN Then
                                            Call SendPlayerDir
                                        End If
                                        Exit Function
                                    End If
                                Else
                                    CanMove = False
                                
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> DIR_DOWN Then
                                        Call SendPlayerDir
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next I
            
            ' Check to see if a npc is already on that tile
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(I).Num > 0 Then
                    If (MapNpc(I).x = GetPlayerX(MyIndex)) And (MapNpc(I).y = GetPlayerY(MyIndex) + 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_DOWN Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next I
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Down > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
                
    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            ' Check to see if the map tile is blocked or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SIGN Then
                If TempTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or TempTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_NONE Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_LEFT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            Else
                If TempTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_LEFT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_CBLOCK Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data1 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data2 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).DoorOpen = NO Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_LEFT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For I = 1 To MAX_PLAYERS
                If I <> MyIndex Then
                    If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                        If (GetPlayerX(I) = GetPlayerX(MyIndex) - 1) And (GetPlayerY(I) = GetPlayerY(MyIndex)) Then
                            CanMove = False
                            
                            ' Set the new direction if they weren't facing that direction
                            If d <> DIR_LEFT Then
                                Call SendPlayerDir
                            End If
                            Exit Function
                        End If
                    End If
                    
                    ' Might as well check for pets too
                    If Player(I).Pet.Alive = YES Then
                        If Player(I).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(I).Pet.x = GetPlayerX(MyIndex) - 1 And Player(I).Pet.y = GetPlayerY(MyIndex) Then
                                CanMove = False
                        
                                ' Set the new direction if they weren't facing that direction
                                If d <> DIR_LEFT Then
                                    Call SendPlayerDir
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    If Player(I).Pet.Alive = YES Then
                        If Player(I).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(I).Pet.x = GetPlayerX(MyIndex) - 1 And Player(I).Pet.y = GetPlayerY(MyIndex) Then
                                If IsValid(GetPlayerX(MyIndex) - 2, GetPlayerY(MyIndex)) Then
                                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 2, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                                        CanMove = False
                                
                                        ' Set the new direction if they weren't facing that direction
                                        If d <> DIR_LEFT Then
                                            Call SendPlayerDir
                                        End If
                                        Exit Function
                                    End If
                                Else
                                    CanMove = False
                                
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> DIR_LEFT Then
                                        Call SendPlayerDir
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next I
        
            ' Check to see if a npc is already on that tile
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(I).Num > 0 Then
                    If (MapNpc(I).x = GetPlayerX(MyIndex) - 1) And (MapNpc(I).y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_LEFT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next I
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Left > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
        
    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < MAX_MAPX Then
            ' Check to see if the map tile is blocked or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SIGN Then
                If TempTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or TempTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_NONE Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_RIGHT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            Else
                If TempTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_RIGHT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_CBLOCK Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data1 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data2 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).DoorOpen = NO Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_RIGHT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For I = 1 To MAX_PLAYERS
                If I <> MyIndex Then
                    If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                        If (GetPlayerX(I) = GetPlayerX(MyIndex) + 1) And (GetPlayerY(I) = GetPlayerY(MyIndex)) Then
                            CanMove = False
                            
                            ' Set the new direction if they weren't facing that direction
                            If d <> DIR_RIGHT Then
                                Call SendPlayerDir
                            End If
                            Exit Function
                        End If
                    End If
                    
                    ' Might as well check for pets too
                    If Player(I).Pet.Alive = YES Then
                        If Player(I).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(I).Pet.x = GetPlayerX(MyIndex) + 1 And Player(I).Pet.y = GetPlayerY(MyIndex) Then
                                CanMove = False
                        
                                ' Set the new direction if they weren't facing that direction
                                If d <> DIR_RIGHT Then
                                    Call SendPlayerDir
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    If Player(I).Pet.Alive = YES Then
                        If Player(I).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(I).Pet.x = GetPlayerX(MyIndex) + 1 And Player(I).Pet.y = GetPlayerY(MyIndex) Then
                                If IsValid(GetPlayerX(MyIndex) + 2, GetPlayerY(MyIndex)) Then
                                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 2, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                                        CanMove = False
                                
                                        ' Set the new direction if they weren't facing that direction
                                        If d <> DIR_RIGHT Then
                                            Call SendPlayerDir
                                        End If
                                        Exit Function
                                    End If
                                Else
                                    CanMove = False
                                
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> DIR_RIGHT Then
                                        Call SendPlayerDir
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next I
        
            ' Check to see if a npc is already on that tile
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(I).Num > 0 Then
                    If (MapNpc(I).x = GetPlayerX(MyIndex) + 1) And (MapNpc(I).y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_RIGHT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next I
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Right > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
End Function

Sub CheckMovement()
    If GettingMap = False Then
        If IsTryingToMove Then
            If CanMove Then
                ' Check if player has the shift key down for running
                If ShiftDown Then
                    Player(MyIndex).Moving = MOVING_RUNNING
                Else
                    Player(MyIndex).Moving = MOVING_WALKING
                End If
                
                Select Case GetPlayerDir(MyIndex)
                    Case DIR_UP
                        Call SendPlayerMove
                        Player(MyIndex).YOffset = PIC_Y
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                
                    Case DIR_DOWN
                        Call SendPlayerMove
                        Player(MyIndex).YOffset = PIC_Y * -1
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                
                    Case DIR_LEFT
                        Call SendPlayerMove
                        Player(MyIndex).XOffset = PIC_X
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                
                    Case DIR_RIGHT
                        Call SendPlayerMove
                        Player(MyIndex).XOffset = PIC_X * -1
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                End Select
            
                ' Gotta check :)
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                    GettingMap = True
                End If
            End If
        End If
    End If
End Sub

Function FindPlayer(ByVal Name As String) As Long
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            ' Make sure we don't try to check a name thats to small
            If Len(GetPlayerName(I)) >= Len(Trim$(Name)) Then
                If UCase(Mid(GetPlayerName(I), 1, Len(Trim$(Name)))) = UCase(Trim$(Name)) Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If
    Next I
    
    FindPlayer = 0
End Function

Public Sub EditorInit()
    Dim I As Long
    
    Call StopMidi

    InEditor = True
    InSpawnEditor = False
    frmMapEditor.Show vbModeless, frmMirage

    frmMapEditor.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles" & EditorSet & ".bmp")

    frmMapEditor.scrlPicture.Max = Int((frmMapEditor.picBackSelect.Height - frmMapEditor.picBack.Height) / PIC_Y)
    frmMapEditor.picBack.Width = 448
   
    If GameTime = TIME_NIGHT Then frmMapEditor.chkDayNight.Value = 1
    If GameTime = TIME_DAY Then frmMapEditor.chkDayNight.Value = 0
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, y1 As Long
Dim x2 As Long, y2 As Long, PicX As Long

    If InEditor Then
        x1 = Int(x / PIC_X)
        y1 = Int(y / PIC_Y)
       
        If frmMapEditor.MousePointer = 2 Then
            If frmMapEditor.optTiles.Value = 1 Then
                With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    If frmMapEditor.optGround.Value = True Then
                        PicX = .Ground
                        EditorSet = .GroundSet
                    End If
                    If frmMapEditor.optMask.Value = True Then
                        PicX = .Mask
                        EditorSet = .MaskSet
                    End If
                    If frmMapEditor.optAnim.Value = True Then
                        PicX = .Anim
                        EditorSet = .AnimSet
                    End If
                    If frmMapEditor.optMask2.Value = True Then
                        PicX = .Mask2
                        EditorSet = .Mask2Set
                    End If
                    If frmMapEditor.optM2Anim.Value = True Then
                        PicX = .M2Anim
                        EditorSet = .M2AnimSet
                    End If
                    If frmMapEditor.optFringe.Value = True Then
                        PicX = .Fringe
                        EditorSet = .FringeSet
                    End If
                    If frmMapEditor.optFAnim.Value = True Then
                        PicX = .FAnim
                        EditorSet = .FAnimSet
                    End If
                    If frmMapEditor.optFringe2.Value = True Then
                        PicX = .Fringe2
                        EditorSet = .Fringe2Set
                    End If
                    If frmMapEditor.optF2Anim.Value = True Then
                        PicX = .F2Anim
                        EditorSet = .F2AnimSet
                    End If
                   
                    EditorTileY = Int(PicX / TilesInSheets)
                    EditorTileX = (PicX - Int(PicX / TilesInSheets) * TilesInSheets)
                    frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                    frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                    frmMapEditor.shpSelected.Height = PIC_Y
                    frmMapEditor.shpSelected.Width = PIC_X
                End With
            ElseIf frmMapEditor.optlight.Value = True Then
                EditorTileY = Int(Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light / TilesInSheets)
                EditorTileX = (Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light - Int(Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light / TilesInSheets) * TilesInSheets)
                frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                frmMapEditor.shpSelected.Height = PIC_Y
                frmMapEditor.shpSelected.Width = PIC_X
            ElseIf frmMapEditor.optAttributes.Value = True Then
                With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    If .Type = TILE_TYPE_BLOCKED Then frmMapEditor.optBlocked.Value = True
                    If .Type = TILE_TYPE_WARP Then
                        EditorWarpMap = .Data1
                        EditorWarpX = .Data2
                        EditorWarpY = .Data3
                        frmMapEditor.optWarp.Value = True
                    End If
                    If .Type = TILE_TYPE_HEAL Then frmMapEditor.optHeal.Value = True
                    If .Type = TILE_TYPE_KILL Then frmMapEditor.optKill.Value = True
                    If .Type = TILE_TYPE_ITEM Then
                        ItemEditorNum = .Data1
                        ItemEditorValue = .Data2
                        frmMapEditor.optItem.Value = True
                    End If
                    If .Type = TILE_TYPE_NPCAVOID Then frmMapEditor.optNpcAvoid.Value = True
                    If .Type = TILE_TYPE_KEY Then
                        KeyEditorNum = .Data1
                        KeyEditorTake = .Data2
                        frmMapEditor.optKey.Value = True
                    End If
                    If .Type = TILE_TYPE_KEYOPEN Then
                        KeyOpenEditorX = .Data1
                        KeyOpenEditorY = .Data2
                        KeyOpenEditorMsg = .String1
                        frmMapEditor.optKeyOpen.Value = True
                    End If
                    If .Type = TILE_TYPE_SHOP Then
                        EditorShopNum = .Data1
                        frmMapEditor.optShop.Value = True
                    End If
                    If .Type = TILE_TYPE_CBLOCK Then
                        EditorItemNum1 = .Data1
                        EditorItemNum2 = .Data2
                        EditorItemNum3 = .Data3
                        frmMapEditor.optCBlock.Value = True
                    End If
                    If .Type = TILE_TYPE_ARENA Then
                        Arena1 = .Data1
                        Arena2 = .Data2
                        Arena3 = .Data3
                        frmMapEditor.optArena.Value = True
                    End If
                    If .Type = TILE_TYPE_SOUND Then
                        SoundFileName = .String1
                        frmMapEditor.optSound.Value = True
                    End If
                    If .Type = TILE_TYPE_SPRITE_CHANGE Then
                        SpritePic = .Data1
                        SpriteItem = .Data2
                        SpritePrice = .Data3
                        frmMapEditor.optSprite.Value = True
                    End If
                    If .Type = TILE_TYPE_SIGN Then
                        SignLine1 = .String1
                        SignLine2 = .String2
                        SignLine3 = .String3
                        frmMapEditor.optSign.Value = True
                    End If
                    If .Type = TILE_TYPE_DOOR Then frmMapEditor.optDoor.Value = True
                    If .Type = TILE_TYPE_NOTICE Then
                        NoticeTitle = .String1
                        NoticeText = .String2
                        NoticeSound = .String3
                        frmMapEditor.optNotice.Value = True
                    End If
                    If .Type = TILE_TYPE_CHEST Then frmMapEditor.optChest.Value = True
                    If .Type = TILE_TYPE_CLASS_CHANGE Then
                        ClassChange = .Data1
                        ClassChangeReq = .Data2
                        frmMapEditor.optClassChange.Value = True
                    End If
                    If .Type = TILE_TYPE_SCRIPTED Then
                        ScriptNum = .Data1
                        frmMapEditor.optScripted.Value = True
                    End If
                End With
            End If
            frmMapEditor.MousePointer = 1
            frmMirage.MousePointer = 1
        Else
            If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmMapEditor.shpSelected.Height <= PIC_Y And frmMapEditor.shpSelected.Width <= PIC_X Then
                    If frmMapEditor.optTiles = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                            If frmMapEditor.optGround.Value = True Then
                                .Ground = EditorTileY * TilesInSheets + EditorTileX
                                .GroundSet = EditorSet
                            End If
                            If frmMapEditor.optMask.Value = True Then
                                .Mask = EditorTileY * TilesInSheets + EditorTileX
                                .MaskSet = EditorSet
                            End If
                            If frmMapEditor.optAnim.Value = True Then
                                .Anim = EditorTileY * TilesInSheets + EditorTileX
                                .AnimSet = EditorSet
                            End If
                            If frmMapEditor.optMask2.Value = True Then
                                .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                                .Mask2Set = EditorSet
                            End If
                            If frmMapEditor.optM2Anim.Value = True Then
                                .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .M2AnimSet = EditorSet
                            End If
                            If frmMapEditor.optFringe.Value = True Then
                                .Fringe = EditorTileY * TilesInSheets + EditorTileX
                                .FringeSet = EditorSet
                            End If
                            If frmMapEditor.optFAnim.Value = True Then
                                .FAnim = EditorTileY * TilesInSheets + EditorTileX
                                .FAnimSet = EditorSet
                            End If
                            If frmMapEditor.optFringe2.Value = True Then
                                .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                                .Fringe2Set = EditorSet
                            End If
                            If frmMapEditor.optF2Anim.Value = True Then
                                .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .F2AnimSet = EditorSet
                            End If
                        End With
                    ElseIf frmMapEditor.optlight.Value = True Then
                        Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light = EditorTileY * TilesInSheets + EditorTileX
                    ElseIf frmMapEditor.optAttributes = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                            If frmMapEditor.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                            If frmMapEditor.optWarp.Value = True Then
                                .Type = TILE_TYPE_WARP
                                .Data1 = EditorWarpMap
                                .Data2 = EditorWarpX
                                .Data3 = EditorWarpY
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
       
                            If frmMapEditor.optHeal.Value = True Then
                                .Type = TILE_TYPE_HEAL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
       
                            If frmMapEditor.optKill.Value = True Then
                                .Type = TILE_TYPE_KILL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
       
                            If frmMapEditor.optItem.Value = True Then
                                .Type = TILE_TYPE_ITEM
                                .Data1 = ItemEditorNum
                                .Data2 = ItemEditorValue
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optNpcAvoid.Value = True Then
                                .Type = TILE_TYPE_NPCAVOID
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optKey.Value = True Then
                                .Type = TILE_TYPE_KEY
                                .Data1 = KeyEditorNum
                                .Data2 = KeyEditorTake
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optKeyOpen.Value = True Then
                                .Type = TILE_TYPE_KEYOPEN
                                .Data1 = KeyOpenEditorX
                                .Data2 = KeyOpenEditorY
                                .Data3 = 0
                                .String1 = KeyOpenEditorMsg
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optShop.Value = True Then
                                .Type = TILE_TYPE_SHOP
                                .Data1 = EditorShopNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optCBlock.Value = True Then
                                .Type = TILE_TYPE_CBLOCK
                                .Data1 = EditorItemNum1
                                .Data2 = EditorItemNum2
                                .Data3 = EditorItemNum3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optArena.Value = True Then
                                .Type = TILE_TYPE_ARENA
                                .Data1 = Arena1
                                .Data2 = Arena2
                                .Data3 = Arena3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSound.Value = True Then
                                .Type = TILE_TYPE_SOUND
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SoundFileName
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSprite.Value = True Then
                                .Type = TILE_TYPE_SPRITE_CHANGE
                                .Data1 = SpritePic
                                .Data2 = SpriteItem
                                .Data3 = SpritePrice
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSign.Value = True Then
                                .Type = TILE_TYPE_SIGN
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SignLine1
                                .String2 = SignLine2
                                .String3 = SignLine3
                            End If
                            If frmMapEditor.optDoor.Value = True Then
                                .Type = TILE_TYPE_DOOR
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optNotice.Value = True Then
                                .Type = TILE_TYPE_NOTICE
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = NoticeTitle
                                .String2 = NoticeText
                                .String3 = NoticeSound
                            End If
                            If frmMapEditor.optChest.Value = True Then
                                .Type = TILE_TYPE_CHEST
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optClassChange.Value = True Then
                                .Type = TILE_TYPE_CLASS_CHANGE
                                .Data1 = ClassChange
                                .Data2 = ClassChangeReq
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optScripted.Value = True Then
                                .Type = TILE_TYPE_SCRIPTED
                                .Data1 = ScriptNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                        End With
                    End If
                Else
                    For y2 = 0 To Int(frmMapEditor.shpSelected.Height / PIC_Y) - 1
                        For x2 = 0 To Int(frmMapEditor.shpSelected.Width / PIC_X) - 1
                            If x1 + x2 <= MAX_MAPX Then
                                If y1 + y2 <= MAX_MAPY Then
                                    If frmMapEditor.optTiles = True Then
                                        With Map(GetPlayerMap(MyIndex)).Tile(x1 + x2, y1 + y2)
                                            If frmMapEditor.optGround.Value = True Then
                                                .Ground = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .GroundSet = EditorSet
                                            End If
                                            If frmMapEditor.optMask.Value = True Then
                                                .Mask = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .MaskSet = EditorSet
                                            End If
                                            If frmMapEditor.optAnim.Value = True Then
                                                .Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .AnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optMask2.Value = True Then
                                                .Mask2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Mask2Set = EditorSet
                                            End If
                                            If frmMapEditor.optM2Anim.Value = True Then
                                                .M2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .M2AnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optFringe.Value = True Then
                                                .Fringe = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FringeSet = EditorSet
                                            End If
                                            If frmMapEditor.optFAnim.Value = True Then
                                                .FAnim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FAnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optFringe2.Value = True Then
                                                .Fringe2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Fringe2Set = EditorSet
                                            End If
                                            If frmMapEditor.optF2Anim.Value = True Then
                                                .F2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .F2AnimSet = EditorSet
                                            End If
                                        End With
                                    ElseIf frmMapEditor.optlight.Value = True Then
                                        Map(GetPlayerMap(MyIndex)).Tile(x1 + x2, y1 + y2).Light = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                    End If
                                End If
                            End If
                        Next x2
                    Next y2
                End If
            End If
           
            If (Button = 2) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmMapEditor.optTiles.Value = True Then
                    With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                        If frmMapEditor.optGround.Value = True Then
                            .Ground = 0
                            .GroundSet = -1
                        End If
                        If frmMapEditor.optMask.Value = True Then
                            .Mask = 0
                            .MaskSet = -1
                        End If
                        If frmMapEditor.optAnim.Value = True Then
                            .Anim = 0
                            .AnimSet = -1
                        End If
                        If frmMapEditor.optMask2.Value = True Then
                            .Mask2 = 0
                            .Mask2Set = -1
                        End If
                        If frmMapEditor.optM2Anim.Value = True Then
                            .M2Anim = 0
                            .M2AnimSet = -1
                        End If
                        If frmMapEditor.optFringe.Value = True Then
                            .Fringe = 0
                            .FringeSet = -1
                        End If
                        If frmMapEditor.optFAnim.Value = True Then
                            .FAnim = 0
                            .FAnimSet = -1
                        End If
                        If frmMapEditor.optFringe2.Value = True Then
                            .Fringe2 = 0
                            .Fringe2Set = -1
                        End If
                        If frmMapEditor.optF2Anim.Value = True Then
                            .F2Anim = 0
                            .F2AnimSet = -1
                        End If
                    End With
                ElseIf frmMapEditor.optlight.Value = True Then
                    Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light = 0
                ElseIf frmMapEditor.optAttributes.Value = True Then
                    With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                        .Type = 0
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End With
                    
                End If
            End If
        End If
    End If
End Sub

Public Sub PetMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, y1 As Long
    If Player(MyIndex).Pet.Alive = NO Then Exit Sub
    
    x1 = Int(x / PIC_X)
    y1 = Int(y / PIC_Y)
    
    If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        Call SendData(PETMOVESELECT_CHAR & SEP_CHAR & x1 & SEP_CHAR & y1 & END_CHAR)
    End If
End Sub

Public Sub EditorChooseTile(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        EditorTileX = Int(x / PIC_X)
        EditorTileY = Int(y / PIC_Y)
    End If
    frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
    frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
    'Call BitBlt(frmMapEditor.picSelect.hDC, 0, 0, PIC_X, PIC_Y, frmMapEditor.picBackSelect.hDC, EditorTileX * PIC_X, EditorTileY * PIC_Y, SRCCOPY)
End Sub

Public Sub EditorTileScroll()
    frmMapEditor.picBackSelect.Top = (frmMapEditor.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    InEditor = False
    InSpawnEditor = False
    frmMapEditor.Visible = False
    frmMapProperties.Visible = False
    'frmAttributes.Visible = False
    frmMirage.Show
    frmadmin.Show
    
    If CurrentSong = vbNullString And Trim$(Map(GetPlayerMap(MyIndex)).Music) <> vbNullString Then
        Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
    End If
    
    frmMapEditor.MousePointer = 1
    frmMirage.MousePointer = 1
    Call LoadMap(GetPlayerMap(MyIndex))
    'frmMirage.picMapEditor.Visible = False
End Sub

Public Sub EditorClearLayer()
Dim YesNo As Long, x As Long, y As Long

    ' Ground layer
    If frmMapEditor.optGround.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the ground layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Ground = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).GroundSet = -1
                Next x
            Next y
        End If
    End If

    ' Mask layer
    If frmMapEditor.optMask.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).MaskSet = -1
                Next x
            Next y
        End If
    End If
   
    ' Mask Animation layer
    If frmMapEditor.optAnim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).AnimSet = -1
                Next x
            Next y
        End If
    End If
   
    ' Mask 2 layer
    If frmMapEditor.optMask2.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask 2 layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2Set = -1
                Next x
            Next y
        End If
    End If
   
    ' Mask 2 Animation layer
    If frmMapEditor.optM2Anim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask 2 animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).M2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).M2AnimSet = -1
                Next x
            Next y
        End If
    End If
   
    ' Fringe layer
    If frmMapEditor.optFringe.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FringeSet = -1
                Next x
            Next y
        End If
    End If
   
    ' Fringe Animation layer
    If frmMapEditor.optFAnim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnimSet = -1
                Next x
            Next y
        End If
    End If
   
    ' Fringe 2 layer
    If frmMapEditor.optFringe2.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe 2 layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2Set = -1
                Next x
            Next y
        End If
    End If
   
    ' Fringe 2 Animation layer
    If frmMapEditor.optF2Anim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe 2 animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).F2AnimSet = -1
                Next x
            Next y
        End If
    End If
End Sub

Public Sub EditorClearMap()
Dim YesNo As Long, x As Long, y As Long

    YesNo = MsgBox("Are you sure you wish to clear the whole map?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Ground = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).GroundSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).MaskSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Anim = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).AnimSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2 = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2Set = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).M2Anim = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).M2AnimSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).FringeSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnim = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnimSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2 = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2Set = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).F2Anim = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).F2AnimSet = -1
            Next x
        Next y
    End If
End Sub

Public Sub EditorClearAttribs()
Dim YesNo As Long, x As Long, y As Long

    YesNo = MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME)
   
    If YesNo = vbYes Then
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = 0
            Next x
        Next y
    End If
End Sub

Public Sub EmoticonEditorInit()
    frmEmoticonEditor.scrlEmoticon.Max = MAX_EMOTICONS
    frmEmoticonEditor.scrlEmoticon.Value = Emoticons(EditorIndex - 1).Pic
    frmEmoticonEditor.txtCommand.Text = Trim$(Emoticons(EditorIndex - 1).Command)
    frmEmoticonEditor.picEmoticons.Picture = LoadPicture(App.Path & "\GFX\emoticons.bmp")
    If Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_BOTH Or Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_IMAGE Then frmEmoticonEditor.chkPic.Value = 1
    If Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_BOTH Or Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_SOUND Then frmEmoticonEditor.chkSound.Value = 1
    
    frmEmoticonEditor.Show vbModal
End Sub

Public Sub EmoticonEditorOk()
    Emoticons(EditorIndex - 1).Pic = frmEmoticonEditor.scrlEmoticon.Value
    Emoticons(EditorIndex - 1).Sound = frmEmoticonEditor.lstSound.Text
    If frmEmoticonEditor.txtCommand.Text <> "/" Then
        Emoticons(EditorIndex - 1).Command = frmEmoticonEditor.txtCommand.Text
    Else
        Emoticons(EditorIndex - 1).Command = vbNullString
    End If
    If frmEmoticonEditor.chkPic.Value = 1 Then
        If frmEmoticonEditor.chkSound.Value = 1 Then
            Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_BOTH
        Else
            Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_IMAGE
        End If
    Else
        If frmEmoticonEditor.chkSound.Value = 1 Then
            Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_SOUND
        Else
            Call MsgBox("You need to pick either a picture, a sound, or both")
        End If
    End If
    
    Call SendSaveEmoticon(EditorIndex - 1)
    Call EmoticonEditorCancel
End Sub

Public Sub EmoticonEditorCancel()
    InEmoticonEditor = False
    Unload frmEmoticonEditor
End Sub

Public Sub ArrowEditorInit()
    frmEditArrows.scrlArrow.Max = MAX_ARROWS
    If Arrows(EditorIndex).Pic = 0 Then Arrows(EditorIndex).Pic = 1
    frmEditArrows.scrlArrow.Value = Arrows(EditorIndex).Pic
    frmEditArrows.txtName.Text = Arrows(EditorIndex).Name
    If Arrows(EditorIndex).Range = 0 Then Arrows(EditorIndex).Range = 1
    frmEditArrows.scrlRange.Value = Arrows(EditorIndex).Range
    frmEditArrows.picArrows.Picture = LoadPicture(App.Path & "\GFX\arrows.bmp")
    frmEditArrows.Show vbModal
End Sub

Public Sub SpeechEditorInit()
Dim p As Long

    frmSpeech.scrlNumber.Max = MAX_SPEECH_OPTIONS
    frmSpeech.scrlNumber.Value = 0
    frmSpeech.scrlGoTo(0).Max = MAX_SPEECH_OPTIONS
    frmSpeech.scrlGoTo(1).Max = MAX_SPEECH_OPTIONS
    frmSpeech.scrlGoTo(2).Max = MAX_SPEECH_OPTIONS
    frmSpeech.lblSection.Caption = "0"
    frmSpeech.chkQuit.Enabled = False
    frmSpeech.chkScript.Enabled = False
    
    If Trim$(Speech(EditorIndex).Name) = vbNullString Then
        frmSpeech.lblWarn.Visible = True
    Else
        frmSpeech.lblWarn.Visible = False
    End If
    
    frmSpeech.txtName.Text = Speech(EditorIndex).Name
    frmSpeech.chkQuit.Value = Speech(EditorIndex).Num(0).Exit
    frmSpeech.txtMainTalk.Text = Speech(EditorIndex).Num(0).Text
    frmSpeech.optSaid(Speech(EditorIndex).Num(0).SaidBy).Value = True
    If Speech(EditorIndex).Num(0).Respond > 0 Then
        frmSpeech.chkRespond.Value = 1
    Else
        frmSpeech.chkRespond.Value = 0
    End If
    
    If Speech(EditorIndex).Num(0).Script > 0 Then
        frmSpeech.chkScript.Value = 1
        frmSpeech.scrlScript.Visible = True
        frmSpeech.scrlScript.Value = Speech(EditorIndex).Num(0).Script
        frmSpeech.lblScript.Visible = True
        frmSpeech.lblScript.Caption = Speech(EditorIndex).Num(0).Script
    Else
        frmSpeech.chkScript.Value = 0
        frmSpeech.scrlScript.Visible = False
        frmSpeech.scrlScript.Value = 0
        frmSpeech.lblScript.Visible = False
        frmSpeech.lblScript.Caption = 0
    End If
    
    For p = 1 To 3
        If frmSpeech.chkRespond.Value = 1 Then
            frmSpeech.optResponces(p - 1).Enabled = True
            frmSpeech.txtTalk(p - 1).Enabled = True
            frmSpeech.scrlGoTo(p - 1).Enabled = True
            frmSpeech.lblGoTo(p - 1).Enabled = True
            frmSpeech.chkExit(p - 1).Enabled = True
            
            If Speech(EditorIndex).Num(0).Respond = p Then
                frmSpeech.optResponces(p - 1).Value = True
            End If
        
            frmSpeech.txtTalk(p - 1).Text = Speech(EditorIndex).Num(0).Responces(p).Text
            frmSpeech.scrlGoTo(p - 1).Value = Speech(EditorIndex).Num(0).Responces(p).GoTo
            frmSpeech.lblGoTo(p - 1).Caption = "Go to " & Speech(EditorIndex).Num(0).Responces(p).GoTo
            frmSpeech.chkExit(p - 1).Value = Speech(EditorIndex).Num(0).Responces(p).Exit
        Else
            frmSpeech.optResponces(p - 1).Enabled = False
            frmSpeech.txtTalk(p - 1).Enabled = False
            frmSpeech.scrlGoTo(p - 1).Enabled = False
            frmSpeech.lblGoTo(p - 1).Enabled = False
            frmSpeech.chkExit(p - 1).Enabled = False
        End If
    Next p
    
    SpeechEditorCurrentNumber = 0
    
    frmSpeech.Show vbModal
End Sub

Public Sub ArrowEditorOk()
    Arrows(EditorIndex).Pic = frmEditArrows.scrlArrow.Value
    Arrows(EditorIndex).Range = frmEditArrows.scrlRange.Value
    Arrows(EditorIndex).Name = frmEditArrows.txtName.Text
    Call SendSaveArrow(EditorIndex)
    Call ArrowEditorCancel
End Sub

Public Sub ArrowEditorCancel()
    InArrowEditor = False
    Unload frmEditArrows
End Sub

Public Sub ItemEditorInit()
Dim I As Long
    EditorItemY = Int(Item(EditorIndex).Pic / 6)
    EditorItemX = (Item(EditorIndex).Pic - Int(Item(EditorIndex).Pic / 6) * 6)
    
    frmItemEditor.scrlClassReq.Max = Max_Classes

    frmItemEditor.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
    
    frmItemEditor.txtName.Text = Trim$(Item(EditorIndex).Name)
    frmItemEditor.txtDesc.Text = Trim$(Item(EditorIndex).desc)
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        frmItemEditor.fraBow.Visible = True
        
        If Item(EditorIndex).Data1 >= 0 Then
            frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        Else
            frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1 * -1
        End If
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlMagicReq.Value = Item(EditorIndex).MagicReq
        frmItemEditor.scrlClassReq.Value = Item(EditorIndex).ClassReq
        frmItemEditor.scrlAccessReq.Value = Item(EditorIndex).AccessReq
        frmItemEditor.scrlAddHP.Value = Item(EditorIndex).AddHP
        frmItemEditor.scrlAddMP.Value = Item(EditorIndex).AddMP
        frmItemEditor.scrlAddSP.Value = Item(EditorIndex).AddSP
        frmItemEditor.scrlAddStr.Value = Item(EditorIndex).AddStr
        frmItemEditor.scrlAddDef.Value = Item(EditorIndex).AddDef
        frmItemEditor.scrlAddMagi.Value = Item(EditorIndex).AddMagi
        frmItemEditor.scrlAddSpeed.Value = Item(EditorIndex).AddSpeed
        frmItemEditor.scrlAddEXP.Value = Item(EditorIndex).AddEXP
        frmItemEditor.scrlAttackSpeed.Value = Item(EditorIndex).AttackSpeed
        
        If Item(EditorIndex).Data3 > 0 Then
            frmItemEditor.chkBow.Value = Checked
        Else
            frmItemEditor.chkBow.Value = Unchecked
        End If
        
        If Item(EditorIndex).Data1 > 0 Then
            frmItemEditor.chkRepair.Value = Checked
        Else
            frmItemEditor.chkRepair.Value = Unchecked
        End If
        
        frmItemEditor.cmbBow.Clear
        If frmItemEditor.chkBow.Value = Checked Then
            For I = 1 To 100
                frmItemEditor.cmbBow.AddItem I & ": " & Arrows(I).Name
            Next I
            frmItemEditor.cmbBow.ListIndex = Item(EditorIndex).Data3 - 1
            frmItemEditor.picBow.Top = (Arrows(Item(EditorIndex).Data3).Pic * 32) * -1
            frmItemEditor.cmbBow.Enabled = True
        Else
            frmItemEditor.cmbBow.AddItem "None"
            frmItemEditor.cmbBow.ListIndex = 0
            frmItemEditor.cmbBow.Enabled = False
        End If
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.scrlVitalMod.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraSpell.Visible = False
    End If

    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).Name = frmItemEditor.txtName.Text
    Item(EditorIndex).desc = frmItemEditor.txtDesc.Text
    Item(EditorIndex).Pic = EditorItemY * 6 + EditorItemX
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        If frmItemEditor.chkRepair.Value = 0 Then Item(EditorIndex).Data1 = Item(EditorIndex).Data1 * -1
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        If frmItemEditor.chkBow.Value = Checked Then
            Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
        Else
            Item(EditorIndex).Data3 = 0
        End If
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        Item(EditorIndex).MagicReq = frmItemEditor.scrlMagicReq.Value
        
        Item(EditorIndex).ClassReq = frmItemEditor.scrlClassReq.Value
        Item(EditorIndex).AccessReq = frmItemEditor.scrlAccessReq.Value
        
        Item(EditorIndex).AddHP = frmItemEditor.scrlAddHP.Value
        Item(EditorIndex).AddMP = frmItemEditor.scrlAddMP.Value
        Item(EditorIndex).AddSP = frmItemEditor.scrlAddSP.Value
        Item(EditorIndex).AddStr = frmItemEditor.scrlAddStr.Value
        Item(EditorIndex).AddDef = frmItemEditor.scrlAddDef.Value
        Item(EditorIndex).AddMagi = frmItemEditor.scrlAddMagi.Value
        Item(EditorIndex).AddSpeed = frmItemEditor.scrlAddSpeed.Value
        Item(EditorIndex).AddEXP = frmItemEditor.scrlAddEXP.Value
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.Value
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = 0
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = 0
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_PET) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlPet.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlPetLevel.Value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = 0
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
    End If
    
    Call SendSaveItem(EditorIndex)
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub ItemEditorCancel()
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub NpcEditorInit()
    
    frmNpcEditor.Picsprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
    
    frmNpcEditor.txtName.Text = Trim$(Npc(EditorIndex).Name)
    frmNpcEditor.txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = STR(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    If Npc(EditorIndex).Range = 0 Then Npc(EditorIndex).Range = 1
    frmNpcEditor.scrlRange.Value = Npc(EditorIndex).Range
    frmNpcEditor.scrlSTR.Value = Npc(EditorIndex).STR
    frmNpcEditor.scrlDEF.Value = Npc(EditorIndex).DEF
    frmNpcEditor.scrlSPEED.Value = Npc(EditorIndex).Speed
    frmNpcEditor.scrlMAGI.Value = Npc(EditorIndex).MAGI
    frmNpcEditor.BigNpc.Value = Npc(EditorIndex).Big
    If Npc(EditorIndex).MaxHP = 0 Then
        frmNpcEditor.StartHP.Value = 1
    Else
        frmNpcEditor.StartHP.Value = Npc(EditorIndex).MaxHP
    End If
    frmNpcEditor.ExpGive.Text = Npc(EditorIndex).EXP
    frmNpcEditor.txtChance.Text = STR(Npc(EditorIndex).ItemNPC(1).Chance)
    frmNpcEditor.scrlNum.Value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmNpcEditor.scrlValue.Value = Npc(EditorIndex).ItemNPC(1).ItemValue
    frmNpcEditor.scrlSpeech.Value = Npc(EditorIndex).Speech
    If Npc(EditorIndex).Speech > 0 Then
        frmNpcEditor.lblSpeechName.Caption = Speech(Npc(EditorIndex).Speech).Name
    Else
        frmNpcEditor.lblSpeechName.Caption = vbNullString
    End If
    If Npc(EditorIndex).SpawnTime = 0 Then
        frmNpcEditor.chkDay.Value = Checked
        frmNpcEditor.chkNight.Value = Checked
    ElseIf Npc(EditorIndex).SpawnTime = 1 Then
        frmNpcEditor.chkDay.Value = Checked
        frmNpcEditor.chkNight.Value = Unchecked
    ElseIf Npc(EditorIndex).SpawnTime = 2 Then
        frmNpcEditor.chkDay.Value = Unchecked
        frmNpcEditor.chkNight.Value = Checked
    End If
    
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).Name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmNpcEditor.scrlSprite.Value
    Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.Value
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.Value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.Value
    Npc(EditorIndex).Speed = frmNpcEditor.scrlSPEED.Value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.Value
    Npc(EditorIndex).Big = frmNpcEditor.BigNpc.Value
    Npc(EditorIndex).MaxHP = frmNpcEditor.StartHP.Value
    Npc(EditorIndex).EXP = frmNpcEditor.ExpGive.Text
    Npc(EditorIndex).Speech = frmNpcEditor.scrlSpeech.Value
    
    If frmNpcEditor.chkDay.Value = Checked And frmNpcEditor.chkNight.Value = Checked Then
        Npc(EditorIndex).SpawnTime = 0
    ElseIf frmNpcEditor.chkDay.Value = Checked And frmNpcEditor.chkNight.Value = Unchecked Then
        Npc(EditorIndex).SpawnTime = 1
    ElseIf frmNpcEditor.chkDay.Value = Unchecked And frmNpcEditor.chkNight.Value = Checked Then
        Npc(EditorIndex).SpawnTime = 2
    End If
    
    Call SendSaveNpc(EditorIndex)
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorCancel()
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorBltSprite()
    If frmNpcEditor.BigNpc.Value = Checked Then
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, 64, 64, frmNpcEditor.Picsprites.hDC, 3 * 64, frmNpcEditor.scrlSprite.Value * 64, SRCCOPY)
    Else
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, SIZE_X, SIZE_Y, frmNpcEditor.Picsprites.hDC, 3 * SIZE_X, frmNpcEditor.scrlSprite.Value * SIZE_Y - (SIZE_Y - PIC_Y), SRCCOPY)
    End If
End Sub

Public Sub ShopEditorInit()
Dim I As Long

    frmShopEditor.txtName.Text = Trim$(Shop(EditorIndex).Name)
    frmShopEditor.txtJoinSay.Text = Trim$(Shop(EditorIndex).JoinSay)
    frmShopEditor.txtLeaveSay.Text = Trim$(Shop(EditorIndex).LeaveSay)
    frmShopEditor.chkFixesItems.Value = Shop(EditorIndex).FixesItems
    
    frmShopEditor.cmbItemGive.Clear
    frmShopEditor.cmbItemGive.AddItem "None"
    frmShopEditor.cmbItemGet.Clear
    frmShopEditor.cmbItemGet.AddItem "None"
    For I = 1 To MAX_ITEMS
        frmShopEditor.cmbItemGive.AddItem I & ": " & Trim$(Item(I).Name)
        frmShopEditor.cmbItemGet.AddItem I & ": " & Trim$(Item(I).Name)
    Next I
    frmShopEditor.cmbItemGive.ListIndex = 0
    frmShopEditor.cmbItemGet.ListIndex = 0
    
    Call UpdateShopTrade
    
    frmShopEditor.Show vbModal
End Sub

Public Sub UpdateShopTrade()
Dim I As Long, GetItem As Long, GetValue As Long, GiveItem As Long, GiveValue As Long, C As Long
    
    For I = 0 To 5
        frmShopEditor.lstTradeItem(I).Clear
    Next I
    
    For C = 1 To 6
        For I = 1 To MAX_TRADES
            GetItem = Shop(EditorIndex).TradeItem(C).Value(I).GetItem
            GetValue = Shop(EditorIndex).TradeItem(C).Value(I).GetValue
            GiveItem = Shop(EditorIndex).TradeItem(C).Value(I).GiveItem
            GiveValue = Shop(EditorIndex).TradeItem(C).Value(I).GiveValue

            If GetItem > 0 And GiveItem > 0 Then
                frmShopEditor.lstTradeItem(C - 1).AddItem I & ": " & GiveValue & " " & Trim$(Item(GiveItem).Name) & " for " & GetValue & " " & Trim$(Item(GetItem).Name)
            Else
                frmShopEditor.lstTradeItem(C - 1).AddItem "Empty Trade Slot"
            End If
        Next I
    Next C
    
    For I = 0 To 5
        frmShopEditor.lstTradeItem(I).ListIndex = 0
    Next I
End Sub

Public Sub ShopEditorOk()
    Shop(EditorIndex).Name = frmShopEditor.txtName.Text
    Shop(EditorIndex).JoinSay = frmShopEditor.txtJoinSay.Text
    Shop(EditorIndex).LeaveSay = frmShopEditor.txtLeaveSay.Text
    Shop(EditorIndex).FixesItems = frmShopEditor.chkFixesItems.Value
    
    Call SendSaveShop(EditorIndex)
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub SpellEditorInit()
Dim I As Long

    frmSpellEditor.cmbClassReq.AddItem "All Classes"
    For I = 1 To Max_Classes
        frmSpellEditor.cmbClassReq.AddItem Trim$(Class(I).Name)
    Next I
    
    frmSpellEditor.txtName.Text = Trim$(Spell(EditorIndex).Name)
    frmSpellEditor.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmSpellEditor.scrlLevelReq.Value = Spell(EditorIndex).LevelReq
        
    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    frmSpellEditor.scrlVitalMod.Value = Spell(EditorIndex).Data1
    
    frmSpellEditor.scrlCost.Value = Spell(EditorIndex).MPCost
    frmSpellEditor.scrlSound.Value = Spell(EditorIndex).Sound
    
    If Spell(EditorIndex).Range = 0 Then Spell(EditorIndex).Range = 1
    frmSpellEditor.scrlRange.Value = Spell(EditorIndex).Range
    
    frmSpellEditor.scrlSpellAnim.Value = Spell(EditorIndex).SpellAnim
    If Spell(EditorIndex).SpellTime < 40 Then Spell(EditorIndex).SpellTime = 40
    frmSpellEditor.scrlSpellTime.Value = Spell(EditorIndex).SpellTime
    If Spell(EditorIndex).SpellDone = 0 Then Spell(EditorIndex).SpellDone = 1
    frmSpellEditor.scrlSpellDone.Value = Spell(EditorIndex).SpellDone
    
    frmSpellEditor.chkArea.Value = Spell(EditorIndex).AE
        
    frmSpellEditor.Show vbModal
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).Name = frmSpellEditor.txtName.Text
    Spell(EditorIndex).ClassReq = frmSpellEditor.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevelReq.Value
    Spell(EditorIndex).Type = frmSpellEditor.cmbType.ListIndex
    Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.Value
    Spell(EditorIndex).Data3 = 0
    Spell(EditorIndex).MPCost = frmSpellEditor.scrlCost.Value
    Spell(EditorIndex).Sound = frmSpellEditor.scrlSound.Value
    Spell(EditorIndex).Range = frmSpellEditor.scrlRange.Value
    
    Spell(EditorIndex).SpellAnim = frmSpellEditor.scrlSpellAnim.Value
    Spell(EditorIndex).SpellTime = frmSpellEditor.scrlSpellTime.Value
    Spell(EditorIndex).SpellDone = frmSpellEditor.scrlSpellDone.Value
    
    Spell(EditorIndex).AE = frmSpellEditor.chkArea.Value
    
    Call SendSaveSpell(EditorIndex)
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub UpdateTradeInventory()
Dim I As Long

    frmPlayerTrade.PlayerInv1.Clear
    
For I = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, I) > 0 And GetPlayerInvItemNum(MyIndex, I) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Then
            frmPlayerTrade.PlayerInv1.AddItem I & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name) & " (" & GetPlayerInvItemValue(MyIndex, I) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = I Or GetPlayerArmorSlot(MyIndex) = I Or GetPlayerHelmetSlot(MyIndex) = I Or GetPlayerShieldSlot(MyIndex) = I Then
                frmPlayerTrade.PlayerInv1.AddItem I & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name) & " (worn)"
            Else
                frmPlayerTrade.PlayerInv1.AddItem I & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name)
            End If
        End If
    Else
        frmPlayerTrade.PlayerInv1.AddItem "<Nothing>"
    End If
Next I
    
    frmPlayerTrade.PlayerInv1.ListIndex = 0
End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1 As Long, y1 As Long

    x1 = Int(x / PIC_X)
    y1 = Int(y / PIC_Y)
    
    If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        Call SendData(SEARCH_CHAR & SEP_CHAR & x1 & SEP_CHAR & y1 & END_CHAR)
    End If
    MouseDownX = x1
    MouseDownY = y1
End Sub
Sub BltTile2(ByVal x As Long, ByVal y As Long, ByVal Tile As Long)
If TileFile(6) = 0 Then Exit Sub

    rec.Top = Int(Tile / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_TileSurf(6), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerText(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim intLoop As Integer
Dim intLoop2 As Integer

Dim bytLineCount As Byte
Dim bytLineLength As Byte
Dim strLine(0 To MAX_LINES - 1) As String
Dim strWords() As String
    strWords() = Split(Bubble(Index).Text, " ")
    
    If Len(Bubble(Index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(Bubble(Index).Text) * 9) \ PIC_X)
        
        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then
            DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
        End If
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If
    
    TextX = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - Int(PIC_Y) + 85
    
    If TextX < ((DISPLAY_BUBBLE_WIDTH * 32) / 2) Then TextX = ((DISPLAY_BUBBLE_WIDTH * 32) / 2)
    If TextX > (MAX_MAPX * PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) Then TextX = (MAX_MAPX * PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2)
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 1)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 2)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 16)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1
            
        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1
            
            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If
            
            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    If bytLineCount > 0 Then
        For intLoop = 6 To (bytLineCount - 2) * PIC_Y + 6
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 17)
            
            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 5)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 16) - 4, 3)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 16) - 4, 4)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) + sx - NewXOffset + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 7, TextY - (NewPlayerY * PIC_Y) + sx - NewYOffset, strLine(intLoop), QBColor(DarkGrey))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub
Sub BltPlayerBars(ByVal Index As Long)
Dim x As Long, y As Long

    x = (GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
    y = (GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
    
    If Player(Index).HP = 0 Then Exit Sub
    'draws the back bars
    Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
    Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
    
    'draws HP
    If Int((GetPlayerHP(Index) / GetPlayerMaxHP(Index)) * 100) > 50 Then
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    End If
    If Int((GetPlayerHP(Index) / GetPlayerMaxHP(Index)) * 100) > 20 And Int((GetPlayerHP(Index) / GetPlayerMaxHP(Index)) * 100) <= 50 Then
        Call DD_BackBuffer.SetFillColor(RGB(255, 255, 0))
    End If
    If Int((GetPlayerHP(Index) / GetPlayerMaxHP(Index)) * 100) <= 20 Then
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    End If
    Call DD_BackBuffer.DrawBox(x, y + PIC_Y, x + ((Player(Index).HP / 100) / (Player(Index).MaxHP / 100) * SIZE_X), y + 36)
End Sub
Sub BltPetBars(ByVal Index As Long)
Dim x As Long, y As Long

    x = (Player(Index).Pet.x * PIC_X + sx + Player(Index).Pet.XOffset) - (NewPlayerX * PIC_X) - NewXOffset
    y = (Player(Index).Pet.y * PIC_Y + sx + Player(Index).Pet.YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
    
    If Player(Index).HP = 0 Then Exit Sub
    'draws the back bars
    Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
    Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
    
    'draws HP
    If Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) > 50 Then
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    End If
    If Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) > 20 And Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) <= 50 Then
        Call DD_BackBuffer.SetFillColor(RGB(255, 255, 0))
    End If
    If Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) <= 20 Then
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    End If
    Call DD_BackBuffer.DrawBox(x, y + PIC_Y, x + ((Player(Index).Pet.HP / 100) / (Player(Index).Pet.MaxHP / 100) * SIZE_X), y + 36)
End Sub
Sub BltNpcBars(ByVal Index As Long)
Dim x As Long, y As Long

If MapNpc(Index).HP = 0 Then Exit Sub
If MapNpc(Index).Num < 1 Then Exit Sub

    If Npc(MapNpc(Index).Num).Big = 1 Then
        x = (MapNpc(Index).x * PIC_X + sx - 9 + MapNpc(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
        Call DD_BackBuffer.DrawBox(x, y + 32, x + 50, y + 36)
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) > 50 Then
            Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        End If
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) <= 50 And Int((MapNpc(Index).HP / MapNpc(Index).MaxHP) * 100) > 20 Then
            Call DD_BackBuffer.SetFillColor(RGB(255, 255, 0))
        End If
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) <= 20 Then
            Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        End If
        Call DD_BackBuffer.DrawBox(x, y + 32, x + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHP / 100) * 50), y + 36)
    Else
        x = (MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
        Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) > 50 Then
            Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        End If
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) <= 50 And Int((MapNpc(Index).HP / MapNpc(Index).MaxHP) * 100) > 20 Then
            Call DD_BackBuffer.SetFillColor(RGB(255, 255, 0))
        End If
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) <= 20 Then
            Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        End If
        Call DD_BackBuffer.DrawBox(x, y + 32, x + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHP / 100) * 32), y + 36)
    End If
End Sub

Public Sub UpdateVisInv()
Dim Index As Long
Dim d As Long

    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) <> Index Then frmMirage.ShieldImage.Picture = LoadPicture()
        If GetPlayerWeaponSlot(MyIndex) <> Index Then frmMirage.WeaponImage.Picture = LoadPicture()
        If GetPlayerHelmetSlot(MyIndex) <> Index Then frmMirage.HelmetImage.Picture = LoadPicture()
        If GetPlayerArmorSlot(MyIndex) <> Index Then frmMirage.ArmorImage.Picture = LoadPicture()
    Next Index
    
    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ShieldImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerWeaponSlot(MyIndex) = Index Then Call BitBlt(frmMirage.WeaponImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerHelmetSlot(MyIndex) = Index Then Call BitBlt(frmMirage.HelmetImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerArmorSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ArmorImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
    Next Index
        
    frmMirage.EquipS(0).Visible = False
    frmMirage.EquipS(1).Visible = False
    frmMirage.EquipS(2).Visible = False
    frmMirage.EquipS(3).Visible = False

    For d = 0 To MAX_INV - 1
        If Player(MyIndex).Inv(d + 1).Num > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type <> ITEM_TYPE_CURRENCY Then
                'frmMirage.descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
            'Else
                If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(0).Visible = True
                    frmMirage.EquipS(0).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(0).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(1).Visible = True
                    frmMirage.EquipS(1).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(1).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(2).Visible = True
                    frmMirage.EquipS(2).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(2).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(3).Visible = True
                    frmMirage.EquipS(3).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(3).Left = frmMirage.picInv(d).Left - 2
                Else
                    'frmMirage.picInv(d).ToolTipText = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
                End If
            End If
        End If
    Next d
End Sub

Sub BltSpriteChange(ByVal x As Long, ByVal y As Long)
Dim x2 As Long, y2 As Long
    
    If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_SPRITE_CHANGE Then

        ' Only used if ever want to switch to blt rather then bltfast
        With rec_pos
            .Top = y * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = x * PIC_X
            .Right = .Left + PIC_X
        End With
                                        
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * PIC_Y + 16
        rec.Bottom = rec.Top + PIC_Y - 16
        rec.Left = 96
        rec.Right = rec.Left + PIC_X
        
        x2 = x * PIC_X + sx
        y2 = y * PIC_Y + sx
                                           
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltSpriteChange2(ByVal x As Long, ByVal y As Long)
Dim x2 As Long, y2 As Long
    
    If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_SPRITE_CHANGE Then

        ' Only used if ever want to switch to blt rather then bltfast
        With rec_pos
            .Top = y * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = x * PIC_X
            .Right = .Left + PIC_X
        End With
                                        
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * PIC_Y
        rec.Bottom = rec.Top + PIC_Y - 16
        rec.Left = 96
        rec.Right = rec.Left + PIC_X
        
        x2 = x * PIC_X + sx
        y2 = y * PIC_Y + (sx / 2) '- 16
               
        If y2 < 0 Then
            Exit Sub
        End If
                            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub SendGameTime()
Dim Packet As String

Packet = GMTIME_CHAR & SEP_CHAR & GameTime & END_CHAR
Call SendData(Packet)
End Sub

Sub ItemSelected(ByVal Index As Long, ByVal Selected As Long)
Dim index2 As Long
index2 = Trade(Selected).Items(Index).ItemGetNum

    frmTrade.shpSelect.Top = frmTrade.picItem(Index - 1).Top - 1
    frmTrade.shpSelect.Left = frmTrade.picItem(Index - 1).Left - 1

    If index2 <= 0 Then
        Call clearItemSelected
        Exit Sub
    End If

    frmTrade.descName.Caption = Trim$(Item(index2).Name)
    frmTrade.descQuantity.Caption = "Quantity: " & Trade(Selected).Items(Index).ItemGetVal
    
    frmTrade.descStr.Caption = "Strength: " & Item(index2).StrReq
    frmTrade.descDef.Caption = "Defense: " & Item(index2).DefReq
    frmTrade.descSpeed.Caption = "Speed: " & Item(index2).SpeedReq
    frmTrade.descMagi.Caption = "Magic: " & Item(index2).MagicReq
    
    frmTrade.descAStr.Caption = "Strength: " & Item(index2).AddStr
    frmTrade.descADef.Caption = "Defense: " & Item(index2).AddDef
    frmTrade.descAMagi.Caption = "Magic: " & Item(index2).AddMagi
    frmTrade.descASpeed.Caption = "Speed: " & Item(index2).AddSpeed
    
    frmTrade.descHp.Caption = "HP: " & Item(index2).AddHP
    frmTrade.descMp.Caption = "MP: " & Item(index2).AddMP
    frmTrade.descSp.Caption = "SP: " & Item(index2).AddSP
    frmTrade.descAExp.Caption = "EXP: " & Item(index2).AddEXP & "%"
    
    frmTrade.desc.Caption = Trim$(Item(index2).desc)
    
    frmTrade.lblTradeFor.Caption = "Trade for: " & Trim$(Item(Trade(Selected).Items(Index).ItemGiveNum).Name)
    frmTrade.lblQuantity.Caption = "Quantity: " & Trade(Selected).Items(Index).ItemGiveVal
End Sub

Sub clearItemSelected()
    frmTrade.lblTradeFor.Caption = vbNullString
    frmTrade.lblQuantity.Caption = vbNullString
    
    frmTrade.descName.Caption = vbNullString
    frmTrade.descQuantity.Caption = vbNullString
    
    frmTrade.descStr.Caption = "Strength: 0"
    frmTrade.descDef.Caption = "Defense: 0"
    frmTrade.descMagi.Caption = "Magic: 0"
    frmTrade.descSpeed.Caption = "Speed: 0"
    
    frmTrade.descAStr.Caption = "Strength: 0"
    frmTrade.descADef.Caption = "Defense: 0"
    frmTrade.descAMagi.Caption = "Magic: 0"
    frmTrade.descASpeed.Caption = "Speed: 0"
    
    frmTrade.descHp.Caption = "HP: 0"
    frmTrade.descMp.Caption = "MP: 0"
    frmTrade.descSp.Caption = "SP: 0"

    frmTrade.descAExp.Caption = "EXP: 0%"
    frmTrade.desc.Caption = vbNullString
End Sub

Public Sub MakeFormTrans(ByVal Form As Form, ByVal Amount As Long)
Dim NormalWindowStyle As Long

    DoEvents
    
    If GetVersion >= 4 Then ' Make sure it's Windows 2000 or better.
        If Amount > 100 Then Amount = 100 ' Make sure they don't have more then 100 percent
        NormalWindowStyle = GetWindowLong(Form.hWnd, GWL_EXSTYLE)
        SetWindowLong Form.hWnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
        ' Sets the transparency level
        SetLayeredWindowAttributes Form.hWnd, 0, 255 * (1 - (Val(Amount) / 100)), LWA_ALPHA
    End If
End Sub

Public Function MakeLoc(ByVal x As Long, ByVal y As Long) As Long
    MakeLoc = (y * MAX_MAPX) + x
End Function

Public Function MakeX(ByVal Loc As Long) As Long
    MakeX = Loc - (MakeY(Loc) * MAX_MAPX)
End Function

Public Function MakeY(ByVal Loc As Long) As Long
    MakeY = Int(Loc / MAX_MAPX)
End Function

Public Function IsValid(ByVal x As Long, _
   ByVal y As Long) As Boolean
    IsValid = True

    If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then IsValid = False
End Function
