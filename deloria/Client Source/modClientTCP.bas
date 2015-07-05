Attribute VB_Name = "modClientTCP"
Option Explicit

Public ServerIP As String
Public PlayerBuffer As String
Public MapBuffer As String
Public InGame As Boolean
Public ItemGiveS(1 To MAX_TRADES) As Long
Public ItemGetS(1 To MAX_TRADES) As Long
Public ItemGiveSS(1 To MAX_TRADES) As Long
Public ItemGetSS(1 To MAX_TRADES) As Long
Public TradePlayer As Long

Public DebugMode As Boolean

Sub TcpInit()
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    PlayerBuffer = ""
    MapBuffer = ""
    
    Dim FileName As String
    FileName = App.Path & "\config.ini"
    If FileExist("config.ini") Then
        WEBSITE = ReadINI("CONFIG", "WebSite", FileName)
        frmMirage.chkplayername.Value = ReadINI("CONFIG", "PlayerName", FileName)
        frmMapEditor.optMapGrid.Value = ReadINI("CONFIG", "MapGrid", FileName)
        frmMirage.chkmusic.Value = ReadINI("CONFIG", "Music", FileName)
        frmMirage.chksound.Value = ReadINI("CONFIG", "Sound", FileName)
    Else
        WriteINI "CONFIG", "Account", "", (App.Path & "\config.ini")
        WriteINI "CONFIG", "Password", "", (App.Path & "\config.ini")
        WriteINI "CONFIG", "WebSite", "", (App.Path & "\config.ini")
        WriteINI "CONFIG", "PlayerName", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "MapGrid", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "Music", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "Sound", 1, App.Path & "\config.ini"
    End If
    
    FileName = App.Path & "\UC.ini"
    If FileExist("UC.ini") Then
        frmMirage.chkWeather.Value = ReadINI("UC", "Weather", FileName)
        frmMirage.chkTime.Value = ReadINI("UC", "Day/Night", FileName)
    Else
        WriteINI "UC", "Weather", 1, FileName
        WriteINI "UC", "Day/Night", 1, FileName
    End If

    FileName = App.Path & "\Controls.ini"
    If FileExist("Controls.ini") Then
        KUp = Val(ReadINI("KEYBOARD", "Up", FileName))
        KDown = Val(ReadINI("KEYBOARD", "Down", FileName))
        KLeft = Val(ReadINI("KEYBOARD", "Left", FileName))
        KRight = Val(ReadINI("KEYBOARD", "Right", FileName))
        KAttack = Val(ReadINI("KEYBOARD", "Attack", FileName))
        KRun = Val(ReadINI("KEYBOARD", "Run", FileName))
        KEnter = Val(ReadINI("KEYBOARD", "Enter", FileName))
        
        JUp = Val(ReadINI("JOYPAD", "Up", FileName))
        JDown = Val(ReadINI("JOYPAD", "Down", FileName))
        JLeft = Val(ReadINI("JOYPAD", "Left", FileName))
        JRight = Val(ReadINI("JOYPAD", "Right", FileName))
        JAttack = Val(ReadINI("JOYPAD", "Attack", FileName))
        JRun = Val(ReadINI("JOYPAD", "Run", FileName))
        JEnter = Val(ReadINI("JOYPAD", "Enter", FileName))
        JUpC = Val(ReadINI("JOYPAD", "UpC", FileName))
        JDownC = Val(ReadINI("JOYPAD", "DownC", FileName))
        JLeftC = Val(ReadINI("JOYPAD", "LeftC", FileName))
        JRightC = Val(ReadINI("JOYPAD", "RightC", FileName))
        JAttackC = Val(ReadINI("JOYPAD", "AttackC", FileName))
        JRunC = Val(ReadINI("JOYPAD", "RunC", FileName))
        JEnterC = Val(ReadINI("JOYPAD", "EnterC", FileName))
    Else
        WriteINI "KEYBOARD", "Up", 0, FileName
        WriteINI "KEYBOARD", "Down", 0, FileName
        WriteINI "KEYBOARD", "Left", 0, FileName
        WriteINI "KEYBOARD", "Right", 0, FileName
        WriteINI "KEYBOARD", "Attack", 0, FileName
        WriteINI "KEYBOARD", "Run", 0, FileName
        WriteINI "KEYBOARD", "Enter", 0, FileName
        
        WriteINI "JOYPAD", "Up", 0, FileName
        WriteINI "JOYPAD", "Down", 65535, FileName
        WriteINI "JOYPAD", "Left", 0, FileName
        WriteINI "JOYPAD", "Right", 65535, FileName
        WriteINI "JOYPAD", "Attack", 128, FileName
        WriteINI "JOYPAD", "Run", 64, FileName
        WriteINI "JOYPAD", "Enter", 4, FileName
        WriteINI "JOYPAD", "UpC", 2, FileName
        WriteINI "JOYPAD", "DownC", 2, FileName
        WriteINI "JOYPAD", "LeftC", 1, FileName
        WriteINI "JOYPAD", "RightC", 1, FileName
        WriteINI "JOYPAD", "AttackC", 3, FileName
        WriteINI "JOYPAD", "RunC", 3, FileName
        WriteINI "JOYPAD", "EnterC", 3, FileName
        
        KUp = Val(ReadINI("KEYBOARD", "Up", FileName))
        KDown = Val(ReadINI("KEYBOARD", "Down", FileName))
        KLeft = Val(ReadINI("KEYBOARD", "Left", FileName))
        KRight = Val(ReadINI("KEYBOARD", "Right", FileName))
        KAttack = Val(ReadINI("KEYBOARD", "Attack", FileName))
        KRun = Val(ReadINI("KEYBOARD", "Run", FileName))
        KEnter = Val(ReadINI("KEYBOARD", "Enter", FileName))
        
        JUp = Val(ReadINI("JOYPAD", "Up", FileName))
        JDown = Val(ReadINI("JOYPAD", "Down", FileName))
        JLeft = Val(ReadINI("JOYPAD", "Left", FileName))
        JRight = Val(ReadINI("JOYPAD", "Right", FileName))
        JAttack = Val(ReadINI("JOYPAD", "Attack", FileName))
        JRun = Val(ReadINI("JOYPAD", "Run", FileName))
        JEnter = Val(ReadINI("JOYPAD", "Enter", FileName))
        JUpC = Val(ReadINI("JOYPAD", "UpC", FileName))
        JDownC = Val(ReadINI("JOYPAD", "DownC", FileName))
        JLeftC = Val(ReadINI("JOYPAD", "LeftC", FileName))
        JRightC = Val(ReadINI("JOYPAD", "RightC", FileName))
        JAttackC = Val(ReadINI("JOYPAD", "AttackC", FileName))
        JRunC = Val(ReadINI("JOYPAD", "RunC", FileName))
        JEnterC = Val(ReadINI("JOYPAD", "EnterC", FileName))
    End If
    
    FileName = App.Path & "\Input.ini"
    If FileExist("Input.ini") Then
        If Val(ReadINI("CONFIG", "Joypad", FileName)) = 1 Then
            ID = True
        Else
            ID = False
        End If
    Else
        WriteINI "CONFIG", "Joypad", 0, FileName
        ID = False
    End If
    
    If ID = True Then
        frmMirage.Check1.Value = Checked
    Else
        frmMirage.Check1.Value = Unchecked
    End If
     
    frmMirage.Socket.RemoteHost = GAME_IP
    frmMirage.Socket.RemotePort = GAME_PORT
    
    frmMirage.MapSocket.RemoteHost = GAME_IP
    frmMirage.MapSocket.RemotePort = 5000
End Sub

Sub TcpDestroy()
    frmMirage.Socket.Close
    frmMirage.MapSocket.Close
End Sub

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim Start As Long

    frmMirage.Socket.GetData Buffer, vbString, DataLength
    PlayerBuffer = PlayerBuffer & Buffer
        
    Start = InStr(PlayerBuffer, END_CHAR)
    Do While Start > 0
        Packet = Mid(PlayerBuffer, 1, Start - 1)
        PlayerBuffer = Mid(PlayerBuffer, Start + 1, Len(PlayerBuffer))
        Start = InStr(PlayerBuffer, END_CHAR)
        If Len(Packet) > 0 Then
            Call HandleData(Packet)
        End If
    Loop
End Sub

Function ConnectToServer() As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMirage.Socket.Close
    frmMirage.Socket.Connect
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop
    
    If IsConnected Then
        ConnectToServer = True
    Else
        ConnectToServer = False
    End If
End Function

Function IsConnected() As Boolean
    If frmMirage.Socket.State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If GetPlayerName(Index) <> "" Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Sub SendData(ByVal Data As String)
    If IsConnected Then
        frmMirage.Socket.SendData Data
        DoEvents
    End If
End Sub

Sub SendNewAccount(ByVal name As String, ByVal Password As String)
Dim Packet As String

    Packet = "newaccount" & SEP_CHAR & Trim(name) & SEP_CHAR & Trim(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelAccount(ByVal name As String, ByVal Password As String)
Dim Packet As String
    
    Packet = "delaccount" & SEP_CHAR & Trim(name) & SEP_CHAR & Trim(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLogin(ByVal name As String, ByVal Password As String)
Dim Packet As String

    Call SendData("configuremaps" & SEP_CHAR & ConfigMaps & SEP_CHAR & END_CHAR)
    Packet = "login" & SEP_CHAR & Trim(name) & SEP_CHAR & Trim(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & SEC_CODE1 & SEP_CHAR & SEC_CODE2 & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAddChar(ByVal name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long, ByVal Pic As Long)
Dim Packet As String

    Packet = "addchar" & SEP_CHAR & Trim(name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & Slot & SEP_CHAR & Pic & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelChar(ByVal Slot As Long)
Dim Packet As String
    
    Packet = "delchar" & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetClasses()
Dim Packet As String

    Packet = "getclasses" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseChar(ByVal CharSlot As Long)
Dim Packet As String
Dim i As Long

    frmMainMenu.Enabled = False
    Call SetStatus("Loading maps...")
    For i = 1 To MAX_MAPS
        Call SetStatus("Loading maps... (" & i & "/" & MAX_MAPS & ")")
        Call LoadMaps(i)
        DoEvents
    Next i
    Call SetStatus("Maps loaded!")
    frmMainMenu.Enabled = True

    Packet = "usechar" & SEP_CHAR & CharSlot & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SayMsg(ByVal Text As String)
Dim Packet As String

    Packet = "saymsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub GlobalMsg(ByVal Text As String)
Dim Packet As String

    Packet = "globalmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub BroadcastMsg(ByVal Text As String)
Dim Packet As String

    Packet = "broadcastmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub EmoteMsg(ByVal Text As String)
Dim Packet As String

    Packet = "emotemsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub MapMsg(ByVal Text As String)
Dim Packet As String

    Packet = "mapmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Packet As String

    Packet = "playermsg" & SEP_CHAR & MsgTo & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub AdminMsg(ByVal Text As String)
Dim Packet As String

    Packet = "adminmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerMove()
Dim Packet As String

    Packet = "playermove" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerDir()
Dim Packet As String

    Packet = "playerdir" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerRequestNewMap()
Dim Packet As String
    
    Packet = "requestnewmap" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendMap()
Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long

    Packet = "MAPDATA" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim(CheckMap(GetPlayerMap(MyIndex)).name) & SEP_CHAR & CheckMap(GetPlayerMap(MyIndex)).Revision & SEP_CHAR & CheckMap(GetPlayerMap(MyIndex)).Moral & SEP_CHAR & CheckMap(GetPlayerMap(MyIndex)).Up & SEP_CHAR & CheckMap(GetPlayerMap(MyIndex)).Down & SEP_CHAR & CheckMap(GetPlayerMap(MyIndex)).Left & SEP_CHAR & CheckMap(GetPlayerMap(MyIndex)).Right & SEP_CHAR & CheckMap(GetPlayerMap(MyIndex)).Music & SEP_CHAR & CheckMap(GetPlayerMap(MyIndex)).BootMap & SEP_CHAR & CheckMap(GetPlayerMap(MyIndex)).BootX & SEP_CHAR & CheckMap(GetPlayerMap(MyIndex)).BootY & SEP_CHAR & CheckMap(GetPlayerMap(MyIndex)).Indoor & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With CheckMap(GetPlayerMap(MyIndex)).Tile(x, y)
                Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR
            End With
        Next x
    Next y
    
    For x = 1 To MAX_MAP_NPCS
        Packet = Packet & CheckMap(GetPlayerMap(MyIndex)).Npc(x) & SEP_CHAR
    Next x
    
    Packet = Packet & END_CHAR
    
    x = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, x)
    P2 = Mid(Packet, x + 1, Len(Packet) - x)
    Call SendData(Packet)
End Sub

Sub WarpMeTo(ByVal name As String)
Dim Packet As String

    Packet = "WARPMETO" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpToMe(ByVal name As String)
Dim Packet As String

    Packet = "WARPTOME" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpTo(ByVal MapNum As Long)
Dim Packet As String
    
    Packet = "WARPTO" & SEP_CHAR & MapNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetAccess(ByVal name As String, ByVal Access As Byte)
Dim Packet As String

    Packet = "SETACCESS" & SEP_CHAR & name & SEP_CHAR & Access & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetSprite(ByVal SpriteNum As Long)
Dim Packet As String

    Packet = "SETSPRITE" & SEP_CHAR & SpriteNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendKick(ByVal name As String)
Dim Packet As String

    Packet = "KICKPLAYER" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBan(ByVal name As String)
Dim Packet As String

    Packet = "BANPLAYER" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBanList()
Dim Packet As String

    Packet = "BANLIST" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditItem()
Dim Packet As String

    Packet = "REQUESTEDITITEM" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveItem(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "SAVEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).desc
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
                
Sub SendRequestEditEmoticon()
Dim Packet As String

    Packet = "REQUESTEDITEMOTICON" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveEmoticon(ByVal EmoNum As Long)
Dim Packet As String

    Packet = "SAVEEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Trim(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
                
Sub SendRequestEditNpc()
Dim Packet As String

    Packet = "REQUESTEDITNPC" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveNpc(ByVal NpcNum As Long)
Dim Packet As String
Dim i As Long
    
    Packet = "SAVENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).speed & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).EXP & SEP_CHAR
    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(NpcNum).ItemNPC(i).Chance
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendMapRespawn()
Dim Packet As String

    Packet = "MAPRESPAWN" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseItem(ByVal InvNum As Long)
Dim Packet As String

    Packet = "USEITEM" & SEP_CHAR & InvNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDropItem(ByVal InvNum, ByVal Ammount As Long)
Dim Packet As String

    Packet = "MAPDROPITEM" & SEP_CHAR & InvNum & SEP_CHAR & Ammount & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendWhosOnline()
Dim Packet As String

    Packet = "WHOSONLINE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendOnlineList()
Dim Packet As String

Packet = "ONLINELIST" & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub
            
Sub SendMOTDChange(ByVal MOTD As String)
Dim Packet As String

    Packet = "SETMOTD" & SEP_CHAR & MOTD & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditShop()
Dim Packet As String

    Packet = "REQUESTEDITSHOP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveShop(ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "SAVESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditSpell()
Dim Packet As String

    Packet = "REQUESTEDITSPELL" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveSpell(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "SAVESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditMap()
Dim Packet As String

    Packet = "REQUESTEDITMAP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendTradeRequest(ByVal name As String)
Dim Packet As String

    Packet = "PPTRADE" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAcceptTrade()
Dim Packet As String

    Packet = "ATRADE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDeclineTrade()
Dim Packet As String

    Packet = "DTRADE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPartyRequest(ByVal name As String)
Dim Packet As String

    Packet = "PARTY" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendJoinParty()
Dim Packet As String

    Packet = "JOINPARTY" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLeaveParty()
Dim Packet As String

    Packet = "LEAVEPARTY" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPartyInvite(ByVal name As String)
Dim Packet As String

    Packet = "INVITE" & SEP_CHAR & name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBanDestroy()
Dim Packet As String
    
    Packet = "BANDESTROY" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestLocation()
Dim Packet As String

    Packet = "REQUESTLOCATION" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendSetPlayerSprite(ByVal name As String, ByVal SpriteNum As Long)
Dim Packet As String

    Packet = "SETPLAYERSPRITE" & SEP_CHAR & name & SEP_CHAR & SpriteNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
