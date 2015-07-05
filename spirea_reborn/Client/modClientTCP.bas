Attribute VB_Name = "modClientTCP"
Option Explicit
Sub TcpInit()
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Replaced hard-coded IP with constant.
'****************************************************************

    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    PlayerBuffer = ""
    XToGo = -1
    YToGo = -1
    toRun = False
     Dim filename As String
    filename = App.Path & "\config.ini"

    frmMirage.Socket.RemoteHost = GetVar(filename, "IPCONFIG", "IP")
    frmMirage.Socket.RemotePort = Val(GetVar(filename, "IPCONFIG", "PORT"))
End Sub

Sub TcpDestroy()
    frmMirage.Socket.Close
    
    If frmChars.Visible Then frmChars.Visible = False
    If frmCredits.Visible Then frmCredits.Visible = False
    If frmDeleteAccount.Visible Then frmDeleteAccount.Visible = False
    If frmLogin.Visible Then frmLogin.Visible = False
    If frmNewAccount.Visible Then frmNewAccount.Visible = False
    If frmNewChar.Visible Then frmNewChar.Visible = False
End Sub

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim top As String * 3
Dim Start As Integer

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

Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim Packet As String

    Packet = "newaccount" & SEP_CHAR & Trim(Name) & SEP_CHAR & Trim(Password) & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
Dim Packet As String
    
    Packet = "delaccount" & SEP_CHAR & Trim(Name) & SEP_CHAR & Trim(Password) & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim Packet As String

    Packet = "login" & SEP_CHAR & Trim(Name) & SEP_CHAR & Trim(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
Dim Packet As String

    Packet = "addchar" & SEP_CHAR & Trim(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & Slot & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelChar(ByVal Slot As Long)
Dim Packet As String
    
    Packet = "delchar" & SEP_CHAR & Slot & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetClasses()
Dim Packet As String

    Packet = "getclasses" & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseChar(ByVal CharSlot As Long)
Dim Packet As String

    Packet = "usechar" & SEP_CHAR & CharSlot & END_CHAR
    Call SendData(Packet)
End Sub

Sub SayMsg(ByVal Text As String)
Dim Packet As String

    Packet = "saymsg" & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub GlobalMsg(ByVal Text As String)
Dim Packet As String

    Packet = "globalmsg" & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub BroadcastMsg(ByVal Text As String)
Dim Packet As String

    Packet = "broadcastmsg" & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub EmoteMsg(ByVal Text As String)
Dim Packet As String

    Packet = "emotemsg" & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub MapMsg(ByVal Text As String)
Dim Packet As String

    Packet = "mapmsg" & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Packet As String

    Packet = "playermsg" & SEP_CHAR & MsgTo & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub AdminMsg(ByVal Text As String)
Dim Packet As String

    Packet = "adminmsg" & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerMove()
Dim Packet As String

    Packet = "playermove" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerDir()
Dim Packet As String

    Packet = "playerdir" & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerRequestNewMap()
Dim Packet As String
    
    Packet = "requestnewmap" & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendMap()
Dim Packet As String, P1 As String, P2 As String
Dim X As Long
Dim Y As Long

    Packet = "MAPDATA" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim(Map.Name) & SEP_CHAR & Map.Revision & SEP_CHAR & Map.Moral & SEP_CHAR & Map.Up & SEP_CHAR & Map.Down & SEP_CHAR & Map.Left & SEP_CHAR & Map.Right & SEP_CHAR & Map.Music & SEP_CHAR & Map.BootMap & SEP_CHAR & Map.BootX & SEP_CHAR & Map.BootY & SEP_CHAR & Map.Shop & SEP_CHAR
    
  For Y = 0 To MAX_MAPY
For X = 0 To MAX_MAPX
With Map.Tile(X, Y)
Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR
End With
Next X
Next Y
    
    For X = 1 To MAX_MAP_NPCS
        Packet = Packet & Map.Npc(X) & SEP_CHAR
    Next X
    
    Packet = Packet & END_CHAR
    
    X = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, X)
    P2 = Mid(Packet, X + 1, Len(Packet) - X)
    Call SendData(Packet)
End Sub

Sub WarpMeTo(ByVal Name As String)
Dim Packet As String

    Packet = "WARPMETO" & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpToMe(ByVal Name As String)
Dim Packet As String

    Packet = "WARPTOME" & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpTo(ByVal MapNum As Long)
Dim Packet As String
    
    Packet = "WARPTO" & SEP_CHAR & MapNum & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim Packet As String

    Packet = "SETACCESS" & SEP_CHAR & Name & SEP_CHAR & Access & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetSprite(ByVal SpriteNum As Long)
Dim Packet As String

    Packet = "SETSPRITE" & SEP_CHAR & SpriteNum & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendKick(ByVal Name As String)
Dim Packet As String

    Packet = "KICKPLAYER" & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBan(ByVal Name As String)
Dim Packet As String

    Packet = "BANPLAYER" & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBanList()
Dim Packet As String

    Packet = "BANLIST" & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendRequestEditItem()
Dim Packet As String

    Packet = "REQUESTEDITITEM" & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveItem(ByVal itemnum As Long)
Dim Packet As String

    Packet = "SAVEITEM" & SEP_CHAR & itemnum & SEP_CHAR & Trim(Item(itemnum).Name) & SEP_CHAR & Item(itemnum).Pic & SEP_CHAR & Item(itemnum).Type & SEP_CHAR & Item(itemnum).Data1 & SEP_CHAR & Item(itemnum).Data2 & SEP_CHAR & Item(itemnum).Data3 & END_CHAR
    Call SendData(Packet)
End Sub
                
Sub SendRequestEditNpc()
Dim Packet As String

    Packet = "REQUESTEDITNPC" & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveNpc(ByVal NpcNum As Long)
Dim Packet As String
    
    Packet = "SAVENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendMapRespawn()
Dim Packet As String

    Packet = "MAPRESPAWN" & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseItem(ByVal InvNum As Long)
Dim Packet As String

    Packet = "USEITEM" & SEP_CHAR & InvNum & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDropItem(ByVal InvNum, ByVal Ammount As Long)
Dim Packet As String

    Packet = "MAPDROPITEM" & SEP_CHAR & InvNum & SEP_CHAR & Ammount & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendWhosOnline()
Dim Packet As String

    Packet = "WHOSONLINE" & END_CHAR
    Call SendData(Packet)
End Sub
            
Sub SendMOTDChange(ByVal MOTD As String)
Dim Packet As String

    Packet = "SETMOTD" & SEP_CHAR & MOTD & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditShop()
Dim Packet As String

    Packet = "REQUESTEDITSHOP" & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveShop(ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "SAVESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditSpell()
Dim Packet As String

    Packet = "REQUESTEDITSPELL" & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveSpell(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "SAVESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditMap()
Dim Packet As String

    Packet = "REQUESTEDITMAP" & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPartyRequest(ByVal Name As String)
Dim Packet As String

    Packet = "PARTY" & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendJoinParty()
Dim Packet As String

    Packet = "JOINPARTY" & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLeaveParty()
Dim Packet As String

    Packet = "LEAVEPARTY" & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBanDestroy()
Dim Packet As String
    
    Packet = "BANDESTROY" & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestLocation()
Dim Packet As String

    Packet = "REQUESTLOCATION" & END_CHAR
    Call SendData(Packet)
End Sub
