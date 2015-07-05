Attribute VB_Name = "modClientTCP"
Option Explicit
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/21/2005  Verrigan   Modified all procedures that send
'*                        packets so they send them as byte
'*                        arrays.
'****************************************************************
Sub TcpInit()
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Replaced hard-coded IP with constant.
'****************************************************************

    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    PlayerBuffer = ""
        
    frmMirage.Socket.RemoteHost = GAME_IP
    frmMirage.Socket.RemotePort = GAME_PORT
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

Public Function ConnectToServer() As Boolean
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************

Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    With frmMirage.Socket
        .Close
        .Connect
    End With
    
    ' Wait until connected or 15 seconds have passed and report the server being down
    ' Changed to 15 seconds to allow users with firewalls enough time to allow the
    ' application access to the internet. -- Verrigan
    Do While (Not IsConnected) And (GetTickCount <= Wait + 15000)
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
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 09/13/2005  Verrigan   SendDataNew() function created to send
'*                        the new byte array packets. It prepends
'*                        the data with the packet type, then it
'*                        prepends the data with the packet size,
'*                        and finally sends the packet. :)
'****************************************************************
Sub SendDataNew(ByRef Data() As Byte, pType As SMsgTypes)
  If IsConnected Then
    Data = PrefixBuffer(Data, VarPtr(CByte(pType)), 1)
    Data = PrefixBuffer(Data, VarPtr(aLen(Data)), 2)
    frmMirage.Socket.SendData Data
    DoEvents
  End If
End Sub
'Sub SendData(ByVal Data As String)
'    If IsConnected Then
'        frmMirage.Socket.SendData Data
'        Debug.Print Data
'        DoEvents
'    End If
'End Sub
Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Name)
  Buffer = AddStringToBuffer(Buffer, Password)
  Buffer = AddIntegerToBuffer(Buffer, App.Major)
  Buffer = AddIntegerToBuffer(Buffer, App.Minor)
  Buffer = AddIntegerToBuffer(Buffer, App.Revision)
  
  Call SendDataNew(Buffer, SMsgNewAccount)
End Sub
Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Name)
  Buffer = AddStringToBuffer(Buffer, Password)
  Buffer = AddIntegerToBuffer(Buffer, App.Major)
  Buffer = AddIntegerToBuffer(Buffer, App.Minor)
  Buffer = AddIntegerToBuffer(Buffer, App.Revision)
  
  'Packet = "delaccount" & SEP_CHAR & Trim(Name) & SEP_CHAR & Trim(Password) & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgDelAccount)
End Sub
Sub SendLogin(ByVal Name As String, ByVal Password As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Name)
  Buffer = AddStringToBuffer(Buffer, Password)
  Buffer = AddIntegerToBuffer(Buffer, App.Major)
  Buffer = AddIntegerToBuffer(Buffer, App.Minor)
  Buffer = AddIntegerToBuffer(Buffer, App.Revision)
  
  'Packet = "login" & SEP_CHAR & Trim(Name) & SEP_CHAR & Trim(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgLogin)
End Sub
Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Name)
  Buffer = AddByteToBuffer(Buffer, CByte(Sex))
  Buffer = AddByteToBuffer(Buffer, CByte(ClassNum))
  Buffer = AddByteToBuffer(Buffer, CByte(Slot))
  
  'Packet = "addchar" & SEP_CHAR & Trim(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgAddChar)
End Sub
Sub SendDelChar(ByVal Slot As Long)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddByteToBuffer(Buffer, CByte(Slot))
    
  'Packet = "delchar" & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgDelChar)
End Sub
Sub SendGetClasses()
  Dim Buffer() As Byte
  
  Buffer = ""

  'Packet = "getclasses" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgGetClasses)
End Sub
Sub SendUseChar(ByVal CharSlot As Long)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddByteToBuffer(Buffer, CByte(CharSlot))
  
  'Packet = "usechar" & SEP_CHAR & CharSlot & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgUseChar)
End Sub
Sub SayMsg(ByVal Text As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Text)
  
  'Packet = "saymsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgSay)
End Sub
Sub GlobalMsg(ByVal Text As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Text)
  
  'Packet = "globalmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgGlobal)
End Sub
Sub BroadcastMsg(ByVal Text As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Text)

  'Packet = "broadcastmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgBroadcast)
End Sub
Sub EmoteMsg(ByVal Text As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Text)
  
  'Packet = "emotemsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgEmote)
End Sub
'This procedure is never called and is not handled by the server. -- Verrigan
'Sub MapMsg(ByVal Text As String)
'Dim Packet As String
'
'    Packet = "mapmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
'    Call SendData(Packet)
'End Sub
Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, MsgTo)
  Buffer = AddStringToBuffer(Buffer, Text)
  
  'Packet = "playermsg" & SEP_CHAR & MsgTo & SEP_CHAR & Text & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgPlayer)
End Sub
Sub AdminMsg(ByVal Text As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Text)
  
  'Packet = "adminmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgAdmin)
End Sub
Sub SendPlayerMove()
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddByteToBuffer(Buffer, CByte(GetPlayerDir(MyIndex)))
  Buffer = AddByteToBuffer(Buffer, Player(MyIndex).Moving)
  
  'Packet = "playermove" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgPlayerMove)
End Sub
Sub SendPlayerDir()
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddByteToBuffer(Buffer, CByte(GetPlayerDir(MyIndex)))
  
  'Packet = "playerdir" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgPlayerDir)
End Sub
Sub SendPlayerRequestNewMap()
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddByteToBuffer(Buffer, CByte(GetPlayerDir(MyIndex)))
  
  'Packet = "requestnewmap" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgRequestNewMap)
End Sub
'Rewrote entire SendMap() sub. -- Verrigan
'Public Sub SendMap()
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************
'
'Dim Packet As String, P1 As String, P2 As String
'Dim x As Long
'Dim y As Long
'
'
'    With Map
'        Packet = "MAPDATA" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim(.Name) & SEP_CHAR & .Revision & SEP_CHAR & .Moral & SEP_CHAR & .Up & SEP_CHAR & .Down & SEP_CHAR & .Left & SEP_CHAR & .Right & SEP_CHAR & .Music & SEP_CHAR & .BootMap & SEP_CHAR & .BootX & SEP_CHAR & .BootY & SEP_CHAR & .Shop & SEP_CHAR
'    End With
'
'    For y = 0 To MAX_MAPY
'        For x = 0 To MAX_MAPX
'            With Map.Tile(x, y)
'                Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Fringe & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR
'            End With
'        Next x
'    Next y
'
'    With Map
'        For x = 1 To MAX_MAP_NPCS
'            Packet = Packet & .Npc(x) & SEP_CHAR
'        Next x
'    End With
'
'    Packet = Packet & END_CHAR
'
'    x = Int(Len(Packet) / 2)
'    P1 = Mid(Packet, 1, x)
'    P2 = Mid(Packet, x + 1, Len(Packet) - x)
'    Call SendData(Packet)
'End Sub
Public Sub SendMap()
  Dim Buffer() As Byte
  
  Buffer = FillBuffer(VarPtr(Map), LenB(Map))
  
  Call SendDataNew(Buffer, SMsgMapData)
End Sub
Sub WarpMeTo(ByVal Name As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Name)

  'Packet = "WARPMETO" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgWarpMeTo)
End Sub
Sub WarpToMe(ByVal Name As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Name)
  
  'Packet = "WARPTOME" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgWarpToMe)
End Sub
Sub WarpTo(ByVal MapNum As Long)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddIntegerToBuffer(Buffer, CInt(MapNum))
  
  'Packet = "WARPTO" & SEP_CHAR & MapNum & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgWarpTo)
End Sub
Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Name)
  Buffer = AddByteToBuffer(Buffer, Access)
  
  'Packet = "SETACCESS" & SEP_CHAR & Name & SEP_CHAR & Access & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgSetAccess)
End Sub
Sub SendSetSprite(ByVal SpriteNum As Long)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddIntegerToBuffer(Buffer, CInt(SpriteNum))
  
  'Packet = "SETSPRITE" & SEP_CHAR & SpriteNum & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgSetSprite)
End Sub
Sub SendKick(ByVal Name As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Name)
  
  'Packet = "KICKPLAYER" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgKickPlayer)
End Sub
Sub SendBan(ByVal Name As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Name)
  
  'Packet = "BANPLAYER" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgBanPlayer)
End Sub
Sub SendBanList()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "BANLIST" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgBanList)
End Sub
Sub SendRequestEditItem()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "REQUESTEDITITEM" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgRequestEditItem)
End Sub
'Rewrote entire SendSaveItem() sub. -- Verrigan
'Public Sub SendSaveItem(ByVal ItemNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************
'
'Dim Packet As String
'
'    With Item(ItemNum)
'        Packet = "SAVEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(.Name) & SEP_CHAR & .Pic & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & END_CHAR
'    End With
'
'    Call SendData(Packet)
'End Sub
Public Sub SendSaveItem(ByVal ItemNum As Long)
  Dim Buffer() As Byte
  Dim tBytes() As Byte
  
  tBytes = FillBuffer(VarPtr(Item(ItemNum)), LenB(Item(ItemNum)))
  
  Buffer = GetFromBuffer(tBytes, NAME_LENGTH * 2, True)
  Buffer = StrConv(Buffer, vbFromUnicode)
  
  Buffer = AddToBuffer(Buffer, tBytes, aLen(tBytes))
  
  Call SendDataNew(Buffer, SMsgSaveItem)
End Sub
Sub SendRequestEditNpc()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "REQUESTEDITNPC" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgRequestEditNPC)
End Sub
'Rewrote entire SendSaveNpc sub. -- Verrigan
'Public Sub SendSaveNpc(ByVal NpcNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************
'
'Dim Packet As String
'
'    With Npc(NpcNum)
'        Packet = "SAVENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(.Name) & SEP_CHAR & Trim(.AttackSay) & SEP_CHAR & .Sprite & SEP_CHAR & .SpawnSecs & SEP_CHAR & .Behavior & SEP_CHAR & .Range & SEP_CHAR & .DropChance & SEP_CHAR & .DropItem & SEP_CHAR & .DropItemValue & SEP_CHAR & .STR & SEP_CHAR & .DEF & SEP_CHAR & .SPEED & SEP_CHAR & .MAGI & SEP_CHAR & END_CHAR
'    End With
'
'    Call SendData(Packet)
'End Sub
Public Sub SendSaveNpc(ByVal NpcNum As Long)
  Dim Buffer() As Byte
  Dim tBytes() As Byte
  
  tBytes = FillBuffer(VarPtr(Npc(NpcNum)), LenB(Npc(NpcNum)))
  
  Buffer = GetFromBuffer(tBytes, NAME_LENGTH * 2, True)
  Buffer = StrConv(Buffer, vbFromUnicode)
  
  Buffer = AddToBuffer(Buffer, tBytes, aLen(tBytes))
  Buffer = PrefixBuffer(Buffer, CInt(NpcNum), 2)
  
  Call SendDataNew(Buffer, SMsgSaveNPC)
End Sub
Sub SendMapRespawn()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "MAPRESPAWN" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgMapRespawn)
End Sub
Sub SendUseItem(ByVal InvNum As Long)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddByteToBuffer(Buffer, CByte(InvNum))
  
  'Packet = "USEITEM" & SEP_CHAR & InvNum & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgUseItem)
End Sub
Sub SendDropItem(ByVal InvNum As Long, ByVal Ammount As Long)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddByteToBuffer(Buffer, CByte(InvNum))
  Buffer = AddLongToBuffer(Buffer, Ammount)
  
  'Packet = "MAPDROPITEM" & SEP_CHAR & InvNum & SEP_CHAR & Ammount & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgMapDropItem)
End Sub
Sub SendWhosOnline()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "WHOSONLINE" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgWhosOnline)
End Sub
Sub SendMOTDChange(ByVal MOTD As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, MOTD)
  
  'Packet = "SETMOTD" & SEP_CHAR & MOTD & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgSetMOTD)
End Sub
Sub SendRequestEditShop()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "REQUESTEDITSHOP" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgRequestEditShop)
End Sub
'Rewrote entire SendSaveShop() sub. -- Verrigan
'Public Sub SendSaveShop(ByVal ShopNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************
'
'Dim Packet As String
'Dim i As Long
'
'    With Shop(ShopNum)
'        Packet = "SAVESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(.Name) & SEP_CHAR & Trim(.JoinSay) & SEP_CHAR & Trim(.LeaveSay) & SEP_CHAR & .FixesItems & SEP_CHAR
'    End With
'
'    For i = 1 To MAX_TRADES
'        With Shop(ShopNum).TradeItem(i)
'            Packet = Packet & .GiveItem & SEP_CHAR & .GiveValue & SEP_CHAR & .GetItem & SEP_CHAR & .GetValue & SEP_CHAR
'        End With
'    Next i
'
'    Packet = Packet & END_CHAR
'    Call SendData(Packet)
'End Sub
Public Sub SendSaveShop(ByVal ShopNum As Long)
  Dim Buffer() As Byte
  Dim tBytes() As Byte
  
  tBytes = FillBuffer(VarPtr(Shop(ShopNum)), LenB(Shop(ShopNum)))
  
  Buffer = GetFromBuffer(tBytes, NAME_LENGTH * 2, True)
  Buffer = StrConv(Buffer, vbFromUnicode)
  
  Buffer = AddToBuffer(Buffer, tBytes, aLen(tBytes))
  
  Call SendDataNew(Buffer, SMsgSaveShop)
End Sub
Sub SendRequestEditSpell()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "REQUESTEDITSPELL" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgRequestEditSpell)
End Sub
'Rewrote entire SendSaveSpell() sub. -- Verrigan
'Public Sub SendSaveSpell(ByVal SpellNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************
'
'Dim Packet As String
'
'    With Spell(SpellNum)
'        Packet = "SAVESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(.Name) & SEP_CHAR & .ClassReq & SEP_CHAR & .LevelReq & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & END_CHAR
'    End With
'
'    Call SendData(Packet)
'End Sub
Public Sub SendSaveSpell(ByVal SpellNum As Long)
  Dim Buffer() As Byte
  Dim tBytes() As Byte
  
  tBytes = FillBuffer(VarPtr(Spell(SpellNum)), LenB(Spell(SpellNum)))
  
  Buffer = GetFromBuffer(tBytes, NAME_LENGTH * 2, True)
  Buffer = StrConv(Buffer, vbFromUnicode)
  
  Buffer = AddToBuffer(Buffer, tBytes, aLen(tBytes))
  
  Call SendDataNew(Buffer, SMsgSaveSpell)
End Sub
Sub SendRequestEditMap()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "REQUESTEDITMAP" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgRequestEditMap)
End Sub

Sub SendPartyRequest(ByVal Name As String)
  Dim Buffer() As Byte
  
  Buffer = ""
  Buffer = AddStringToBuffer(Buffer, Name)
  
  'Packet = "PARTY" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgParty)
End Sub
Sub SendJoinParty()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "JOINPARTY" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgJoinParty)
End Sub
Sub SendLeaveParty()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "LEAVEPARTY" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgLeaveParty)
End Sub
Sub SendBanDestroy()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "BANDESTROY" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgBanDestroy)
End Sub
Sub SendRequestLocation()
  Dim Buffer() As Byte
  
  Buffer = ""
  
  'Packet = "REQUESTLOCATION" & SEP_CHAR & END_CHAR
  Call SendDataNew(Buffer, SMsgRequestLocation)
End Sub
