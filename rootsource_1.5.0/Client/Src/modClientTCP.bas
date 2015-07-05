Attribute VB_Name = "ClientTCP"
Option Explicit

' ******************************************
' **           Playerworlds Lite          **
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************

' TCP variables
Private PlayerBuffer As clsBuffer

Public Sub TcpInit()
    
    Set PlayerBuffer = New clsBuffer
    
    InitMessages
    
    ' check if IP is valid
    If IsIP(GAME_IP) Then
        frmMainGame.Socket.RemoteHost = GAME_IP
        frmMainGame.Socket.RemotePort = GAME_PORT
    Else
        MsgBox GAME_IP & " does not appear as a valid IP address!"
        DestroyGame
    End If
        
End Sub

Public Sub DestroyTCP()
    frmMainGame.Socket.Close
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Integer

    frmMainGame.Socket.GetData Buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes Buffer()
    
    If PlayerBuffer.Length >= 2 Then pLength = PlayerBuffer.ReadInteger(False)
    
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 2
        If pLength <= PlayerBuffer.Length - 2 Then
            PlayerBuffer.ReadInteger
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.Length >= 2 Then pLength = PlayerBuffer.ReadInteger(False)
    Loop
    PlayerBuffer.Trim
End Sub

Public Function ConnectToServer() As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If

    Wait = GetTickCount
    
    With frmMainGame.Socket
        .Close
        .Connect
    End With
    
    Call SetStatus("Connecting to server...")
    
    ' Wait until connected or a few seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3500)
        DoEvents
        Sleep 20
    Loop
    
    ' return value
    If IsConnected Then
        ConnectToServer = True
    End If
    
End Function

Private Function IsIP(ByVal IPAddress As String) As Boolean
Dim s() As String
Dim i As Long

    ' Check if connecting to localhost or URL
    If IPAddress = "localhost" Or InStr(1, IPAddress, "http://", vbTextCompare) = 1 Then
        IsIP = True
        Exit Function
    End If

    'If there are no periods, I have no idea what we have...
    If InStr(1, IPAddress, ".") = 0 Then Exit Function
    
    'Split up the string by the periods
    s = Split(IPAddress, ".")
    
    'Confirm we have ubound = 3, since xxx.xxx.xxx.xxx has 4 elements and we start at index 0
    If UBound(s) <> 3 Then Exit Function
    
    'Check that the values are numeric and in a valid range
    For i = 0 To 3
        If Val(s(i)) < 0 Then Exit Function
        If Val(s(i)) > 255 Then Exit Function
    Next
    
    'Looks like we were passed a valid IP!
    IsIP = True
    
End Function

Public Function IsConnected() As Boolean
    If frmMainGame.Socket.State = sckConnected Then
        IsConnected = True
    End If
End Function

Public Function IsPlaying(ByVal Index As Long) As Boolean
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If
End Function

Public Sub SendData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
    ' check if connection exist, otherwise will error
    If IsConnected Then
        Set Buffer = New clsBuffer
        Buffer.WriteInteger (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteBytes Data
        frmMainGame.Socket.SendData Buffer.ToArray()
    End If
End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************

Public Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + Len(Password) + 6
    Buffer.WriteInteger CNewAccount
    Buffer.WriteString Name
    Buffer.WriteString Password
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + Len(Password) + 6
    Buffer.WriteInteger CDelAccount
    Buffer.WriteString Name
    Buffer.WriteString Password
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + Len(Password) + 9
    Buffer.WriteInteger CLogin
    Buffer.WriteString Name
    Buffer.WriteString Password
    Buffer.WriteByte App.Major
    Buffer.WriteByte App.Minor
    Buffer.WriteByte App.Revision
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 16
    Buffer.WriteInteger CAddChar
    Buffer.WriteString Name
    Buffer.WriteLong Sex
    Buffer.WriteLong ClassNum + 1
    Buffer.WriteLong Slot
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendDelChar(ByVal Slot As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger CDelChar
    Buffer.WriteLong Slot
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendGetClasses()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CGetClasses
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger CUseChar
    Buffer.WriteLong CharSlot
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SayMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + 4
    Buffer.WriteInteger CSayMsg
    Buffer.WriteString Text
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub GlobalMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + 4
    Buffer.WriteInteger CGlobalMsg
    Buffer.WriteString Text
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub BroadcastMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + 4
    Buffer.WriteInteger CBroadcastMsg
    Buffer.WriteString Text
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub EmoteMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + 4
    Buffer.WriteInteger CEmoteMsg
    Buffer.WriteString Text
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + Len(MsgTo) + 6
    Buffer.WriteInteger CSayMsg
    Buffer.WriteString MsgTo
    Buffer.WriteString Text
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub AdminMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + 4
    Buffer.WriteInteger CAdminMsg
    Buffer.WriteString Text
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendPlayerMove()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 10
    Buffer.WriteInteger CPlayerMove
    Buffer.WriteLong GetPlayerDir(MyIndex)
    Buffer.WriteLong Player(MyIndex).Moving
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendPlayerDir()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger CPlayerDir
    Buffer.WriteLong GetPlayerDir(MyIndex)
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendPlayerRequestNewMap()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger CRequestNewMap
    Buffer.WriteLong GetPlayerDir(MyIndex)
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SyncPacket()
Dim Buffer As clsBuffer
    
    SentSync = True
    
    Set Buffer = New clsBuffer
    
    SyncX = Player(MyIndex).x
    SyncY = Player(MyIndex).y
    SyncMap = Player(MyIndex).map
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CSync
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendMap()
Dim Buffer As clsBuffer
Dim MapSize As Long
Dim MapData() As Byte

    Set Buffer = New clsBuffer
    
    MapSize = LenB(map(5))
    ReDim MapData(MapSize - 1)
    CopyMemory MapData(0), ByVal VarPtr(map(5)), MapSize
    
    Buffer.PreAllocate MapSize + 2
    Buffer.WriteInteger CMapData
    Buffer.WriteBytes MapData
    Call SendData(Buffer.ToArray())
End Sub

Public Sub WarpMeTo(ByVal Name As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 4
    Buffer.WriteInteger CWarpMeTo
    Buffer.WriteString Name
    
    Call SendData(Buffer.ToArray())
End Sub

Public Sub WarpToMe(ByVal Name As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 4
    Buffer.WriteInteger CWarpToMe
    Buffer.WriteString Name
    Call SendData(Buffer.ToArray())
End Sub

Public Sub WarpTo(ByVal MapNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger CWarpTo
    Buffer.WriteLong MapNum
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteInteger CSetAccess
    Buffer.WriteString Name
    Buffer.WriteLong Access
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger CSetSprite
    Buffer.WriteLong SpriteNum
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendKick(ByVal Name As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 4
    Buffer.WriteInteger CKickPlayer
    Buffer.WriteString Name
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendBan(ByVal Name As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 4
    Buffer.WriteInteger CBanPlayer
    Buffer.WriteString Name
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendBanList()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CBanList
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendRequestEditItem()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CRequestEditItem
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendSaveItem(ByVal ItemNum As Long)
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    Set Buffer = New clsBuffer
    
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    Buffer.PreAllocate ItemSize + 2
    Buffer.WriteInteger CSaveItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendRequestEditNpc()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CRequestEditNpc
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendSaveNpc(ByVal NpcNum As Long)
Dim Buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte

    Set Buffer = New clsBuffer
    
    NpcSize = LenB(Npc(NpcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(Npc(NpcNum)), NpcSize
    
    Buffer.PreAllocate NpcSize + 2
    Buffer.WriteInteger CSaveNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteBytes NpcData
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendMapRespawn()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CMapRespawn
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendUseItem(ByVal InvNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 6
    Buffer.WriteInteger CUseItem
    Buffer.WriteLong InvNum
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendDropItem(ByVal InvNum As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CMapDropItem
    Buffer.WriteLong InvNum
    Buffer.WriteLong Amount
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendWhosOnline()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CWhosOnline
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(MOTD) + 4
    Buffer.WriteInteger CSetMotd
    Buffer.WriteString MOTD
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendRequestEditShop()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CRequestEditShop
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendSaveShop(ByVal ShopNum As Long)
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    
    Buffer.PreAllocate ShopSize + 2
    Buffer.WriteInteger CSaveShop
    Buffer.WriteLong ShopNum
    Buffer.WriteBytes ShopData
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendRequestEditSpell()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CRequestEditSpell
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte

    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    Buffer.PreAllocate SpellSize + 2
    Buffer.WriteInteger CSaveSpell
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendRequestEditMap()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CRequestEditMap
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendPartyRequest(ByVal Name As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 4
    Buffer.WriteInteger CParty
    Buffer.WriteString Name
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendJoinParty()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CJoinParty
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendLeaveParty()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CLeaveParty
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendBanDestroy()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CBanDestroy
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendCreateGuild(ByVal user As String, ByVal Guild As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Guild) + Len(user) + 6
    Buffer.WriteInteger CCreateGuild
    Buffer.WriteString user
    Buffer.WriteString Guild
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendRemoveFromGuild(ByVal user As String, ByVal Guild As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Guild) + Len(user) + 6
    Buffer.WriteInteger CRemoveFromGuild
    Buffer.WriteString user
    Buffer.WriteString Guild
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendGuildInvite(ByVal user As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(user) + 4
    Buffer.WriteInteger CInviteGuild
    Buffer.WriteString user
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendGuildKick(ByVal user As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(user) + 4
    Buffer.WriteInteger CKickGuild
    Buffer.WriteString user
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendGuildPromote(ByVal user As String, ByVal Access As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(user) + 4 + 4
    Buffer.WriteInteger CGuildPromote
    Buffer.WriteString user
    Buffer.WriteLong Access
    Call SendData(Buffer.ToArray())
End Sub

Public Sub SendLeaveGuild()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CLeaveGuild
    Call SendData(Buffer.ToArray())
End Sub
