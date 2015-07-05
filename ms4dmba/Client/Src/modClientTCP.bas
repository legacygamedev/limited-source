Attribute VB_Name = "modClientTCP"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************

Private PlayerBuffer As clsBuffer

Sub TcpInit()

    Set PlayerBuffer = New clsBuffer

    ' used for parsing packets
    SEP_CHAR = vbNullChar ' ChrW$(0)
    END_CHAR = ChrW$(237)
    
    ' check if IP is valid
    If IsIP(GAME_IP) Then
        frmMirage.Socket.RemoteHost = GAME_IP
        frmMirage.Socket.RemotePort = GAME_PORT
    Else
        MsgBox GAME_IP & " does not appear as a valid IP address!"
        DestroyGame
    End If
        
End Sub

Sub DestroyTCP()
    frmMirage.Socket.Close
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long
Dim Data() As Byte

    frmMirage.Socket.GetData Buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes Buffer()
    
    'PlayerBuffer.DecompressBuffer
    
    If PlayerBuffer.Length >= 4 Then
        pLength = PlayerBuffer.ReadLong(False)
    End If
    
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4
        'make sure we have the right plength and pbuffer
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            
            Data() = PlayerBuffer.ReadBytes(pLength + 1)
            
            'If EncryptPackets Then
            '    Encryption_XOR_DecryptByte Data(), PacketKeys(PacketInIndex)
            '    PacketInIndex = PacketInIndex + 1
            '    If PacketInIndex > PacketEncKeys - 1 Then PacketInIndex = 0
            'End If
            
            HandleData Data()
        End If

        pLength = 0
        If PlayerBuffer.Length >= 4 Then
            pLength = PlayerBuffer.ReadLong(False)
        End If
    Loop

    ' Check if the playbuffer is empty
    If PlayerBuffer.Length <= 1 Then PlayerBuffer.Flush

End Sub

Public Function ConnectToServer(ByVal i As Long) As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    If i = 4 Then Exit Function
    
    Wait = GetTickCount
    
    With frmMirage.Socket
        .Close
        .Connect
    End With
    
    Call SetStatus("Connecting to server...(" & i & ")")
    
    ' Wait until connected or a few seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3500)
        DoEvents
        Sleep 20
    Loop
    
    ' return value
    If IsConnected Then
        ConnectToServer = True
    End If
    
    If Not ConnectToServer Then
        Call ConnectToServer(i + 1)
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

Function IsConnected() As Boolean
    If frmMirage.Socket.State = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If
End Function

Sub SendData(ByRef Data() As Byte)
Dim Buffer As clsBuffer

    If IsConnected Then
        Set Buffer = New clsBuffer
        
        'If EncryptPackets Then
        '    Encryption_XOR_EncryptByte Data, PacketKeys(PacketOutIndex)
        '    PacketOutIndex = PacketOutIndex + 1
        '    If PacketOutIndex > PacketEncKeys - 1 Then PacketOutIndex = 0
        'End If
                
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteBytes Data
        
        'Buffer.CompressBuffer
        
        frmMirage.Socket.SendData Buffer.ToArray()
        
        Set Buffer = Nothing
    End If
End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************

Public Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CNewAccount
    Buffer.WriteString Name
    Buffer.WriteString Password
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Public Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Name) + Len(Password) + 6
    Buffer.WriteLong CDelAccount
    Buffer.WriteString Name
    Buffer.WriteString Password
    
    SendData Buffer.ToArray()
End Sub

Public Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Name) + Len(Password) + 9
    Buffer.WriteLong CLogin
    Buffer.WriteString Name
    Buffer.WriteString Password
    Buffer.WriteLong App.Major
    Buffer.WriteLong App.Minor
    Buffer.WriteLong App.Revision
    
    SendData Buffer.ToArray()
End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Name) + 16
    Buffer.WriteLong CAddChar
    Buffer.WriteString Name
    Buffer.WriteLong Sex
    Buffer.WriteLong ClassNum
    Buffer.WriteLong Slot
    
    SendData Buffer.ToArray()
End Sub

Public Sub SendDelChar(ByVal Slot As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 6
    Buffer.WriteLong CDelChar
    Buffer.WriteLong Slot
    
    SendData Buffer.ToArray()
End Sub

Public Sub SendGetClasses()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CGetClasses
    
    SendData Buffer.ToArray()
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 6
    Buffer.WriteLong CUseChar
    Buffer.WriteLong CharSlot
    
    SendData Buffer.ToArray()
End Sub

Public Sub SayMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Text) + 4
    Buffer.WriteLong CSayMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Public Sub GlobalMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Text) + 4
    Buffer.WriteLong CGlobalMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Public Sub BroadcastMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Text) + 4
    Buffer.WriteLong CBroadcastMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Public Sub EmoteMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Text) + 4
    Buffer.WriteLong CEmoteMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Public Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Text) + Len(MsgTo) + 6
    Buffer.WriteLong CSayMsg
    Buffer.WriteString MsgTo
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Public Sub AdminMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Text) + 4
    Buffer.WriteLong CAdminMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Public Sub SendPlayerMove()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 10
    Buffer.WriteLong CPlayerMove
    Buffer.WriteLong GetPlayerDir(MyIndex)
    Buffer.WriteLong Player(MyIndex).Moving
    
    SendData Buffer.ToArray()
End Sub

Public Sub SendPlayerDir()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 6
    Buffer.WriteLong CPlayerDir
    Buffer.WriteLong GetPlayerDir(MyIndex)
    
    SendData Buffer.ToArray()
End Sub

Public Sub SendPlayerRequestNewMap()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 6
    Buffer.WriteLong CRequestNewMap
    Buffer.WriteLong GetPlayerDir(MyIndex)
    
    SendData Buffer.ToArray()
End Sub

Public Sub SendMap()
Dim Packet As String
Dim X As Long
Dim Y As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    CanMoveNow = False
    
    With Map
        Buffer.WriteLong CMapData
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteLong .Moral
        Buffer.WriteLong .TileSet
        Buffer.WriteLong .Up
        Buffer.WriteLong .Down
        Buffer.WriteLong .Left
        Buffer.WriteLong .Right
        Buffer.WriteLong .Music
        Buffer.WriteLong .BootMap
        Buffer.WriteLong .BootX
        Buffer.WriteLong .BootY
        Buffer.WriteLong .Shop
        Buffer.WriteLong .MaxX
        Buffer.WriteLong .MaxY
    End With
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                Buffer.WriteLong .Ground
                Buffer.WriteLong .Mask
                Buffer.WriteLong .Anim
                Buffer.WriteLong .Mask2
                Buffer.WriteLong .Fringe
                Buffer.WriteLong .Fringe2
                Buffer.WriteLong .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
            End With
        Next
    Next
    
    With Map
        For X = 1 To MAX_MAP_NPCS
            Buffer.WriteLong .Npc(X)
        Next
    End With
    
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
    
End Sub

Public Sub WarpMeTo(ByVal Name As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Name) + 4
    Buffer.WriteLong CWarpMeTo
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Public Sub WarpToMe(ByVal Name As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Name) + 4
    Buffer.WriteLong CWarpToMe
    Buffer.WriteString Name
    SendData Buffer.ToArray()
End Sub

Public Sub WarpTo(ByVal MapNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 6
    Buffer.WriteLong CWarpTo
    Buffer.WriteLong MapNum
    SendData Buffer.ToArray()
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 8
    Buffer.WriteLong CSetAccess
    Buffer.WriteString Name
    Buffer.WriteLong Access
    SendData Buffer.ToArray()
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 6
    Buffer.WriteLong CSetSprite
    Buffer.WriteLong SpriteNum
    SendData Buffer.ToArray()
End Sub

Public Sub SendKick(ByVal Name As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Name) + 4
    Buffer.WriteLong CKickPlayer
    Buffer.WriteString Name
    SendData Buffer.ToArray()
End Sub

Public Sub SendBan(ByVal Name As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Name) + 4
    Buffer.WriteLong CBanPlayer
    Buffer.WriteString Name
    SendData Buffer.ToArray()
End Sub

Public Sub SendBanList()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CBanList
    SendData Buffer.ToArray()
End Sub

Public Sub SendRequestEditItem()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CRequestEditItem
    SendData Buffer.ToArray()
End Sub

Public Sub SendSaveItem(ByVal ItemNum As Long)
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    Set Buffer = New clsBuffer
    
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    'buffer.preallocate ItemSize + 2
    Buffer.WriteLong CSaveItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    SendData Buffer.ToArray()
End Sub

Public Sub SendRequestEditNpc()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CRequestEditNpc
    SendData Buffer.ToArray()
End Sub

Public Sub SendSaveNpc(ByVal NpcNum As Long)
Dim Buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte

    Set Buffer = New clsBuffer
    
    NpcSize = LenB(Npc(NpcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(Npc(NpcNum)), NpcSize
    
    'buffer.preallocate NpcSize + 2
    Buffer.WriteLong CSaveNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteBytes NpcData
    SendData Buffer.ToArray()
End Sub

Public Sub SendMapRespawn()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CMapRespawn
    SendData Buffer.ToArray()
End Sub

Public Sub SendUseItem(ByVal InvNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 6
    Buffer.WriteLong CUseItem
    Buffer.WriteLong InvNum
    SendData Buffer.ToArray()
End Sub

Public Sub SendDropItem(ByVal InvNum As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CMapDropItem
    Buffer.WriteLong InvNum
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
End Sub

Public Sub SendWhosOnline()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CWhosOnline
    SendData Buffer.ToArray()
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(MOTD) + 4
    Buffer.WriteLong CSetMotd
    Buffer.WriteString MOTD
    SendData Buffer.ToArray()
End Sub

Public Sub SendRequestEditShop()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CRequestEditShop
    SendData Buffer.ToArray()
End Sub

Public Sub SendSaveShop(ByVal ShopNum As Long)
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    
    'buffer.preallocate ShopSize + 2
    Buffer.WriteLong CSaveShop
    Buffer.WriteLong ShopNum
    Buffer.WriteBytes ShopData
    SendData Buffer.ToArray()
End Sub

Public Sub SendRequestEditSpell()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CRequestEditSpell
    SendData Buffer.ToArray()
End Sub

Public Sub SendSaveSpell(ByVal SpellNum As Long)
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte

    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    'buffer.preallocate SpellSize + 2
    Buffer.WriteLong CSaveSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes SpellData
    SendData Buffer.ToArray()
End Sub

Public Sub SendRequestEditMap()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CRequestEditMap
    SendData Buffer.ToArray()
End Sub

Public Sub SendPartyRequest(ByVal Name As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate Len(Name) + 4
    Buffer.WriteLong CParty
    Buffer.WriteString Name
    SendData Buffer.ToArray()
End Sub

Public Sub SendJoinParty()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CJoinParty
    SendData Buffer.ToArray()
End Sub

Public Sub SendLeaveParty()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CLeaveParty
    SendData Buffer.ToArray()
End Sub

Public Sub SendBanDestroy()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    'buffer.preallocate 2
    Buffer.WriteLong CBanDestroy
    SendData Buffer.ToArray()
End Sub

