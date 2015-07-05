Attribute VB_Name = "modClientTCP"
Option Explicit

Sub InitTcp()

    frmSendGetData.Visible = True
    SetStatus "Initializing TCP settings..."
    
    InitMessages
    EncryptPackets = 1
    If EncryptPackets Then GenerateEncryptionKeys PacketKeys
    
    PacketInIndex = 0
    PacketOutIndex = 0
    
    Set PlayerBuffer = New clsBuffer
    
    frmMainGame.Socket.RemoteHost = GAME_IP
    frmMainGame.Socket.RemotePort = GAME_PORT
    
End Sub

Sub TcpDestroy()
    frmMainGame.Socket.Close
    
    If frmChars.Visible Then frmChars.Visible = False
    If frmLogin.Visible Then frmLogin.Visible = False
    If frmNewAccount.Visible Then frmNewAccount.Visible = False
    If frmNewChar.Visible Then frmNewChar.Visible = False
    If frmEvent.Visible Then frmEvent.Visible = False
    If frmDeleteCharacter.Visible Then frmDeleteCharacter.Visible = False
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    frmMainGame.Socket.GetData Buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes Buffer()
    
    If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
End Sub

Function ConnectToServer() As Boolean
Dim Wait As Long
    
    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMainGame.Socket.Close
    frmMainGame.Socket.Connect
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop
    
    ConnectToServer = IsConnected
End Function

Function IsConnected() As Boolean
'    If frmMainGame.Socket.State = sckConnected Then
'        IsConnected = True
'    End If
    IsConnected = frmMainGame.Socket.State = sckConnected
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If LenB(Current_Name(Index)) > 0 Then IsPlaying = True
End Function

Sub SendData(ByRef Data() As Byte)
Dim Buffer As clsBuffer

    If IsConnected Then
        Set Buffer = New clsBuffer
        
        If EncryptPackets Then
            Encryption_XOR_EncryptByte Data, PacketKeys(PacketOutIndex)
            PacketOutIndex = PacketOutIndex + 1
            If PacketOutIndex > PacketEncKeys - 1 Then PacketOutIndex = 0
        End If
                
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1  ' Write the length
        Buffer.WriteBytes Data()                            ' Write the data to the packet
        frmMainGame.Socket.SendData Buffer.ToArray()        ' Send the data
    End If
End Sub

Sub SendGetClasses()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgGetClasses
    
    SendData Buffer.ToArray()
End Sub

Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Name)) + Len(Trim$(Password)) + 12
    Buffer.WriteLong SMsgNewAccount
    Buffer.WriteString Trim$(Name)
    Buffer.WriteString Trim$(Password)
    
    SendData Buffer.ToArray()
End Sub

Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Name)) + Len(Trim$(Password)) + Len("3.0.7") + 16
    Buffer.WriteLong SMsgLogin
    Buffer.WriteString Trim$(Name)
    Buffer.WriteString Trim$(Password)
    Buffer.WriteString "3.0.7"
    
    SendData Buffer.ToArray()
End Sub

Sub SendRequestEditEmoticon()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgRequestEditEmoticon
    
    SendData Buffer.ToArray()
End Sub

Sub SendEditEmoticon()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgEditEmoticon
    Buffer.WriteLong EditorIndex
    
    SendData Buffer.ToArray()
End Sub

Sub SendSaveEmoticon(ByVal EmoticonNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Emoticons(EmoticonNum).Command)) + 16
    Buffer.WriteLong SMsgSaveEmoticon
    Buffer.WriteLong EmoticonNum
    Buffer.WriteString Trim$(Emoticons(EmoticonNum).Command)
    Buffer.WriteLong Emoticons(EmoticonNum).Pic
    
    SendData Buffer.ToArray()
End Sub

Sub SendCheckEmoticon(ByVal EmoticonNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgCheckEmoticon
    Buffer.WriteLong EmoticonNum
    
    SendData Buffer.ToArray()
End Sub

Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long, ByVal SpriteNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Trim$(Name)) + 24
    Buffer.WriteLong SMsgAddChar
    Buffer.WriteString Trim$(Name)
    Buffer.WriteLong Sex
    Buffer.WriteLong ClassNum
    Buffer.WriteLong Slot
    Buffer.WriteLong SpriteNum
    
    SendData Buffer.ToArray()
End Sub

Sub SendDelChar(ByVal Slot As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgDelChar
    Buffer.WriteByte Slot
    
    SendData Buffer.ToArray()
End Sub

Sub SendUseChar(ByVal Slot As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgUseChar
    Buffer.WriteByte Slot
    
    SendData Buffer.ToArray()
End Sub

Sub SayMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + 8
    Buffer.WriteLong SMsgSayMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Sub EmoteMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + 8
    Buffer.WriteLong SMsgEmoteMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Sub RealmMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + 8
    Buffer.WriteLong SMsgGlobalMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Sub AdminMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + 8
    Buffer.WriteLong SMsgAdminMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Sub PartyMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + 8
    Buffer.WriteLong SMsgPartyMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + Len(MsgTo) + 12
    Buffer.WriteLong SMsgPlayerMsg
    Buffer.WriteString MsgTo
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Sub SendPlayerMove()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong SMsgPlayerMove
    Buffer.WriteLong Current_Dir(MyIndex)
    Buffer.WriteLong Player(MyIndex).Moving
    
    SendData Buffer.ToArray()
End Sub

Sub SendPlayerDir()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgPlayerDir
    Buffer.WriteLong Current_Dir(MyIndex)
    
    SendData Buffer.ToArray()
End Sub

Sub SendUseItem(ByVal InvNum As Long)
Dim Buffer As clsBuffer

    ' Doesnt' matter if they are dead
    If Player(MyIndex).IsDead Then Exit Sub
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgUseItem
    Buffer.WriteLong InvNum
    
    SendData Buffer.ToArray()
End Sub

Sub SendUnequipSlot(ByVal EquipmentSlot As Long)
Dim Buffer As clsBuffer

    ' Doesnt' matter if they are dead
    If Player(MyIndex).IsDead Then Exit Sub
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgUnequipSlot
    Buffer.WriteLong EquipmentSlot
    
    SendData Buffer.ToArray()
End Sub

Sub SendAttack()
Dim Buffer As clsBuffer

    ' Doesnt' matter if they are dead
    If Player(MyIndex).IsDead Then Exit Sub
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgAttack
    
    SendData Buffer.ToArray()
End Sub

Sub SendUseStatPoint(ByVal Stat As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgUseStatPoint
    Buffer.WriteLong Stat
    
    SendData Buffer.ToArray()
End Sub

Sub SendPlayerInfoRequest(ByVal Name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong SMsgPlayerInfoRequest
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Sub WarpMeTo(ByVal Name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong SMsgWarpMeTo
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Sub WarpToMe(ByVal Name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong SMsgWarpToMe
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Sub WarpTo(ByVal MapNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgWarpTo
    Buffer.WriteLong MapNum
    
    SendData Buffer.ToArray()
End Sub

Sub SendSetSprite(ByVal Name As String, ByVal SpriteNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 12
    Buffer.WriteLong SMsgSetSprite
    Buffer.WriteString Name
    Buffer.WriteLong SpriteNum
    
    SendData Buffer.ToArray()
End Sub

Sub SendGetStats()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgGetStats
    
    SendData Buffer.ToArray()
End Sub

Sub SendClickWarp(ByVal X As Long, ByVal Y As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong SMsgClickWarp
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    SendData Buffer.ToArray()
End Sub

Sub SendPlayerRequestNewMap()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgRequestNewMap
    Buffer.WriteLong Current_Dir(MyIndex)
    
    SendData Buffer.ToArray()
End Sub

Sub SendMap()
Dim Buffer As clsBuffer
Dim MapData As clsBuffer
Dim X As Long
Dim Y As Long
Dim TileSize As Long
Dim TileData() As Byte
    
    Set Buffer = New clsBuffer
    Set MapData = New clsBuffer
    
    MapData.PreAllocate LenB(Map)
    With Map
        MapData.WriteString Trim$(.Name)
        MapData.WriteLong .Revision
        MapData.WriteByte .Moral
        MapData.WriteInteger .Up
        MapData.WriteInteger .Down
        MapData.WriteInteger .Left
        MapData.WriteInteger .Right
        MapData.WriteByte .Music
        MapData.WriteInteger .BootMap
        MapData.WriteByte .BootX
        MapData.WriteByte .BootY
        MapData.WriteByte .TileSet
        MapData.WriteByte .MaxX
        MapData.WriteByte .MaxY
    
        For X = 1 To MAX_MOBS
            MapData.WriteLong .Mobs(X).NpcCount
            If .Mobs(X).NpcCount > 0 Then
                For Y = 1 To .Mobs(X).NpcCount
                    MapData.WriteLong .Mobs(X).Npc(Y)
                Next
            End If
        Next
        
        TileSize = LenB(.Tile(0, 0)) * ((UBound(.Tile, 1) + 1) * (UBound(.Tile, 2) + 1))
        ReDim TileData(0 To TileSize - 1)
        CopyMemory TileData(0), ByVal VarPtr(.Tile(0, 0)), TileSize
        MapData.WriteBytes TileData
    End With
    MapData.CompressBuffer
    
    ' Now write the mapdata to the real buffer
    Buffer.PreAllocate Buffer.Length + 4
    Buffer.WriteLong SMsgMapData
    Buffer.WriteBytes MapData.ToArray()

    SendData Buffer.ToArray()
End Sub

Sub SendNeedMap(ByVal Revision As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgNeedMap
    Buffer.WriteLong Revision
    
    SendData Buffer.ToArray()
End Sub

Sub SendMapGetItem()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgMapGetItem
    
    SendData Buffer.ToArray()
End Sub

Sub SendDropItem(ByVal InvNum As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer

    ' Doesnt' matter if they are dead
    If Player(MyIndex).IsDead Then Exit Sub
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong SMsgMapDropItem
    Buffer.WriteLong InvNum
    Buffer.WriteLong Amount
    
    SendData Buffer.ToArray()
End Sub

Sub SendMapRespawn()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgMapRespawn
    
    SendData Buffer.ToArray()
End Sub

Sub SendMapReport()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgMapReport
    
    SendData Buffer.ToArray()
End Sub

Sub SendKick(ByVal Name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong SMsgKickPlayer
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Sub SendBanList()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgListBans
    
    SendData Buffer.ToArray()
End Sub

Sub SendBanDestroy()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgBanDestroy
    
    SendData Buffer.ToArray()
End Sub

Sub SendBan(ByVal Name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong SMsgBanPlayer
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Sub SendRequestEditMap()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgRequestEditMap
    
    SendData Buffer.ToArray()
End Sub

'
' Items
'
Sub SendRequestEditItem()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgRequestEditItem
    
    SendData Buffer.ToArray()
End Sub

Sub SendEditItem()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgEditItem
    Buffer.WriteLong EditorIndex
    
    SendData Buffer.ToArray()
End Sub

Sub SendSaveItem(ByVal ItemNum As Long)
Dim Buffer As clsBuffer
Dim ItemData() As Byte
Dim ItemSize As Long

    Set Buffer = New clsBuffer
    
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(0 To ItemSize - 1)

    Buffer.PreAllocate ItemSize + 8
    Buffer.WriteLong SMsgSaveItem
    Buffer.WriteLong ItemNum
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteBytes ItemData
    
    SendData Buffer.ToArray()
End Sub

'
' Npcs
'
Sub SendRequestEditNpc()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgRequestEditNpc
    
    SendData Buffer.ToArray()
End Sub

Sub SendEditNpc()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgEditNpc
    Buffer.WriteLong EditorIndex
    
    SendData Buffer.ToArray()
End Sub

Sub SendSaveNpc(ByVal NpcNum As Long)
Dim Buffer As clsBuffer
Dim NpcData() As Byte
Dim NpcSize As Long

    Set Buffer = New clsBuffer
    
    NpcSize = LenB(Npc(NpcNum))
    ReDim NpcData(0 To NpcSize - 1)

    Buffer.PreAllocate NpcSize + 8
    Buffer.WriteLong SMsgSaveNpc
    Buffer.WriteLong NpcNum
    CopyMemory NpcData(0), ByVal VarPtr(Npc(NpcNum)), NpcSize
    Buffer.WriteBytes NpcData
    
    SendData Buffer.ToArray()
End Sub

'
' Shops
'
Sub SendRequestEditShop()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgRequestEditShop
    
    SendData Buffer.ToArray()
End Sub

Sub SendEditShop()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgEditShop
    Buffer.WriteLong EditorIndex
    
    SendData Buffer.ToArray()
End Sub

Sub SendSaveShop(ByVal ShopNum As Long)
Dim Buffer As clsBuffer
Dim ShopData() As Byte
Dim ShopSize As Long

    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(0 To ShopSize - 1)

    Buffer.PreAllocate ShopSize + 8
    Buffer.WriteLong SMsgSaveShop
    Buffer.WriteLong ShopNum
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    Buffer.WriteBytes ShopData
    
    SendData Buffer.ToArray()
End Sub

'
' Spell
'
Sub SendRequestEditSpell()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgRequestEditSpell
    
    SendData Buffer.ToArray()
End Sub

Sub SendEditSpell()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgEditSpell
    Buffer.WriteLong EditorIndex
    
    SendData Buffer.ToArray()
End Sub

Sub SendSaveSpell(ByVal SpellNum As Long)
Dim Buffer As clsBuffer
'Dim SpellData() As Byte
'Dim SpellSize As Long

    Set Buffer = New clsBuffer
    
'    SpellSize = LenB(Spell(SpellNum))
'    ReDim SpellData(0 To SpellSize - 1)

    Buffer.PreAllocate SpellSize + 8
    Buffer.WriteLong SMsgSaveSpell
    Buffer.WriteLong SpellNum
'    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
'    Buffer.WriteBytes SpellData
    Buffer.WriteBytes Get_SpellData(SpellNum)
    
    SendData Buffer.ToArray()
End Sub

'
' Animation
'
Sub SendRequestEditAnimation()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgRequestEditAnimation
    
    SendData Buffer.ToArray()
End Sub

Sub SendEditAnimation()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgEditAnimation
    Buffer.WriteLong EditorIndex
    
    SendData Buffer.ToArray()
End Sub

Sub SendSaveAnimation(ByVal AnimationNum As Long)
Dim Buffer As clsBuffer
'Dim AnimationData() As Byte
'Dim AnimationSize As Long

    Set Buffer = New clsBuffer
'
'    AnimationSize = LenB(Animation(AnimationNum))
'    ReDim AnimationData(0 To AnimationSize - 1)

    Buffer.PreAllocate AnimationSize + 8
    Buffer.WriteLong SMsgSaveAnimation
    Buffer.WriteLong AnimationNum
    'CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
'    Buffer.WriteBytes AnimationData
    Buffer.WriteBytes Get_AnimationData(AnimationNum)
    
    SendData Buffer.ToArray()
End Sub

Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 9
    Buffer.WriteLong SMsgSetAccess
    Buffer.WriteString Name
    Buffer.WriteByte Access
    
    SendData Buffer.ToArray()
End Sub

Sub SendWhosOnline()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgWhosOnline
    
    SendData Buffer.ToArray()
End Sub

Sub SendMOTDChange(ByVal MOTD As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(MOTD) + 8
    Buffer.WriteLong SMsgSetMOTD
    Buffer.WriteString MOTD
    
    SendData Buffer.ToArray()
End Sub

Sub SendTradeRequest(ByVal TradeSlot As Long)
Dim Buffer As clsBuffer

    ' Doesnt' matter if they are dead
    If Player(MyIndex).IsDead Then Exit Sub
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 16
    Buffer.WriteLong SMsgTradeRequest
    Buffer.WriteLong ShopNpcNum
    Buffer.WriteLong InShop
    Buffer.WriteLong TradeSlot
    
    SendData Buffer.ToArray()
End Sub

Sub SendSearch(ByVal X As Long, ByVal Y As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong SMsgSearch
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    SendData Buffer.ToArray()
End Sub

Sub SendPartyRequest(ByVal Name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong SMsgParty
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Sub SendJoinParty()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgJoinParty
    
    SendData Buffer.ToArray()
End Sub

Sub SendLeaveParty()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgLeaveParty
    
    SendData Buffer.ToArray()
End Sub

Sub SendCast(ByVal SpellSlot As Long)
Dim Buffer As clsBuffer

    ' Doesnt' matter if they are dead
    If Player(MyIndex).IsDead Then Exit Sub
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgCast
    Buffer.WriteLong SpellSlot
    
    SendData Buffer.ToArray()
End Sub

Sub SendRequestLocation()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgRequestLocation
    
    SendData Buffer.ToArray()
End Sub

Sub SendFix()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgFix
    
    SendData Buffer.ToArray()
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Byte, ByVal NewSlot As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong SMsgChangeInvSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    
    SendData Buffer.ToArray()
End Sub

Sub SendClearTarget()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgClearTarget
    
    SendData Buffer.ToArray()
End Sub

Sub SendGCreate(ByVal Name As String, ByVal Abbreviation As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + Len(Abbreviation) + 12
    Buffer.WriteLong SMsgGCreate
    Buffer.WriteString Name
    Buffer.WriteString Abbreviation
    
    SendData Buffer.ToArray()
End Sub

Sub SendSetGMOTD(ByVal GMOTD As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(GMOTD) + 8
    Buffer.WriteLong SMsgSetGMOTD
    Buffer.WriteString GMOTD
    
    SendData Buffer.ToArray()
End Sub

Sub SendGQuit()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgGQuit
    
    SendData Buffer.ToArray()
End Sub

Sub SendGDelete()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgGDelete
    
    SendData Buffer.ToArray()
End Sub

Sub SendGPromote(ByVal Name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong SMsgGPromote
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Sub SendGDemote(ByVal Name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong SMsgGDemote
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Sub SendGKick(ByVal Name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong SMsgGKick
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Sub SendGInvite(ByVal Name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong SMsgGInvite
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Sub SendGJoin()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgGJoin
    
    SendData Buffer.ToArray()
End Sub

Sub SendGDecline()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgGDecline
    
    SendData Buffer.ToArray()
End Sub

Sub GuildMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Text) + 8
    Buffer.WriteLong SMsgGuildMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
End Sub

Sub SendKill(ByVal Name As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(Name) + 8
    Buffer.WriteLong SMsgKill
    Buffer.WriteString Name
    
    SendData Buffer.ToArray()
End Sub

Sub SendSetBound()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 12
    Buffer.WriteLong SMsgSetBound
    Buffer.WriteLong ShopNpcNum
    Buffer.WriteLong InShop
    
    SendData Buffer.ToArray()
End Sub

Sub SendCancelSpell()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgCancelSpell
    
    SendData Buffer.ToArray()
End Sub

Sub SendRelease()
Dim Buffer As clsBuffer

    ' If not dead exit
    If Not Current_IsDead(MyIndex) Then Exit Sub
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgRelease
    
    SendData Buffer.ToArray()
End Sub

Sub SendRevive()
Dim Buffer As clsBuffer

    ' If not dead exit
    If Not Current_IsDead(MyIndex) Then Exit Sub
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgRevive
    
    SendData Buffer.ToArray()
End Sub
