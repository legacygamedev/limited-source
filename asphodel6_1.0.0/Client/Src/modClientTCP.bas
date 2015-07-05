Attribute VB_Name = "modClientTCP"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' -- Communcation to server, TCP          --
' -- Winsock Control (mswinsck.ocx)       --
' -- String packets (slow and big)        --
' ------------------------------------------

Public Sub TcpInit()

    ' used for parsing packets
    SEP_CHAR = vbNullChar ' ChrW$(0)
    END_CHAR = ChrW$(237)
    
    ' check if IP is valid
    'If IsIP(GAME_IP) Then
        frmMainGame.Socket.RemoteHost = GAME_IP
        frmMainGame.Socket.RemotePort = GAME_PORT
    'Else
    '    MsgBox GAME_IP & " does not appear as a valid IP address!"
    '    DestroyGame
    'End If
    
End Sub

Public Sub DestroyTCP()
    frmMainGame.Socket.Close
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim Start As Long

    frmMainGame.Socket.GetData Buffer, vbString, DataLength
    PlayerBuffer = PlayerBuffer & Buffer
        
    Start = InStr(PlayerBuffer, END_CHAR)
    Do While Start > 0
        Packet = Mid$(PlayerBuffer, 1, Start - 1)
        PlayerBuffer = Mid$(PlayerBuffer, Start + 1, Len(PlayerBuffer))
        Start = InStr(PlayerBuffer, END_CHAR)
        
        If Len(Packet) > 0 Then
            HandleData Packet
        End If
    Loop
End Sub

Public Function ConnectToServer() As Boolean
Dim Wait As Currency
    
    ' Check to see if we are already connected, if so just exit
    If IsConnected Then: ConnectToServer = IsConnected: Exit Function
    
    With frmMainGame.Socket
        .Close
        .Connect
    End With
    
    Wait = GetTickCountNew + 2000
    
    ' Wait until connected or 2 seconds have passed and report the server being down
    Do Until Wait < GetTickCountNew
        If IsConnected Then Exit Do
        DoEvents
    Loop
    
    ' return value
    ConnectToServer = IsConnected

End Function

Private Function IsIP(ByVal IPAddress As String) As Boolean
Dim s() As String
Dim i As Long
    
    'If there are no periods, I have no idea what we have...
    If InStr(1, IPAddress, ".") = 0 Then Exit Function
    
    'Split up the string by the periods
    s = Split(IPAddress, ".")
    
    'Confirm we have ubound = 3, since xxx.xxx.xxx.xxx has 4 elements and we start at index 0
    If UBound(s) <> 3 Then Exit Function
    
    'Check that the values are numeric and in a valid range
    For i = 0 To 3
        If Not IsNumeric(s(i)) Then Exit Function
        If s(i) < 0 Then Exit Function
        If s(i) > 255 Then Exit Function
    Next
    
    'Looks like we were passed a valid IP!
    IsIP = True
    
End Function

Function IsConnected() As Boolean
    IsConnected = (frmMainGame.Socket.State = sckConnected)
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    IsPlaying = (LenB(GetPlayerName(Index)) > 0)
End Function

Public Sub SendData(ByVal Data As String)
    If IsConnected Then
        frmMainGame.Socket.SendData Data
        DoEvents
    End If
End Sub

' ******************************
' ** Outcoming Client Packets **
' ******************************

Public Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim Packet As String

    Packet = CNewAccount & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & END_CHAR
    SendData Packet
End Sub

Public Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim Packet As String

    Packet = CLogin & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & END_CHAR
    SendData Packet
End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
Dim Packet As String

    Packet = CAddChar & SEP_CHAR & Trim$(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & Slot & END_CHAR
    SendData Packet
End Sub

Public Sub SendDelChar(ByVal Slot As Long)
Dim Packet As String

    Packet = CDelChar & SEP_CHAR & Slot & END_CHAR
    SendData Packet
End Sub

Public Sub SendGetClasses()
Dim Packet As String

    Packet = CGetClasses & END_CHAR
    SendData Packet
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
Dim Packet As String

    Packet = CUseChar & SEP_CHAR & CharSlot & END_CHAR
    SendData Packet
End Sub

Public Sub SendMessage(ByVal ChatType As Byte, ByVal Message As String)
    SendData CMessage & SEP_CHAR & ChatType & SEP_CHAR & Message & END_CHAR
End Sub

Public Sub SendPlayerMove()
Dim Packet As String

    Packet = CPlayerMove & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & END_CHAR
    SendData Packet
End Sub

Public Sub SendPlayerDir()
Dim Packet As String

    Packet = CPlayerDir & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR
    SendData Packet
End Sub

Public Sub SendPlayerRequestNewMap()
Dim Packet As String
    
    Packet = CRequestNewMap & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR
    SendData Packet
End Sub

Public Sub SendMap()
Dim Packet As String
Dim X As Long
Dim Y As Long
Dim i As Long

    CanMoveNow = False
    
    With Map
        Packet = CMapData & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .Moral & SEP_CHAR & .Up & SEP_CHAR & .Down & SEP_CHAR & .Left & SEP_CHAR & .Right & SEP_CHAR & .Music & SEP_CHAR & .BootMap & SEP_CHAR & .BootX & SEP_CHAR & .BootY
    End With
    
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            With Map.Tile(X, Y)
                For i = 0 To UBound(.Layer)
                    Packet = Packet & SEP_CHAR & .Layer(i) & SEP_CHAR & .LayerSet(i)
                Next
                Packet = Packet & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3
            End With
        Next
    Next
    
    With MapSpawn
        Packet = Packet & SEP_CHAR & UBound(.Npc)
        For X = 1 To UBound(.Npc)
            Packet = Packet & SEP_CHAR & .Npc(X).Num & SEP_CHAR & .Npc(X).X & SEP_CHAR & .Npc(X).Y
        Next
    End With
    
    Packet = Packet & END_CHAR
    
    SendData Packet
    
End Sub

Public Sub WarpMeTo(ByVal Name As String)
Dim Packet As String

    Packet = CWarpMeTo & SEP_CHAR & Name & END_CHAR
    SendData Packet
End Sub

Public Sub WarpToMe(ByVal Name As String)
Dim Packet As String

    Packet = CWarpToMe & SEP_CHAR & Name & END_CHAR
    SendData Packet
End Sub

Public Sub WarpTo(ByVal MapNum As Long)
Dim Packet As String
    
    Packet = CWarpTo & SEP_CHAR & MapNum & END_CHAR
    SendData Packet
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim Packet As String

    Packet = CSetAccess & SEP_CHAR & Name & SEP_CHAR & Access & END_CHAR
    SendData Packet
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
Dim Packet As String

    Packet = CSetSprite & SEP_CHAR & SpriteNum & END_CHAR
    SendData Packet
End Sub

Public Sub SendKick(ByVal Name As String)
Dim Packet As String

    Packet = CKickPlayer & SEP_CHAR & Name & END_CHAR
    SendData Packet
End Sub

Public Sub SendBan(ByVal Name As String)
Dim Packet As String

    Packet = CBanPlayer & SEP_CHAR & Name & END_CHAR
    SendData Packet
End Sub

Public Sub SendBanList()
Dim Packet As String

    Packet = CBanList & END_CHAR
    SendData Packet
End Sub

Public Sub SendRequestEditItem()
Dim Packet As String

    Packet = CRequestEditItem & END_CHAR
    SendData Packet
End Sub

Public Sub SendSaveItem(ByVal ItemNum As Long)
Dim Packet As String
Dim LoopI As Long

    With Item(ItemNum)
        Packet = CSaveItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .Pic & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .Durability & SEP_CHAR & .Anim & SEP_CHAR & .CostItem & SEP_CHAR & .CostAmount
        For LoopI = 1 To Stats.Stat_Count - 1
            Packet = Packet & SEP_CHAR & .BuffStats(LoopI)
        Next
        For LoopI = 1 To Vitals.Vital_Count - 1
            Packet = Packet & SEP_CHAR & .BuffVitals(LoopI)
        Next
        For LoopI = 0 To Item_Requires.Count - 1
            Packet = Packet & SEP_CHAR & .Required(LoopI)
        Next
    End With
    
    Packet = Packet & END_CHAR
    
    SendData Packet
    
End Sub

Public Sub SendRequestEditNpc()
Dim Packet As String

    Packet = CRequestEditNpc & END_CHAR
    SendData Packet
End Sub

Public Sub SendSaveNpc(ByVal NpcNum As Long)
Dim Packet As String
Dim i As Long

    With Npc(NpcNum)
        
        Packet = CSaveNpc & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & Trim$(.AttackSay) & SEP_CHAR & .Sprite & SEP_CHAR & .SpawnSecs & SEP_CHAR & .Behavior & SEP_CHAR & .Range & SEP_CHAR & .DropChance & SEP_CHAR & .DropItem & SEP_CHAR & .DropItemValue & SEP_CHAR & .Stat(Stats.Strength) & SEP_CHAR & .Stat(Stats.Defense) & SEP_CHAR & .Stat(Stats.Speed) & SEP_CHAR & .Stat(Stats.Magic) & SEP_CHAR & .HP & SEP_CHAR & .Experience & SEP_CHAR & .GivesGuild
        
        For i = 0 To UBound(.Sound)
            Packet = Packet & SEP_CHAR & .Sound(i)
        Next
        
        For i = 0 To UBound(.Reflection)
            Packet = Packet & SEP_CHAR & .Reflection(i)
        Next
        
        Packet = Packet & END_CHAR
        
    End With
    
    SendData Packet
End Sub

Public Sub SendMapRespawn()
Dim Packet As String

    Packet = CMapRespawn & END_CHAR
    SendData Packet
End Sub

Public Sub SendUseItem(ByVal InvNum As Long)
Dim Packet As String

    Packet = CUseItem & SEP_CHAR & InvNum & END_CHAR
    SendData Packet
End Sub

Public Sub SendDropItem(ByVal InvNum As Long, ByVal Amount As Long)
Dim Packet As String

    Packet = CMapDropItem & SEP_CHAR & InvNum & SEP_CHAR & Amount & END_CHAR
    SendData Packet
End Sub

Public Sub SendWhosOnline()
Dim Packet As String

    Packet = CWhosOnline & END_CHAR
    SendData Packet
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim Packet As String

    Packet = CSetMotd & SEP_CHAR & MOTD & END_CHAR
    SendData Packet
End Sub

Public Sub SendRequestEditShop()
Dim Packet As String

    Packet = CRequestEditShop & END_CHAR
    SendData Packet
End Sub

Public Sub SendSaveShop(ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long

    With Shop(ShopNum)
        Packet = CSaveShop & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & Trim$(.JoinSay) & SEP_CHAR & Trim$(.LeaveSay) & SEP_CHAR & .FixesItems
    End With
    
    For i = 1 To MAX_TRADES
        With Shop(ShopNum).TradeItem(i)
            Packet = Packet & SEP_CHAR & .GiveItem & SEP_CHAR & .GiveValue & SEP_CHAR & .GetItem & SEP_CHAR & .GetValue
        End With
    Next
    
    Packet = Packet & END_CHAR
    SendData Packet
End Sub

Public Sub SendRequestEditAnim()
Dim Packet As String

    Packet = CRequestEditAnim & END_CHAR
    SendData Packet
    
End Sub

Public Sub SendRequestEditSign()
Dim Packet As String

    Packet = CRequestEditSign & END_CHAR
    SendData Packet
    
End Sub

Public Sub SendRequestEditSpell()
Dim Packet As String

    Packet = CRequestEditSpell & END_CHAR
    SendData Packet
End Sub

Public Sub SendSaveSpell(ByVal SpellNum As Long)
Dim Packet As String

    With Spell(SpellNum)
        Packet = CSaveSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .CastSound & SEP_CHAR & .MPReq & SEP_CHAR & .Type & SEP_CHAR & .Anim & SEP_CHAR & .Icon & SEP_CHAR & .Range & SEP_CHAR & .AOE & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .Timer & END_CHAR
    End With
    
    SendData Packet
    
End Sub

Public Sub SendSaveSign(ByVal SignNum As Long)
Dim Packet As String
Dim LoopI As Long

    With Sign(SignNum)
        Packet = CSaveSign & SEP_CHAR & SignNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & UBound(.Section)
        For LoopI = 0 To UBound(.Section)
            Packet = Packet & SEP_CHAR & SignSection(LoopI)
        Next
    End With
    
    Packet = Packet & END_CHAR
    
    SendData Packet
    
End Sub

Public Sub SendSaveAnim(ByVal AnimNum As Long)
Dim Packet As String

    With Animation(AnimNum)
        Packet = CSaveAnim & SEP_CHAR & AnimNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .Delay & SEP_CHAR & .Width & SEP_CHAR & .Height & SEP_CHAR & .Pic & END_CHAR
    End With
    
    SendData Packet
    
End Sub

Public Sub SendRequestEditMap()
Dim Packet As String

    Packet = CRequestEditMap & END_CHAR
    SendData Packet
End Sub

Public Sub SendPartyRequest(ByVal Name As String)
Dim Packet As String

    Packet = CParty & SEP_CHAR & Name & END_CHAR
    SendData Packet
End Sub

Public Sub SendJoinParty()
Dim Packet As String

    Packet = CJoinParty & END_CHAR
    SendData Packet
End Sub

Public Sub SendLeaveParty()
Dim Packet As String

    Packet = CLeaveParty & END_CHAR
    SendData Packet
End Sub

Public Sub SendBanDestroy()
Dim Packet As String
    
    Packet = CBanDestroy & END_CHAR
    SendData Packet
End Sub

