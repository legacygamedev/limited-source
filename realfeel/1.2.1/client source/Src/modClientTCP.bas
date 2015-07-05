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

'Check data to prevent error
If GetVar(App.Path & "\data.dat", "Address", "IP") = "" Then
    Call PutVar(App.Path & "\data.dat", "Address", "IP", CStr(frmDualSolace.Socket.LocalIP))
Else
    GAME_IP = GetVar(App.Path & "\data.dat", "Address", "IP")
End If
If GetVar(App.Path & "\data.dat", "Address", "Port") = "" Then
    Call PutVar(App.Path & "\data.dat", "Address", "Port", "2000")
Else
    GAME_PORT = GetVar(App.Path & "\data.dat", "Address", "Port")
End If

    frmDualSolace.Socket.RemoteHost = GAME_IP
    frmDualSolace.Socket.RemotePort = GAME_PORT
End Sub

Sub TcpDestroy()
    frmDualSolace.Socket.Close
    
    If frmChars.Visible Then frmChars.Visible = False
    If frmCredits.Visible Then frmCredits.Visible = False
    If frmDeleteAccount.Visible Then frmDeleteAccount.Visible = False
    If frmNewAccount.Visible Then frmNewAccount.Visible = False
    If frmNewChar.Visible Then frmNewChar.Visible = False
End Sub

' Converts the binary to a string
Public Function CBuf(bData() As Byte, ByVal Total As Integer) As String
Dim n As Long
CBuf = ""

For n = 0 To Total - 1
CBuf = CBuf & ChrW(bData(n))
Next n

End Function

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim top As String * 3
Dim Start As Integer
Dim StringData As String
Dim ByteData() As Byte

    frmDualSolace.Socket.GetData ByteData, vbByte + vbArray, DataLength
    ' convert the byte buffer
    StringData = CBuf(ByteData(), DataLength)
    
    'Dim i As Long
    'For i = 0 To DataLength - 1
    '    Debug.Print "BYTEARRAY(" & i & "): " & ByteData(i)
    'Next i
    
    'Debug.Print "STRINGPACKET: " & StringData
    
    PlayerBuffer = PlayerBuffer & StringData
        
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
    With frmDualSolace.Socket
        .Close
        .Connect
    End With
    
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
    If frmDualSolace.Socket.State = sckConnected Then
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
        frmDualSolace.Socket.SendData Data
        DoEvents
    End If
End Sub

Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim Packet As String

    Packet = "newaccount" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
Dim Packet As String
    
    Packet = "delaccount" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim Packet As String

    Packet = "login" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
Dim Packet As String

    Packet = "addchar" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
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

Sub GetClassData()
Dim Packet As String

    Packet = "getclassdata" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseChar(ByVal CharSlot As Long)
Dim Packet As String

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

Public Sub SendMap()
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************

Dim Packet As String, P1 As String, P2 As String
Dim X As Long
Dim Y As Long
    
    'First, we will send the pause map notice
    Call SendData("PAUSEMAP" & SEP_CHAR & "LOCK" & SEP_CHAR & "Updating map..." & SEP_CHAR & END_CHAR)
    
    With Map
        Packet = "MAPDATA" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .Revision & SEP_CHAR & .Moral & SEP_CHAR & .Up & SEP_CHAR & .Down & SEP_CHAR & .Left & SEP_CHAR & .Right & SEP_CHAR & .Music & SEP_CHAR & .BootMap & SEP_CHAR & .BootX & SEP_CHAR & .BootY & SEP_CHAR & .Shop & SEP_CHAR
    End With
    
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map.Tile(X, Y)
                Packet = Packet & .Ground & SEP_CHAR
                Packet = Packet & .Mask & SEP_CHAR
                Packet = Packet & .Mask2 & SEP_CHAR
                Packet = Packet & .Anim & SEP_CHAR
                Packet = Packet & .Anim2 & SEP_CHAR
                Packet = Packet & .Fringe & SEP_CHAR
                Packet = Packet & .FringeAnim & SEP_CHAR
                Packet = Packet & .Fringe2 & SEP_CHAR
            End With
        Next X
    Next Y
    
    With Map
        For X = 1 To MAX_MAP_NPCS
            Packet = Packet & .Npc(X) & SEP_CHAR
        Next X
    End With
    
    Packet = Packet & END_CHAR
    
    X = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, X)
    P2 = Mid(Packet, X + 1, Len(Packet) - X)
    Call SendData(Packet)
    Call SendMapAttributes
    Call SendData("NOWSAVEMAP" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & END_CHAR)
End Sub

Sub SendMapAttributes()
'On Error GoTo errorhandler:
Dim Packet As String
Dim X As Long
Dim Y As Long

    Packet = "MAPATTRIBUTES" & SEP_CHAR
    
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map.Tile(X, Y)
                Packet = Packet & .Walkable & SEP_CHAR
                Packet = Packet & .Blocked & SEP_CHAR
                Packet = Packet & .Warp & SEP_CHAR & .WarpMap & SEP_CHAR & .WarpX & SEP_CHAR & .WarpY & SEP_CHAR
                Packet = Packet & .Item & SEP_CHAR & .ItemNum & SEP_CHAR & .ItemValue & SEP_CHAR
                Packet = Packet & .NpcAvoid & SEP_CHAR
                Packet = Packet & .Key & SEP_CHAR & .KeyNum & SEP_CHAR & .KeyOpen & SEP_CHAR
                Packet = Packet & .KeyOpen & SEP_CHAR & .KeyOpenX & SEP_CHAR & .KeyOpenY & SEP_CHAR
                Packet = Packet & .North & SEP_CHAR
                Packet = Packet & .West & SEP_CHAR
                Packet = Packet & .East & SEP_CHAR
                Packet = Packet & .South & SEP_CHAR
                Packet = Packet & .Shop & SEP_CHAR & .ShopNum & SEP_CHAR
                Packet = Packet & .Bank & SEP_CHAR
                Packet = Packet & .Heal & SEP_CHAR & .HealValue & SEP_CHAR
                Packet = Packet & .Damage & SEP_CHAR & .DamageValue & SEP_CHAR
            End With
        Next X
    Next Y

    Packet = Packet & END_CHAR
    Call SendData(Packet)

ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  'Call ReportError("modServerTCP.bas", "SendMapAttributes", Err.Number, Err.Description)
End Sub

Sub WarpMeTo(ByVal Name As String)
Dim Packet As String

    Packet = "WARPMETO" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpToMe(ByVal Name As String)
Dim Packet As String

    Packet = "WARPTOME" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpTo(ByVal MapNum As Long)
Dim Packet As String
    
    Packet = "WARPTO" & SEP_CHAR & MapNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim Packet As String

    Packet = "SETACCESS" & SEP_CHAR & Name & SEP_CHAR & Access & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetSprite(ByVal SpriteNum As Long)
Dim Packet As String

    Packet = "SETSPRITE" & SEP_CHAR & SpriteNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendKick(ByVal Name As String)
Dim Packet As String

    Packet = "KICKPLAYER" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBan(ByVal Name As String)
Dim Packet As String

    Packet = "BANPLAYER" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
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

Public Sub SendSaveItem(ByVal ItemNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************

Dim Packet As String
    
    With Item(ItemNum)
        Packet = "SAVEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & Trim$(.Description) & SEP_CHAR & .Pic & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .Data4 & SEP_CHAR & .Data5 & SEP_CHAR & .Sound & SEP_CHAR & END_CHAR
    End With
    
    Call SendData(Packet)
End Sub
                
Sub SendRequestEditNpc()
Dim Packet As String

    Packet = "REQUESTEDITNPC" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Public Sub SendSaveNpc(ByVal NpcNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************

Dim Packet As String
    
    With Npc(NpcNum)
        Packet = "SAVENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & Trim$(.AttackSay) & SEP_CHAR & .Sprite & SEP_CHAR & .SpawnSecs & SEP_CHAR & .Behavior & SEP_CHAR & .Range & SEP_CHAR & .DropChance & SEP_CHAR & .DropItem & SEP_CHAR & .DropItemValue & SEP_CHAR & .HP & SEP_CHAR & .STR & SEP_CHAR & .DEF & SEP_CHAR & .Speed & SEP_CHAR & .MAGI & SEP_CHAR & .EXP & SEP_CHAR & .Fear & SEP_CHAR & .TintR & SEP_CHAR & .TintG & SEP_CHAR & .TintB & SEP_CHAR & END_CHAR
    End With
    Debug.Print "Save: Fear = " & Npc(NpcNum).Fear
    
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

Public Sub SendSaveShop(ByVal ShopNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************

Dim Packet As String
Dim i As Long

    With Shop(ShopNum)
        Packet = "SAVESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & Trim$(.JoinSay) & SEP_CHAR & Trim$(.LeaveSay) & SEP_CHAR & .FixesItems & SEP_CHAR & .Restock & SEP_CHAR
    End With
    
    For i = 1 To MAX_TRADES
        With Shop(ShopNum).TradeItem(i)
            If .Stock = 0 Or .MaxStock = 0 Then
                If IsNumeric(Mid(frmEditor.lstTradeItem.List(i - 1), (Len(frmEditor.lstTradeItem.List(i - 1)) - 1), 1)) Then
                    .Stock = CLng(Mid(frmEditor.lstTradeItem.List(i - 1), (Len(frmEditor.lstTradeItem.List(i - 1)) - 1), 1))
                    .MaxStock = CLng(Mid(frmEditor.lstTradeItem.List(i - 1), (Len(frmEditor.lstTradeItem.List(i - 1)) - 1), 1))
                End If
            End If
            Packet = Packet & .GiveItem & SEP_CHAR & .GiveValue & SEP_CHAR & .GetItem & SEP_CHAR & .GetValue & SEP_CHAR & .Stock & SEP_CHAR & .MaxStock & SEP_CHAR
        End With
    Next i
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditSpell()
Dim Packet As String

    Packet = "REQUESTEDITSPELL" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Public Sub SendSaveSpell(ByVal SpellNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************

Dim Packet As String

    With Spell(SpellNum)
        Packet = "SAVESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .ClassReq & SEP_CHAR & .LevelReq & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & END_CHAR
    End With
    
    Call SendData(Packet)
End Sub

Sub SendRequestEditClass()
Dim Packet As String

    Packet = "REQUESTEDITCLASS" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Public Sub SendSaveClass(ByVal ClassNum As Byte)
Dim Packet As String

    With Class(ClassNum)
        Packet = "SAVECLASS" & SEP_CHAR & ClassNum & SEP_CHAR & Trim$(.Name) & SEP_CHAR & .Sprite & SEP_CHAR & .HP & SEP_CHAR & .MP & SEP_CHAR & .SP & SEP_CHAR & .STR & SEP_CHAR & .DEF & SEP_CHAR & .MAGI & SEP_CHAR & .Speed & SEP_CHAR & .Map & SEP_CHAR & .X & SEP_CHAR & .Y & SEP_CHAR & END_CHAR
    End With
    
    Call SendData(Packet)
End Sub

Sub SendRequestEditMap()
Dim Packet As String

    Packet = "REQUESTEDITMAP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPartyRequest(ByVal Name As String)
Dim Packet As String

    Packet = "PARTY" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
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

Sub SendAddTracker(ByVal Name As String)
Dim Packet As String
    TrackName = Trim$(Name)
    Packet = "ADDTRACKER" & SEP_CHAR & Trim$(Name) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRemoveTracker(ByVal Name As String)
Dim Packet As String
    TrackName = ""
    Packet = "REMOVETRACKER" & SEP_CHAR & Trim$(Name) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

