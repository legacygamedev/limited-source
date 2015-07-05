Attribute VB_Name = "modClientTCP"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Sub TcpInit()
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    PlayerBuffer = ""
        
    frmCClient.Socket.RemoteHost = GAME_IP
    frmCClient.Socket.RemotePort = GAME_PORT
End Sub

Sub TcpDestroy()
    frmCClient.Socket.Close
    
    If frmMainMenu.picChars.Visible Then frmMainMenu.picChars.Visible = False
    If frmMainMenu.picCreditsMenu.Visible Then frmMainMenu.picCreditsMenu.Visible = False
    If frmMainMenu.picDeleteAccountMenu.Visible Then frmMainMenu.picDeleteAccountMenu.Visible = False
    If frmMainMenu.picLoginMenu.Visible Then frmMainMenu.picLoginMenu.Visible = False
    If frmMainMenu.picNewAccountMenu.Visible Then frmMainMenu.picNewAccountMenu.Visible = False
    If frmMainMenu.picCreateChar.Visible Then frmMainMenu.picCreateChar.Visible = False
End Sub

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim Top As String * 3
Dim Start As Integer

    frmCClient.Socket.GetData Buffer, vbString, DataLength
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
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    With frmCClient.Socket
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
    If frmCClient.Socket.State = sckConnected Then
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
    Dim dbytes() As Byte
       
        dbytes = StrConv(Data, vbFromUnicode)
        If IsConnected Then
            frmCClient.Socket.SendData dbytes
            DoEvents
        End If
    End Sub

Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim Packet As String

    Packet = "newaccount" & SEP_CHAR & Trim(Name) & SEP_CHAR & Trim(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
Dim Packet As String
    
    Packet = "delaccount" & SEP_CHAR & Trim(Name) & SEP_CHAR & Trim(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim Packet As String

    Packet = "login" & SEP_CHAR & Trim(Name) & SEP_CHAR & Trim(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
Dim Packet As String

    Packet = "addchar" & SEP_CHAR & Trim(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
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

Sub SendMapRespawn()
Dim Packet As String

    Packet = "MAPRESPAWN" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendWhosOnline()
Dim Packet As String

    Packet = "WHOSONLINE" & SEP_CHAR & END_CHAR
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
Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long
    
    With Map
        Packet = "MAPDATA" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim(.Name) & SEP_CHAR & .Revision & SEP_CHAR & Trim(.Owner) & SEP_CHAR & .Moral & SEP_CHAR & .Up & SEP_CHAR & .Down & SEP_CHAR & .Left & SEP_CHAR & .Right & SEP_CHAR & .Music & SEP_CHAR & .BootMap & SEP_CHAR & .BootX & SEP_CHAR & .BootY & SEP_CHAR & .Indoors & SEP_CHAR
    End With
    
    For x = 1 To MAX_MAP_NPCS
        With Map
            Packet = Packet & .NSpawn(x).NSx & SEP_CHAR & .NSpawn(x).NSy & SEP_CHAR
        End With
    Next x
    
    For x = 1 To MAX_MAP_RESOURCES
        With Map
            Packet = Packet & .RSpawn(x).RSx & SEP_CHAR & .RSpawn(x).RSy & SEP_CHAR
        End With
    Next x
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map.Tile(x, y)
                Packet = Packet & .Tileset & SEP_CHAR & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Mask2 & SEP_CHAR & .Anim & SEP_CHAR & .Fringe & SEP_CHAR & .Fringe2 & SEP_CHAR & .FAnim & SEP_CHAR & .Light & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .WalkUp & SEP_CHAR & .WalkDown & SEP_CHAR & .WalkLeft & SEP_CHAR & .WalkRight & SEP_CHAR & .Build & SEP_CHAR
            End With
        Next x
    Next y
    
    With Map
        For x = 1 To MAX_MAP_NPCS
            Packet = Packet & .Npc(x) & SEP_CHAR
        Next x
    End With
    
    With Map
        For x = 1 To MAX_MAP_RESOURCES
            Packet = Packet & .Resource(x) & SEP_CHAR
        Next x
    End With
    
    Packet = Packet & END_CHAR
    
    x = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, x)
    P2 = Mid(Packet, x + 1, Len(Packet) - x)
    Call SendData(Packet)
End Sub

Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim Packet As String

    Packet = "SETACCESS" & SEP_CHAR & Name & SEP_CHAR & Access & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditItem()
Dim Packet As String

    Packet = "REQUESTEDITITEM" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Public Sub SendSaveItem(ByVal ItemNum As Long)
Dim Packet As String
    
    With Item(ItemNum)
        Packet = "SAVEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(.Name) & SEP_CHAR & .Pic & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & END_CHAR
    End With
    
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
Dim n As Long
    
    Packet = "SAVENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Respawn & SEP_CHAR & Npc(NpcNum).HitOnlyWith & SEP_CHAR & Npc(NpcNum).ShopLink & SEP_CHAR & Npc(NpcNum).ExpType & SEP_CHAR & Npc(NpcNum).EXP & SEP_CHAR
    For i = 1 To MAX_NPC_QUESTS
        Packet = Packet & Npc(NpcNum).QuestNPC(i) & SEP_CHAR
    Next i
    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(NpcNum).ItemNPC(i).Chance
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditShop()
Dim Packet As String

    Packet = "REQUESTEDITSHOP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Public Sub SendSaveShop(ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long
Dim f As Long

    With Shop(ShopNum)
        Packet = "SAVESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(.Name) & SEP_CHAR & .FixesItems & SEP_CHAR
    End With
    
    For i = 1 To MAX_TRADES
        With Shop(ShopNum).TradeItem(i)
            For f = 1 To MAX_GIVE_ITEMS
                Packet = Packet & .GiveItem(f) & SEP_CHAR
            Next f
            For f = 1 To MAX_GIVE_VALUE
                Packet = Packet & .GiveValue(f) & SEP_CHAR
            Next f
            For f = 1 To MAX_GET_ITEMS
                Packet = Packet & .GetItem(f) & SEP_CHAR
            Next f
            For f = 1 To MAX_GET_VALUE
                Packet = Packet & .GetValue(f) & SEP_CHAR
            Next f
        End With
        Packet = Packet & Shop(ShopNum).ItemStock(i) & SEP_CHAR
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
Dim Packet As String

    With Spell(SpellNum)
        Packet = "SAVESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(.Name) & SEP_CHAR & .SpellSprite & SEP_CHAR & .ClassReq & SEP_CHAR & .LevelReq & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & END_CHAR
    End With
    
    Call SendData(Packet)
End Sub

Sub SendRequestEditSkill()
Dim Packet As String

    Packet = "REQUESTEDITSKILL" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Public Sub SendSaveSkill(ByVal SkillNum As Long)
Dim Packet As String

    With Skill(SkillNum)
        Packet = "SAVESKILL" & SEP_CHAR & SkillNum & SEP_CHAR & Trim(.Name) & SEP_CHAR & .SkillSprite & SEP_CHAR & .ClassReq & SEP_CHAR & .LevelReq & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & END_CHAR
    End With
    
    Call SendData(Packet)
End Sub

Sub SendRequestEditQuest()
Dim Packet As String

    Packet = "REQUESTEDITQUEST" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Public Sub SendSaveQuest(ByVal QuestNum As Long)
Dim Packet As String

    With Quest(QuestNum)
        Packet = "SAVEQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim(.Name) & SEP_CHAR & .SetBy & SEP_CHAR & .ClassReq & SEP_CHAR & .LevelMin & SEP_CHAR & .LevelMax & SEP_CHAR & .Type & SEP_CHAR & .Reward & SEP_CHAR & .RewardValue & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & Trim(.Description) & SEP_CHAR & END_CHAR
    End With
    
    Call SendData(Packet)
End Sub

Sub SendRequestEditGUI()
Dim Packet As String

    Packet = "REQUESTEDITGUI" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Public Sub SendSaveGUI(ByVal GUINum As Long)
Dim Packet As String
Dim i As Long

    With GUI(GUINum)
        Packet = "SAVEGUI" & SEP_CHAR & GUINum & SEP_CHAR & Trim(.Name) & SEP_CHAR & Trim(.Designer) & SEP_CHAR & .Revision
        For i = 1 To 7
            Packet = Packet & SEP_CHAR & .Background(i).Data1 & SEP_CHAR & .Background(i).Data2 & SEP_CHAR & .Background(i).Data3 & SEP_CHAR & .Background(i).Data4 & SEP_CHAR & .Background(i).Data5
        Next i
        For i = 1 To 5
            Packet = Packet & SEP_CHAR & .Menu(i).Data1 & SEP_CHAR & .Menu(i).Data2 & SEP_CHAR & .Menu(i).Data3 & SEP_CHAR & .Menu(i).Data4
        Next i
        For i = 1 To 4
            Packet = Packet & SEP_CHAR & .Login(i).Data1 & SEP_CHAR & .Login(i).Data2 & SEP_CHAR & .Login(i).Data3 & SEP_CHAR & .Login(i).Data4
        Next i
        For i = 1 To 4
            Packet = Packet & SEP_CHAR & .NewAcc(i).Data1 & SEP_CHAR & .NewAcc(i).Data2 & SEP_CHAR & .NewAcc(i).Data3 & SEP_CHAR & .NewAcc(i).Data4
        Next i
        For i = 1 To 4
            Packet = Packet & SEP_CHAR & .DelAcc(i).Data1 & SEP_CHAR & .DelAcc(i).Data2 & SEP_CHAR & .DelAcc(i).Data3 & SEP_CHAR & .DelAcc(i).Data4
        Next i
        For i = 1 To 2
            Packet = Packet & SEP_CHAR & .Credits(i).Data1 & SEP_CHAR & .Credits(i).Data2 & SEP_CHAR & .Credits(i).Data3 & SEP_CHAR & .Credits(i).Data4
        Next i
        For i = 1 To 5
            Packet = Packet & SEP_CHAR & .Chars(i).Data1 & SEP_CHAR & .Chars(i).Data2 & SEP_CHAR & .Chars(i).Data3 & SEP_CHAR & .Chars(i).Data4
        Next i
        For i = 1 To 14
            Packet = Packet & SEP_CHAR & .NewChar(i).Data1 & SEP_CHAR & .NewChar(i).Data2 & SEP_CHAR & .NewChar(i).Data3 & SEP_CHAR & .NewChar(i).Data4
        Next i
        Packet = Packet & SEP_CHAR & END_CHAR
    End With
    
    Call SendData(Packet)
End Sub

Sub SendRequestEditMap()
Dim Packet As String

    Packet = "REQUESTEDITMAP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
