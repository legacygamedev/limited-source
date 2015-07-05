Attribute VB_Name = "modHandleData"
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

Public Sub HandleData(ByVal Data As String)
Dim Parse() As String
Dim Name As String
Dim Password As String
Dim Msg As String
Dim Dir As Long
Dim Level As Long
Dim i As Long, n As Long, x As Long, y As Long, f As Long, z As Long
Dim ShopNum As Long
Dim TradeDescription As String
'Dim GiveItem(1 To MAX_GIVE_ITEMS) As Byte
'Dim GiveValue(1 To MAX_GIVE_ITEMS) As Byte
'Dim GetItem(1 To MAX_GET_ITEMS) As Byte
'Dim GetValue(1 To MAX_GET_VALUE) As Byte

    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
    
    ' ::::::::::::::::::
    ' :: Get Max Info ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "maxinfo" Then
        GAME_NAME = Trim(Parse(1))
        GAME_WEBSITE = Trim(Parse(2))
        MAX_PLAYERS = Val(Parse(3))
        MAX_ITEMS = Val(Parse(4))
        MAX_NPCS = Val(Parse(5))
        MAX_SHOPS = Val(Parse(6))
        MAX_SPELLS = Val(Parse(7))
        MAX_SKILLS = Val(Parse(8))
        MAX_MAPS = Val(Parse(9))
        MAX_MAP_ITEMS = Val(Parse(10))
        MAX_GUILDS = Val(Parse(11))
        MAX_GUILD_MEMBERS = Val(Parse(12))
        'MAX_MAPX = Val(Parse(9))
        'MAX_MAPY = Val(Parse(10))
        'MAX_EMOTICONS = Val(Parse(11))
        MAX_QUESTS = Val(Parse(13))

        ReDim Player(1 To MAX_PLAYERS) As PlayerRec
        ReDim Item(1 To MAX_ITEMS) As ItemRec
        ReDim Npc(1 To MAX_NPCS) As NpcRec
        ReDim MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim Shop(1 To MAX_SHOPS) As ShopRec
        ReDim Spell(1 To MAX_SPELLS) As SpellRec
        ReDim Skill(1 To MAX_SKILLS) As SkillRec
        'ReDim Bubble(1 To MAX_PLAYERS) As ChatBubble
        ReDim SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        'ReDim SaveMap.Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        'ReDim Map.Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        'ReDim TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
        'ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
        ReDim Quest(1 To MAX_QUESTS) As QuestRec
        'ReDim MapReport(1 To MAX_MAPS) As MapRec
        
        'For i = 0 To MAX_EMOTICONS
            'Emoticons(i).Pic = 0
            'Emoticons(i).Command = ""
        'Next i
        
        Call ClearTempTile
        Call ClearPushTile
        
        ' Clear out players
        For i = 1 To MAX_PLAYERS
            Call ClearPlayer(i)
        Next i
    
        frmCClient.Caption = Trim(GAME_NAME)
        App.Title = GAME_NAME
 
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "alertmsg" Then
        frmSendGetData.Visible = False
        frmMainMenu.Visible = True
        
        Msg = Parse(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "allchars" Then
        n = 1
        
        frmMainMenu.picChars.Visible = True
        frmSendGetData.Visible = False
        
        frmMainMenu.lstChars.Clear
        
        For i = 1 To MAX_CHARS
            Name = Parse(n)
            Msg = Parse(n + 1)
            Level = Val(Parse(n + 2))
            
            If Trim(Name) = "" Then
                frmMainMenu.lstChars.AddItem "Free Character Slot"
            Else
                frmMainMenu.lstChars.AddItem Name & " a level " & Level & " " & Msg
            End If
            
            n = n + 3
        Next i
        
        frmMainMenu.lstChars.ListIndex = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "loginok" Then
        ' Now we can receive game data
        MyIndex = Val(Parse(1))
        
        frmSendGetData.Visible = True
        frmMainMenu.picChars.Visible = False
        
        Call SetStatus("Receiving game data...")
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "newcharclasses" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For i = 0 To Max_Classes
            Class(i).Name = Parse(n)
            
            Class(i).HP = Val(Parse(n + 1))
            Class(i).MP = Val(Parse(n + 2))
            Class(i).SP = Val(Parse(n + 3))
            
            Class(i).STR = Val(Parse(n + 4))
            Class(i).DEF = Val(Parse(n + 5))
            Class(i).SPEED = Val(Parse(n + 6))
            Class(i).MAGI = Val(Parse(n + 7))
            Class(i).DEX = Val(Parse(n + 8))
            
            n = n + 9
        Next i
        
        ' Used for if the player is creating a new character
        frmMainMenu.picCreateChar.Visible = True
        frmSendGetData.Visible = False

        frmMainMenu.cmbClass.Clear

        For i = 0 To Max_Classes
            frmMainMenu.cmbClass.AddItem Trim(Class(i).Name)
        Next i
            
        frmMainMenu.cmbClass.ListIndex = 0
        frmMainMenu.lblHP.Caption = STR(Class(0).HP)
        frmMainMenu.lblMP.Caption = STR(Class(0).MP)
        frmMainMenu.lblSP.Caption = STR(Class(0).SP)
    
        frmMainMenu.lblSTR.Caption = STR(Class(0).STR)
        frmMainMenu.lblDEF.Caption = STR(Class(0).DEF)
        frmMainMenu.lblSPEED.Caption = STR(Class(0).SPEED)
        frmMainMenu.lblMAGI.Caption = STR(Class(0).MAGI)
        frmMainMenu.lblDEX.Caption = STR(Class(0).DEX)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "classesdata" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For i = 0 To Max_Classes
            Class(i).Name = Parse(n)
            
            Class(i).HP = Val(Parse(n + 1))
            Class(i).MP = Val(Parse(n + 2))
            Class(i).SP = Val(Parse(n + 3))
            
            Class(i).STR = Val(Parse(n + 4))
            Class(i).DEF = Val(Parse(n + 5))
            Class(i).SPEED = Val(Parse(n + 6))
            Class(i).MAGI = Val(Parse(n + 7))
            Class(i).DEX = Val(Parse(n + 8))
            
            n = n + 9
        Next i
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "ingame" Then
        InGame = True
        Call GameInit
        Call GameLoop
        If Parse(1) = END_CHAR Then
            MsgBox ("here")
            End
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinv" Then
        n = 1
        For i = 1 To MAX_INV
            Call SetPlayerInvItemNum(MyIndex, i, Val(Parse(n)))
            Call SetPlayerInvItemValue(MyIndex, i, Val(Parse(n + 1)))
            Call SetPlayerInvItemDur(MyIndex, i, Val(Parse(n + 2)))
            
            n = n + 3
        Next i
        Call UpdateInventory
        Call UpdateVisInventory
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinvupdate" Then
        n = Val(Parse(1))
        
        Call SetPlayerInvItemNum(MyIndex, n, Val(Parse(2)))
        Call SetPlayerInvItemValue(MyIndex, n, Val(Parse(3)))
        Call SetPlayerInvItemDur(MyIndex, n, Val(Parse(4)))
        Call UpdateInventory
        Call UpdateVisInventory
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerworneq" Then
        Call SetPlayerWeaponSlot(MyIndex, Val(Parse(1)))
        Call SetPlayerArmorSlot(MyIndex, Val(Parse(2)))
        Call SetPlayerHelmetSlot(MyIndex, Val(Parse(3)))
        Call SetPlayerShieldSlot(MyIndex, Val(Parse(4)))
        Call SetPlayerAmuletSlot(MyIndex, Val(Parse(5)))
        Call SetPlayerRingSlot(MyIndex, Val(Parse(6)))
        Call SetPlayerArrowSlot(MyIndex, Val(Parse(7)))
        Call UpdateInventory
        Call UpdateVisInventory
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playerhp" Then
        Player(MyIndex).MaxHp = Val(Parse(1))
        Call SetPlayerHP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then
            frmCClient.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
            'frmCClient.lblHP.Caption = Int(GetPlayerHP(MyIndex) / GetPlayerMaxHP(MyIndex) * 100) & "%"
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player mp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playermp" Then
        Player(MyIndex).MaxMP = Val(Parse(1))
        Call SetPlayerMP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxMP(MyIndex) > 0 Then
            frmCClient.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
            'frmCClient.lblMP.Caption = Int(GetPlayerMP(MyIndex) / GetPlayerMaxMP(MyIndex) * 100) & "%"
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playersp" Then
        Player(MyIndex).MaxSP = Val(Parse(1))
        Call SetPlayerSP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxSP(MyIndex) > 0 Then
            frmCClient.lblSP.Caption = GetPlayerSP(MyIndex) & " / " & GetPlayerMaxSP(MyIndex)
            'frmCClient.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Player stats packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerstats" Then
        Call SetPlayerSTR(MyIndex, Val(Parse(1)))
        Call SetPlayerDEF(MyIndex, Val(Parse(2)))
        Call SetPlayerSPEED(MyIndex, Val(Parse(3)))
        Call SetPlayerMAGI(MyIndex, Val(Parse(4)))
        Call SetPlayerDEX(MyIndex, Val(Parse(5)))
        Call SetPlayerLevel(MyIndex, Val(Parse(6)))
        Call SetPlayerExp(MyIndex, Val(Parse(7)))
        frmCClient.lblSTRENGTH.Caption = GetPlayerSTR(MyIndex)
        frmCClient.lblDEFENCE.Caption = GetPlayerDEF(MyIndex)
        frmCClient.lblSPEED.Caption = GetPlayerSPEED(MyIndex)
        frmCClient.lblMAGIC.Caption = GetPlayerMAGI(MyIndex)
        frmCClient.lblDEXTERITY.Caption = GetPlayerDEX(MyIndex)
        frmCClient.lblLEVEL.Caption = GetPlayerLevel(MyIndex)
        frmCClient.lblEXP.Caption = GetPlayerExp(MyIndex)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerdata" Then
        i = Val(Parse(1))
        
        Call SetPlayerName(i, Parse(2))
        Call SetPlayerSprite(i, Val(Parse(3)))
        Call SetPlayerMap(i, Val(Parse(4)))
        Call SetPlayerX(i, Val(Parse(5)))
        Call SetPlayerY(i, Val(Parse(6)))
        Call SetPlayerDir(i, Val(Parse(7)))
        Call SetPlayerAccess(i, Val(Parse(8)))
        Call SetPlayerPK(i, Val(Parse(9)))
        
        ' Make sure they aren't walking
        Player(i).Moving = 0
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        
        ' Check if the player is the client player, and if so reset directions
        If i = MyIndex Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
        
        'Remove shop window if open
        If frmCClient.picNpcQuests.Visible = True Then
            frmCClient.picNpcQuests.Visible = False
            frmCClient.lstNpcQuests.Clear
            frmCClient.cmdCompleteQuest.Visible = False
            frmCClient.cmdAbandonQuest.Visible = False
            QuestNpcNum = 0
            frmCClient.picScreen.SetFocus
        End If
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playermove") Then
        i = Val(Parse(1))
        x = Val(Parse(2))
        y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        Call SetPlayerX(i, x)
        Call SetPlayerY(i, y)
        Call SetPlayerDir(i, Dir)
                
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).Moving = n
        
        Select Case GetPlayerDir(i)
            Case DIR_UP
                Player(i).YOffset = PIC_Y
            Case DIR_DOWN
                Player(i).YOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(i).XOffset = PIC_X
            Case DIR_RIGHT
                Player(i).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcmove") Then
        i = Val(Parse(1))
        x = Val(Parse(2))
        y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        MapNpc(i).x = x
        MapNpc(i).y = y
        MapNpc(i).Dir = Dir
        MapNpc(i).XOffset = 0
        MapNpc(i).YOffset = 0
        MapNpc(i).Moving = n
        
        Select Case MapNpc(i).Dir
            Case DIR_UP
                MapNpc(i).YOffset = PIC_Y
            Case DIR_DOWN
                MapNpc(i).YOffset = PIC_Y * -1
            Case DIR_LEFT
                MapNpc(i).XOffset = PIC_X
            Case DIR_RIGHT
                MapNpc(i).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        Call SetPlayerDir(i, Dir)
        
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        MapNpc(i).Dir = Dir
        
        MapNpc(i).XOffset = 0
        MapNpc(i).YOffset = 0
        MapNpc(i).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "attack") Then
        i = Val(Parse(1))
        
        ' Set player to attacking
        Player(i).Attacking = 1
        Player(i).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Arrow check packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "checkarrows") Then
        n = Val(Parse(1))
        z = Val(Parse(2))
        i = Val(Parse(3))
       
        For x = 1 To MAX_PLAYER_ARROWS
             If Player(n).Arrow(x).Arrow = 0 Then
                 Player(n).Arrow(x).Arrow = 1
                 Player(n).Arrow(x).ArrowNum = z
                 Player(n).Arrow(x).ArrowAnim = z
                 Player(n).Arrow(x).ArrowTime = GetTickCount
                 Player(n).Arrow(x).ArrowVarX = 0
                 Player(n).Arrow(x).ArrowVarY = 0
                 Player(n).Arrow(x).ArrowY = GetPlayerY(n)
                 Player(n).Arrow(x).ArrowX = GetPlayerX(n)
                
                 If i = DIR_DOWN Then
                     Player(n).Arrow(x).ArrowY = GetPlayerY(n) + 1
                     Player(n).Arrow(x).ArrowPosition = 0
                     If Player(n).Arrow(x).ArrowY - 1 > MAX_MAPY Then
                           Player(n).Arrow(x).Arrow = 0
                           Exit Sub
                     End If
                 End If
                 If i = DIR_UP Then
                     Player(n).Arrow(x).ArrowY = GetPlayerY(n) - 1
                     Player(n).Arrow(x).ArrowPosition = 1
                     If Player(n).Arrow(x).ArrowY + 1 < 0 Then
                           Player(n).Arrow(x).Arrow = 0
                           Exit Sub
                     End If
                 End If
                 If i = DIR_RIGHT Then
                     Player(n).Arrow(x).ArrowX = GetPlayerX(n) + 1
                     Player(n).Arrow(x).ArrowPosition = 2
                     If Player(n).Arrow(x).ArrowX - 1 > MAX_MAPX Then
                           Player(n).Arrow(x).Arrow = 0
                           Exit Sub
                     End If
                 End If
                 If i = DIR_LEFT Then
                     Player(n).Arrow(x).ArrowX = GetPlayerX(n) - 1
                     Player(n).Arrow(x).ArrowPosition = 3
                     If Player(n).Arrow(x).ArrowX + 1 < 0 Then
                           Player(n).Arrow(x).Arrow = 0
                           Exit Sub
                     End If
                 End If
                 Exit For
             End If
        Next x
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcattack") Then
        i = Val(Parse(1))
        
        ' Set player to attacking
        MapNpc(i).Attacking = 1
        MapNpc(i).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "checkformap") Then
        ' Erase all players except self
        For i = 1 To HighIndex
            If i <> MyIndex Then
                Call SetPlayerMap(i, 0)
            End If
        Next i
        
        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        x = Val(Parse(1))
        
        ' Get revision
        y = Val(Parse(2))
        
        If FileExist(MAP_PATH & "map" & x & MAP_EXT, False) Then
            ' Check to see if the revisions match
            If GetMapRevision(x) = y Then
                ' We do so we dont need the map
                
                ' Load the map
                Call LoadMap(x)
                
                Call SendData("needmap" & SEP_CHAR & "no" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
        
        ' Either the revisions didn't match or we dont have the map, so we need it
        Call SendData("needmap" & SEP_CHAR & "yes" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "mapdata" Then
        n = 1
        
        SaveMap.Name = Parse(n + 1)
        SaveMap.Revision = Val(Parse(n + 2))
        SaveMap.Owner = Parse(n + 3)
        SaveMap.Moral = Val(Parse(n + 4))
        SaveMap.Up = Val(Parse(n + 5))
        SaveMap.Down = Val(Parse(n + 6))
        SaveMap.Left = Val(Parse(n + 7))
        SaveMap.Right = Val(Parse(n + 8))
        SaveMap.Music = Val(Parse(n + 9))
        SaveMap.BootMap = Val(Parse(n + 10))
        SaveMap.BootX = Val(Parse(n + 11))
        SaveMap.BootY = Val(Parse(n + 12))
        SaveMap.Indoors = Val(Parse(n + 13))
        
        n = n + 14
        
        For x = 1 To MAX_MAP_NPCS
            SaveMap.NSpawn(x).NSx = Val(Parse(n))
            SaveMap.NSpawn(x).NSy = Val(Parse(n + 1))
            
            n = n + 2
        Next x
        
        For x = 1 To MAX_MAP_RESOURCES
            SaveMap.RSpawn(x).RSx = Val(Parse(n))
            SaveMap.RSpawn(x).RSy = Val(Parse(n + 1))
            
            n = n + 2
        Next x
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                SaveMap.Tile(x, y).Tileset = Val(Parse(n))
                SaveMap.Tile(x, y).Ground = Val(Parse(n + 1))
                SaveMap.Tile(x, y).Mask = Val(Parse(n + 2))
                SaveMap.Tile(x, y).Mask2 = Val(Parse(n + 3))
                SaveMap.Tile(x, y).Anim = Val(Parse(n + 4))
                SaveMap.Tile(x, y).Fringe = Val(Parse(n + 5))
                SaveMap.Tile(x, y).Fringe2 = Val(Parse(n + 6))
                SaveMap.Tile(x, y).FAnim = Val(Parse(n + 7))
                SaveMap.Tile(x, y).Light = Val(Parse(n + 8))
                SaveMap.Tile(x, y).Type = Val(Parse(n + 9))
                SaveMap.Tile(x, y).Data1 = Val(Parse(n + 10))
                SaveMap.Tile(x, y).Data2 = Val(Parse(n + 11))
                SaveMap.Tile(x, y).Data3 = Val(Parse(n + 12))
                SaveMap.Tile(x, y).WalkUp = Parse(n + 13)
                SaveMap.Tile(x, y).WalkDown = Parse(n + 14)
                SaveMap.Tile(x, y).WalkLeft = Parse(n + 15)
                SaveMap.Tile(x, y).WalkRight = Parse(n + 16)
                SaveMap.Tile(x, y).Build = Parse(n + 17)
                
                n = n + 18
            Next x
        Next y
        
        For x = 1 To MAX_MAP_NPCS
            SaveMap.Npc(x) = Val(Parse(n))
            n = n + 1
        Next x
        
        For x = 1 To MAX_MAP_RESOURCES
            SaveMap.Resource(x) = Val(Parse(n))
            n = n + 1
        Next x
                
        ' Save the map
        Call SaveLocalMap(Val(Parse(1)))
        
        '' Check if we get a map from someone else and if we were editing a map cancel it out
        'If InEditor Then
            'InEditor = False
            'frmCClient.picMapEditor.Visible = False
            
            'If frmMapWarp.Visible Then
                'Unload frmMapWarp
            'End If
            
            'If frmMapProperties.Visible Then
                'Unload frmMapProperties
            'End If
        'End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapitemdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).Num = Val(Parse(n))
            SaveMapItem(i).Value = Val(Parse(n + 1))
            SaveMapItem(i).Dur = Val(Parse(n + 2))
            SaveMapItem(i).x = Val(Parse(n + 3))
            SaveMapItem(i).y = Val(Parse(n + 4))
            
            n = n + 5
        Next i
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapnpcdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_NPCS
            SaveMapNpc(i).Num = Val(Parse(n))
            SaveMapNpc(i).x = Val(Parse(n + 1))
            SaveMapNpc(i).y = Val(Parse(n + 2))
            SaveMapNpc(i).Dir = Val(Parse(n + 3))
            
            n = n + 4
        Next i
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Map resource data packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapresourcedata" Then
        n = 1
        
        For i = 1 To MAX_MAP_RESOURCES
            SaveMapResource(i).Num = Val(Parse(n))
            SaveMapResource(i).x = Val(Parse(n + 1))
            SaveMapResource(i).y = Val(Parse(n + 2))
            
            n = n + 3
        Next i
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapdone" Then
        Map = SaveMap
        
        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i
        
        For i = 1 To MAX_MAP_RESOURCES
            MapResource(i) = SaveMapResource(i)
        Next i
        
        GettingMap = False
        
        '' Play music
        'Call StopMidi
        'If Map.Music > 0 Then
            'Call PlayMidi("music" & Trim(STR(Map.Music)) & ".mid")
        'End If
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "saymsg") Or (LCase(Parse(0)) = "broadcastmsg") Or (LCase(Parse(0)) = "globalmsg") Or (LCase(Parse(0)) = "playermsg") Or (LCase(Parse(0)) = "mapmsg") Or (LCase(Parse(0)) = "adminmsg") Then
        Call AddText(Parse(1), Val(Parse(2)))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: PushBlock packet ::
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "pushblock") Then
        x = Val(Parse(1)) 'x co-ordinate
        y = Val(Parse(2)) 'y co-ordinate
        n = Val(Parse(3)) 'Pushed Value
        i = Val(Parse(4)) 'Player Direction
        f = Val(Parse(5)) 'Movement
        
        PushTile(x, y).Pushed = n
        PushTile(x, y).Moving = f
        If PushTile(x, y).Pushed = NO Then
            Select Case PushTile(x, y).Dir
                Case DIR_UP
                    Map.Tile(x, y - 1).Mask = 0
                    Map.Tile(x, y - 1).Type = TILE_TYPE_NPCAVOID
                    PushTile(x, y).XOffset = 0
                    PushTile(x, y).YOffset = -32
                    
                Case DIR_DOWN
                    Map.Tile(x, y + 1).Mask = 0
                    Map.Tile(x, y + 1).Type = TILE_TYPE_NPCAVOID
                    PushTile(x, y).XOffset = 0
                    PushTile(x, y).YOffset = 32
                    
                Case DIR_LEFT
                    Map.Tile(x - 1, y).Mask = 0
                    Map.Tile(x - 1, y).Type = TILE_TYPE_NPCAVOID
                    PushTile(x, y).XOffset = -32
                    PushTile(x, y).YOffset = 0
                    
                Case DIR_RIGHT
                    Map.Tile(x + 1, y).Mask = 0
                    Map.Tile(x + 1, y).Type = TILE_TYPE_NPCAVOID
                    PushTile(x, y).XOffset = 32
                    PushTile(x, y).YOffset = 0
            End Select
            Exit Sub
        End If
        PushTile(x, y).Dir = i
        If PushTile(x, y).Pushed = YES Then
            Select Case PushTile(x, y).Dir
                Case DIR_UP
                    Map.Tile(x, y - 1).Type = TILE_TYPE_BLOCKED
                    
                Case DIR_DOWN
                    Map.Tile(x, y + 1).Type = TILE_TYPE_BLOCKED
                    
                Case DIR_LEFT
                    Map.Tile(x - 1, y).Type = TILE_TYPE_BLOCKED
                    
                Case DIR_RIGHT
                    Map.Tile(x + 1, y).Type = TILE_TYPE_BLOCKED
            End Select
        End If
        PushTile(x, y).XOffset = 0
        PushTile(x, y).YOffset = 0
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnitem" Then
        n = Val(Parse(1))
        
        MapItem(n).Num = Val(Parse(2))
        MapItem(n).Value = Val(Parse(3))
        MapItem(n).Dur = Val(Parse(4))
        MapItem(n).x = Val(Parse(5))
        MapItem(n).y = Val(Parse(6))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateitem") Then
        n = Val(Parse(1))
        
        ' Update the item
        Item(n).Name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc resource packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnresource" Then
        n = Val(Parse(1))
        
        MapResource(n).Num = Val(Parse(2))
        MapResource(n).x = Val(Parse(3))
        MapResource(n).y = Val(Parse(4))
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnnpc" Then
        n = Val(Parse(1))
        
        MapNpc(n).Num = Val(Parse(2))
        MapNpc(n).x = Val(Parse(3))
        MapNpc(n).y = Val(Parse(4))
        MapNpc(n).Dir = Val(Parse(5))
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "npcdead" Then
        n = Val(Parse(1))
        
        MapNpc(n).Num = 0
        MapNpc(n).x = 0
        MapNpc(n).y = 0
        MapNpc(n).Dir = 0
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Resource dead packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "resourcedead" Then
        n = Val(Parse(1))
        
        MapResource(n).Num = 0
        MapResource(n).x = 0
        MapResource(n).y = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatenpc") Then
        n = Val(Parse(1))
        
        ' Update the item
        Npc(n).Name = Parse(2)
        Npc(n).Sprite = Val(Parse(3))
        Npc(n).SpawnSecs = 0
        Npc(n).Behavior = Val(Parse(4))
        Npc(n).Range = 0
        Npc(n).STR = 0
        Npc(n).DEF = 0
        Npc(n).SPEED = 0
        Npc(n).MAGI = 0
        Npc(n).Big = Val(Parse(5))
        Npc(n).MaxHp = 0
        Npc(n).Respawn = 0
        Npc(n).HitOnlyWith = 0
        Npc(n).ShopLink = 0
        Npc(n).EXP = 0
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = 0
            Npc(n).ItemNPC(i).ItemNum = 0
            Npc(n).ItemNPC(i).ItemValue = 0
        Next i
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "mapkey") Then
        x = Val(Parse(1))
        y = Val(Parse(2))
        n = Val(Parse(3))
        
        TempTile(x, y).DoorOpen = n
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateshop") Then
        n = Val(Parse(1))
        
        ' Update the shop name
        Shop(n).Name = Parse(2)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatespell") Then
        n = Val(Parse(1))
        
        ' Update the spell
        Spell(n).Name = Parse(2)
        Spell(n).SpellSprite = Val(Parse(3))
        Spell(n).ClassReq = Val(Parse(4))
        Spell(n).LevelReq = Val(Parse(5))
        Spell(n).Type = Val(Parse(6))
        Spell(n).Data1 = Val(Parse(7))
        Spell(n).Data2 = Val(Parse(8))
        Spell(n).Data3 = Val(Parse(9))
        'Call UpdateVisSpells
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update skill packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateskill") Then
        n = Val(Parse(1))
        
        ' Update the spell name
        Skill(n).Name = Parse(2)
        Skill(n).SkillSprite = Val(Parse(3))
        Skill(n).ClassReq = Val(Parse(4))
        Skill(n).LevelReq = Val(Parse(5))
        Skill(n).Type = Val(Parse(6))
        Skill(n).Data1 = Val(Parse(7))
        Skill(n).Data2 = Val(Parse(8))
        Skill(n).Data3 = Val(Parse(9))
        'Call UpdateVisSkills
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (LCase(Parse(0)) = "trade") Then
        ShopNum = Val(Parse(1))
        Player(MyIndex).ShopNum = Val(Parse(1))
        If Val(Parse(2)) = 1 Then
            frmCClient.cmdShopFixItems.Enabled = True
        Else
            frmCClient.cmdShopFixItems.Enabled = False
        End If
        
        n = 3
        frmCClient.lstShopTrade.Clear
        For i = 1 To MAX_TRADES
            For f = 1 To MAX_GIVE_ITEMS
                Shop(ShopNum).TradeItem(i).GiveItem(f) = Val(Parse((n - 1) + f))
            Next f
            n = n + MAX_GIVE_ITEMS
            For f = 1 To MAX_GIVE_VALUE
                Shop(ShopNum).TradeItem(i).GiveValue(f) = Val(Parse((n - 1) + f))
            Next f
            n = n + MAX_GIVE_VALUE
            For f = 1 To MAX_GET_ITEMS
                Shop(ShopNum).TradeItem(i).GetItem(f) = Val(Parse((n - 1) + f))
            Next f
            n = n + MAX_GET_ITEMS
            For f = 1 To MAX_GET_VALUE
                Shop(ShopNum).TradeItem(i).GetValue(f) = Val(Parse((n - 1) + f))
            Next f
            n = n + MAX_GET_VALUE
            
            Shop(ShopNum).ItemStock(i) = Val(Parse(n))
            n = n + 1
            
            TradeDescription = ""
        
            For f = 1 To MAX_GET_ITEMS
                If (Shop(ShopNum).TradeItem(i).GetItem(f) > 0) And (Shop(ShopNum).TradeItem(i).GetValue(f) > 0) Then
                    TradeDescription = TradeDescription & Shop(ShopNum).TradeItem(i).GetValue(f) & " x " & Trim(Item(Shop(ShopNum).TradeItem(i).GetItem(f)).Name) & " + "
                End If
            Next f
            If TradeDescription = "" Then
                frmCClient.lstShopTrade.AddItem "Empty Trade Slot"
                frmCClient.lstShopTrade.ListIndex = 0
                If frmCClient.lstShopTrade.ListCount > 0 Then
                    frmCClient.lstShopTrade.ListIndex = 0
                End If
                frmCClient.lblShopName = Trim(Shop(ShopNum).Name)
                frmCClient.picShopTrade.Visible = True
                Exit Sub
            End If
            TradeDescription = Left(TradeDescription, (Len(TradeDescription) - 2))
            TradeDescription = TradeDescription & "for "
            For f = 1 To MAX_GIVE_ITEMS
                If (Shop(ShopNum).TradeItem(i).GiveItem(f) > 0) And (Shop(ShopNum).TradeItem(i).GiveValue(f) > 0) Then
                    TradeDescription = TradeDescription & Shop(ShopNum).TradeItem(i).GiveValue(f) & " x " & Trim(Item(Shop(ShopNum).TradeItem(i).GiveItem(f)).Name) & " + "
                End If
            Next f
            If TradeDescription = "" Then
                frmCClient.lstShopTrade.AddItem i & "Empty Trade Slot"
                frmCClient.lstShopTrade.ListIndex = 0
                If frmCClient.lstShopTrade.ListCount > 0 Then
                    frmCClient.lstShopTrade.ListIndex = 0
                End If
                frmCClient.lblShopName = Trim(Shop(ShopNum).Name)
                frmCClient.picShopTrade.Visible = True
                Exit Sub
            End If
            TradeDescription = Left(TradeDescription, (Len(TradeDescription) - 2))
            frmCClient.lstShopTrade.AddItem i & ": " & TradeDescription
            frmCClient.lstShopTrade.ListIndex = 0
        Next i
        
        If frmCClient.lstShopTrade.ListCount > 0 Then
            frmCClient.lstShopTrade.ListIndex = 0
        End If
        frmCClient.lblShopName = Trim(Shop(ShopNum).Name)
        frmCClient.picShopTrade.Visible = True
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (LCase(Parse(0)) = "spells") Then
        
        'frmCClient.picPlayerSpells.Visible = True
        frmCClient.lstPlayerSpells.Clear
        
        n = 1
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spells(i).Num = Val(Parse(n))
            Player(MyIndex).Spells(i).Level = Val(Parse(n + 1))
            Player(MyIndex).Spells(i).EXP = Val(Parse(n + 2))
            If Player(MyIndex).Spells(i).Num <> 0 Then
                frmCClient.lstPlayerSpells.AddItem i & ": " & Trim(Spell(Player(MyIndex).Spells(i).Num).Name)
            Else
                frmCClient.lstPlayerSpells.AddItem "<free spells slot>"
            End If
            
            n = n + 3
        Next i
        
        frmCClient.lstPlayerSpells.ListIndex = 0
        Call UpdateVisSpells
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Player Spells Level Update Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerspellslvl" Then
        i = Val(Parse(1))
        
        Player(MyIndex).Spells(i).Level = Val(Parse(2))
        Call UpdateInventory
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::
    ' :: Player Spells Exp Update Packet ::
    ' :::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerspellsexp" Then
        i = Val(Parse(1))
        
        Player(MyIndex).Spells(i).EXP = Val(Parse(2))
        Call UpdateInventory
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Skills packet ::
    ' :::::::::::::::::::
    If (LCase(Parse(0)) = "skills") Then
        
        'frmCClient.picPlayerSkills.Visible = True
        frmCClient.lstPlayerSkills.Clear
        
        n = 1
        
        ' Put skills known in player record
        For i = 1 To MAX_PLAYER_SKILLS
            Player(MyIndex).Skills(i).Num = Val(Parse(n))
            Player(MyIndex).Skills(i).Level = Val(Parse(n + 1))
            Player(MyIndex).Skills(i).EXP = Val(Parse(n + 2))
            If Player(MyIndex).Skills(i).Num <> 0 Then
                frmCClient.lstPlayerSkills.AddItem i & ": " & Trim(Skill(Player(MyIndex).Skills(i).Num).Name)
            Else
                frmCClient.lstPlayerSkills.AddItem "<free skills slot>"
            End If
            
            n = n + 3
        Next i
        
        frmCClient.lstPlayerSkills.ListIndex = 0
        Call UpdateVisSkills
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Player Skills Level Update Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerskillslvl" Then
        i = Val(Parse(1))
        
        Player(MyIndex).Skills(i).Level = Val(Parse(2))
        Call UpdateInventory
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::
    ' :: Player Skills Exp Update Packet ::
    ' :::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerskillsexp" Then
        i = Val(Parse(1))
        
        Player(MyIndex).Skills(i).EXP = Val(Parse(2))
        Call UpdateInventory
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Player quests ::
    ' :::::::::::::::::::
    If (LCase(Parse(0)) = "quests") Then
        'Add current quests to player rec
        n = 1
        
        For i = 1 To MAX_PLAYER_QUESTS
            Player(MyIndex).Quests(i).Num = Val(Parse(n))
            Player(MyIndex).Quests(i).SetMap = Val(Parse(n + 1))
            Player(MyIndex).Quests(i).SetBy = Val(Parse(n + 2))
            Player(MyIndex).Quests(i).Amount = Val(Parse(n + 3))
            Player(MyIndex).Quests(i).Count = Val(Parse(n + 4))
            
            n = n + 5
        Next i
        
        frmCClient.scrlPlayerQuestNum.Value = 1
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Player quest update ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerquest") Then
        n = Val(Parse(1))
        
        Player(MyIndex).Quests(n).Num = Val(Parse(2))
        Player(MyIndex).Quests(n).SetMap = Val(Parse(3))
        Player(MyIndex).Quests(n).SetBy = Val(Parse(4))
        Player(MyIndex).Quests(n).Amount = Val(Parse(5))
        Player(MyIndex).Quests(n).Count = Val(Parse(6))
        
        If Player(MyIndex).Quests(n).Num > 0 Then
            If (frmCClient.picPlayerQuests.Visible = True) Then 'And (frmCClient.scrlPlayerQuestNum.Value = n) Then
                frmCClient.scrlPlayerQuestNum.Value = n
                frmCClient.lblPlayerQuestName.Caption = Trim(Quest(Player(MyIndex).Quests(n).Num).Name)
                frmCClient.txtPlayerQuestDesc.Text = Trim(Quest(Player(MyIndex).Quests(n).Num).Description)
                If Quest(Player(MyIndex).Quests(n).Num).RewardValue > 1 Then
                    frmCClient.lblPlayerQuestReward.Caption = STR(Quest(Player(MyIndex).Quests(n).Num).RewardValue) & "  x  " & Trim(Item(Quest(Player(MyIndex).Quests(n).Num).Reward).Name)
                Else
                    frmCClient.lblPlayerQuestReward.Caption = Trim(Item(Quest(Player(MyIndex).Quests(n).Num).Reward).Name)
                End If
                frmCClient.lblPlayerQuestCount.Caption = STR(Player(MyIndex).Quests(n).Count) & " / " & STR(Player(MyIndex).Quests(n).Amount)
                If Player(MyIndex).Quests(n).Count = Player(MyIndex).Quests(n).Amount Then
                    frmCClient.cmdAbandonQuest.Visible = False
                    frmCClient.cmdCompleteQuest.Visible = True
                Else
                    frmCClient.cmdAbandonQuest.Visible = True
                End If
            End If
        Else
            frmCClient.scrlPlayerQuestNum.Value = n
            frmCClient.lblPlayerQuestName.Caption = "No Quest"
            frmCClient.txtPlayerQuestDesc.Text = ""
            frmCClient.lblPlayerQuestReward.Caption = ""
            frmCClient.lblPlayerQuestCount.Caption = ""
            frmCClient.cmdAbandonQuest.Visible = False
            frmCClient.cmdCompleteQuest.Visible = False
        End If
        'frmCClient.scrlPlayerQuestNum.Value = n
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Update quest packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatequest") Then
        n = Val(Parse(1))
        
        ' Update the quest
        Quest(n).Name = Parse(2)
        'Quest(n).SetBy = Val(Parse(3))
        Quest(n).ClassReq = Val(Parse(3))
        Quest(n).LevelMin = Val(Parse(4))
        Quest(n).LevelMax = Val(Parse(5))
        'Quest(n).Type = Val(Parse(7))
        Quest(n).Reward = Val(Parse(6))
        Quest(n).RewardValue = Val(Parse(7))
        'Quest(n).Data1 = Val(Parse(10))
        'Quest(n).Data2 = Val(Parse(11))
        'Quest(n).Data3 = Val(Parse(12))
        Quest(n).Description = Parse(8)
        Exit Sub
    End If
    
    ' ::::::::::::::::
    ' :: Npc Quests ::
    ' ::::::::::::::::
    If LCase(Parse(0)) = "npcquests" Then
        n = Val(Parse(1))
        QuestNpcNum = Val(Parse(1))
        
        i = 2
        
        frmCClient.lstNpcQuests.Clear
        For f = 1 To MAX_NPC_QUESTS
            Npc(n).QuestNPC(f) = Val(Parse(i))
            If Npc(n).QuestNPC(f) > 0 Then
                frmCClient.lstNpcQuests.AddItem f & ": " & Trim(Quest(Npc(n).QuestNPC(f)).Name)
            Else
                frmCClient.lstNpcQuests.AddItem ""
            End If
            
            i = i + 1
        Next f
        
        frmCClient.lblQuestNpc.Caption = Trim(Npc(n).Name) & "'s Quests"
        
        frmCClient.lstNpcQuests.ListIndex = 0
        frmCClient.picNpcQuests.Visible = True
        frmCClient.scrlPlayerQuestNum.Value = 1
        frmCClient.picPlayerQuests.Visible = True
        frmCClient.chkPlayerQuestPin.Value = Checked
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: High Index Packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "highindex" Then
        HighIndex = Val(Parse(1))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "blitplayerdmg" Then
        DmgDamage = Val(Parse(1))
        NPCWho = Val(Parse(2))
        DmgColor = Val(Parse(3))
        DmgTime = GetTickCount
        iii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "blitnpcdmg" Then
        NPCDmgDamage = Val(Parse(1))
        NPCDmgColor = Val(Parse(2))
        NPCDmgTime = GetTickCount
        ii = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' : Blit PK Damage :
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "blitpkdmg" Then
        PKDmgDamage = Val(Parse(1))
        PKWho = Val(Parse(2))
        PKDmgColor = Val(Parse(3))
        PKDmgTime = GetTickCount
        viii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' : Blit PK Message :
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "blitpkmsg" Then
        PKMsgMessage = Parse(1)
        PKMsgWho = Val(Parse(2))
        PKMsgColor = Val(Parse(3))
        PKMsgTime = GetTickCount
        viv = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' : Blit Player Message :
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "blitplayermsg" Then
        MsgMessage = Parse(1)
        MessageColor = Val(Parse(2))
        MessageTime = GetTickCount
        iv = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' : Blit Npc Message :
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "blitnpcmsg" Then
        NpcMsgMessage = Parse(1)
        NpcMsgWho = Val(Parse(2))
        NpcMessageColor = Val(Parse(3))
        NpcMessageTime = GetTickCount
        vi = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' : Blit Player Warn Message :
    ' ::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "blitwarnmsg" Then
        WarnMessage = Parse(1)
        WarnMsgColor = Val(Parse(2))
        WarnMsgTime = GetTickCount
        vii = 15
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' : Blit Resource Message :
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "blitresourcemsg" Then
        ResourceWho = Val(Parse(1))
        ResourceMsgMessage = Parse(2)
        ResourceMsgColor = Val(Parse(3))
        ResourceMsgTime = GetTickCount
        xi = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' : Blit Resource Damage :
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "blitresourcedmg" Then
        ResourceDmgWho = Val(Parse(1))
        ResourceDmgDamage = Parse(2)
        ResourceDmgColor = Val(Parse(3))
        ResourceDmgTime = GetTickCount
        xii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' : Blit Item Message :
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "blititemmsg" Then
        ItemWho = Val(Parse(2))
        ItemMsgMessage = Trim(Parse(1))
        ItemMsgColor = Val(Parse(3))
        ItemMsgTime = GetTickCount
        xiv = 0
        Exit Sub
    End If

End Sub
