Attribute VB_Name = "modHandleData"
Option Explicit


Sub HandleData(ByVal Data As String)
Dim Parse() As String
Dim dData() As String
Dim Name As String
Dim Password As String
Dim Sex As Long
Dim ClassNum As Long
Dim CharNum As Long
Dim Msg As String
Dim IPMask As String
Dim BanSlot As Long
Dim MsgTo As Long
Dim Dir As Long
Dim InvNum As Long
Dim Ammount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Level As Long
Dim I As Long, N As Long, X As Long, Y As Long, Q As Long
Dim ShopNum As Long, GiveItem As Long, GiveValue As Long, GetItem As Long, GetValue As Long

'Debug.Print "R: " & Data

    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
    
    ' Add the data to the debug window if we are in debug mode
    If Trim(Command) = "-debug" Then
        If frmDebug.Visible = False Then frmDebug.Visible = True
        Call TextAdd(frmDebug.txtDebug, "((( Processed Packet " & Parse(0) & " )))", True)
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
        N = 1
        
        frmChars.Visible = True
        frmSendGetData.Visible = False
        
        frmChars.lstChars.Clear
        
        For I = 1 To MAX_CHARS
            Name = Parse(N)
            Msg = Parse(N + 1)
            Level = Val(Parse(N + 2))
            
            If Trim(Name) = "" Then
                frmChars.lstChars.AddItem "Free Character Slot"
            Else
                frmChars.lstChars.AddItem Name & " a level " & Level & " " & Msg
            End If
            
            N = N + 3
        Next I
        
        frmChars.lstChars.ListIndex = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "loginok" Then
        ' Now we can receive game data
        MyIndex = Val(Parse(1))
        'Grab key for future use.
        MyKEY = Parse(2)
        
        frmSendGetData.Visible = True
        frmChars.Visible = False
        
        Call SetStatus("Receiving game data...")
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "newcharclasses" Then
        N = 1
        
        ' Max classes
        Max_Classes = Val(Parse(N))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        N = N + 1
        
        For I = 0 To Max_Classes
            Class(I).Name = Parse(N)
            
            Class(I).HP = Val(Parse(N + 1))
            Class(I).MP = Val(Parse(N + 2))
            Class(I).SP = Val(Parse(N + 3))
            
            Class(I).str = Val(Parse(N + 4))
            Class(I).DEF = Val(Parse(N + 5))
            Class(I).SPEED = Val(Parse(N + 6))
            Class(I).MAGI = Val(Parse(N + 7))
            
            N = N + 8
        Next I
        
        ' Used for if the player is creating a new character
        frmNewChar.Visible = True
        frmSendGetData.Visible = False

        frmNewChar.cmbClass.Clear

        For I = 0 To Max_Classes
            frmNewChar.cmbClass.AddItem Trim(Class(I).Name)
        Next I
            
        frmNewChar.cmbClass.ListIndex = 0
        frmNewChar.lblHP.Caption = str(Class(0).HP)
        frmNewChar.lblMP.Caption = str(Class(0).MP)
        frmNewChar.lblSP.Caption = str(Class(0).SP)
    
        frmNewChar.lblSTR.Caption = str(Class(0).str)
        frmNewChar.lblDEF.Caption = str(Class(0).DEF)
        frmNewChar.lblSPEED.Caption = str(Class(0).SPEED)
        frmNewChar.lblMAGI.Caption = str(Class(0).MAGI)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "classesdata" Then
        N = 1
        
        ' Max classes
        Max_Classes = Val(Parse(N))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        N = N + 1
        
        For I = 0 To Max_Classes
            Class(I).Name = Parse(N)
            
            Class(I).HP = Val(Parse(N + 1))
            Class(I).MP = Val(Parse(N + 2))
            Class(I).SP = Val(Parse(N + 3))
            
            Class(I).str = Val(Parse(N + 4))
            Class(I).DEF = Val(Parse(N + 5))
            Class(I).SPEED = Val(Parse(N + 6))
            Class(I).MAGI = Val(Parse(N + 7))
            
            N = N + 8
        Next I
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
            'MsgBox ("here")
            End
        End If
        Exit Sub
    End If
    
    'CACHES!!!
    
    'ItemCache
    If LCase(Parse(0)) = "itemcache" Then
        N = 2
        For I = 1 To MAX_ITEMS
            'Debug.Print "I: " & I
            With Item(I)
                .Name = Trim(CStr(Parse(N)))
                .Pic = CLng(Parse(N + 1))
                .Type = CLng(Parse(N + 2))
            End With
            N = N + 3
        Next I
        Exit Sub
    End If
    
    'NPCCache
    If LCase(Parse(0)) = "npccache" Then
        N = 2
        For I = 1 To MAX_NPCS
            With Npc(I)
                .Name = Trim(CStr(Parse(N)))
                .Sprite = CLng(Parse(N + 1))
            End With
            N = N + 2
        Next I
        Exit Sub
    End If
    
    
    
    ' Player can see the assignment box :) Only Owners for now.
    If LCase(Parse(0)) = "assignok" Then
        'Syntax: ASSIGNOK <Count> <Name1> <Name2> <Name3> <Name4>
        'Will only return staff members for assignment.
        frmAssigned.Show ' vbModal
        With frmAssigned.cboCharacter
            N = CLng(Parse(1))
            Q = 2
            For I = 1 To N
                .AddItem Trim(CStr(Parse(Q)))
                Q = Q + 1
            Next I
            .ListIndex = 0
        End With
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinv" Then
        N = 1
        For I = 1 To MAX_INV
            Call SetPlayerInvItemNum(MyIndex, I, Val(Parse(N)))
            Call SetPlayerInvItemValue(MyIndex, I, Val(Parse(N + 1)))
            Call SetPlayerInvItemDur(MyIndex, I, Val(Parse(N + 2)))
            
            N = N + 3
        Next I
        Call UpdateInventory
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinvupdate" Then
        N = Val(Parse(1))
        
        Call SetPlayerInvItemNum(MyIndex, N, Val(Parse(2)))
        Call SetPlayerInvItemValue(MyIndex, N, Val(Parse(3)))
        Call SetPlayerInvItemDur(MyIndex, N, Val(Parse(4)))
        Call UpdateInventory
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerworneq" Then
        Call SetPlayerArmorSlot(MyIndex, Val(Parse(1)))
        Call SetPlayerWeaponSlot(MyIndex, Val(Parse(2)))
        Call SetPlayerHelmetSlot(MyIndex, Val(Parse(3)))
        Call SetPlayerShieldSlot(MyIndex, Val(Parse(4)))
        Call UpdateInventory
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playerhp" Then
        Player(MyIndex).MaxHP = Val(Parse(1))
        Call SetPlayerHP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then
            frmMirage.lblHP.Caption = Int(GetPlayerHP(MyIndex) / GetPlayerMaxHP(MyIndex) * 100) & "%"
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
            frmMirage.lblMP.Caption = Int(GetPlayerMP(MyIndex) / GetPlayerMaxMP(MyIndex) * 100) & "%"
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
            frmMirage.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
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
        Exit Sub
    End If
                
    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerdata" Then
        I = Val(Parse(1))
        
        Call SetPlayerName(I, Parse(2))
        Call SetPlayerSprite(I, Val(Parse(3)))
        Call SetPlayerMap(I, Val(Parse(4)))
        Call SetPlayerX(I, Val(Parse(5)))
        Call SetPlayerY(I, Val(Parse(6)))
        Call SetPlayerDir(I, Val(Parse(7)))
        Call SetPlayerAccess(I, Val(Parse(8)))
        Call SetPlayerPK(I, Val(Parse(9)))
        
        ' Make sure they aren't walking
        Player(I).Moving = 0
        Player(I).XOffset = 0
        Player(I).YOffset = 0
        
        ' Check if the player is the client player, and if so reset directions
        If I = MyIndex Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playermove") Then
        I = Val(Parse(1))
        X = Val(Parse(2))
        Y = Val(Parse(3))
        Dir = Val(Parse(4))
        N = Val(Parse(5))

        Call SetPlayerX(I, X)
        Call SetPlayerY(I, Y)
        Call SetPlayerDir(I, Dir)
                
        Player(I).XOffset = 0
        Player(I).YOffset = 0
        Player(I).Moving = N
        
        Select Case GetPlayerDir(I)
            Case DIR_UP
                Player(I).YOffset = PIC_Y
            Case DIR_DOWN
                Player(I).YOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(I).XOffset = PIC_X
            Case DIR_RIGHT
                Player(I).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcmove") Then
        I = Val(Parse(1))
        X = Val(Parse(2))
        Y = Val(Parse(3))
        Dir = Val(Parse(4))
        N = Val(Parse(5))

        MapNpc(I).X = X
        MapNpc(I).Y = Y
        MapNpc(I).Dir = Dir
        MapNpc(I).XOffset = 0
        MapNpc(I).YOffset = 0
        MapNpc(I).Moving = N
        
        Select Case MapNpc(I).Dir
            Case DIR_UP
                MapNpc(I).YOffset = PIC_Y
            Case DIR_DOWN
                MapNpc(I).YOffset = PIC_Y * -1
            Case DIR_LEFT
                MapNpc(I).XOffset = PIC_X
            Case DIR_RIGHT
                MapNpc(I).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerdir") Then
        I = Val(Parse(1))
        Dir = Val(Parse(2))
        Call SetPlayerDir(I, Dir)
        
        Player(I).XOffset = 0
        Player(I).YOffset = 0
        Player(I).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcdir") Then
        I = Val(Parse(1))
        Dir = Val(Parse(2))
        MapNpc(I).Dir = Dir
        
        MapNpc(I).XOffset = 0
        MapNpc(I).YOffset = 0
        MapNpc(I).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerxy") Then
        X = Val(Parse(1))
        Y = Val(Parse(2))
        
        Call SetPlayerX(MyIndex, X)
        Call SetPlayerY(MyIndex, Y)
        
        ' Make sure they aren't walking
        Player(MyIndex).Moving = 0
        Player(MyIndex).XOffset = 0
        Player(MyIndex).YOffset = 0
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "attack") Then
        I = Val(Parse(1))
        
        ' Set player to attacking
        Player(I).Attacking = 1
        Player(I).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcattack") Then
        I = Val(Parse(1))
        
        ' Set player to attacking
        MapNpc(I).Attacking = 1
        MapNpc(I).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "checkformap") Then
        ' Erase all players except self
        For I = 1 To MAX_PLAYERS
            If I <> MyIndex Then
                Call SetPlayerMap(I, 0)
            End If
        Next I
        
        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        X = Val(Parse(1))
        
        ' Get revision
        Y = Val(Parse(2))
        
        If FileExist("maps\map" & X & ".dat") Then
            ' Check to see if the revisions match
            If GetMapRevision(X) = Y Then
                ' We do so we dont need the map
                
                ' Load the map
                Call LoadMap(X)
                
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
        N = 1
        
        SaveMap.Name = Parse(N + 1)
        SaveMap.Revision = Val(Parse(N + 2))
        SaveMap.Moral = Val(Parse(N + 3))
        SaveMap.Up = Val(Parse(N + 4))
        SaveMap.Down = Val(Parse(N + 5))
        SaveMap.Left = Val(Parse(N + 6))
        SaveMap.Right = Val(Parse(N + 7))
        SaveMap.Music = Val(Parse(N + 8))
        SaveMap.BootMap = Val(Parse(N + 9))
        SaveMap.BootX = Val(Parse(N + 10))
        SaveMap.BootY = Val(Parse(N + 11))
        SaveMap.Shop = Val(Parse(N + 12))
        
        N = N + 13
        
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                SaveMap.Tile(X, Y).Ground = Val(Parse(N))
                SaveMap.Tile(X, Y).Mask = Val(Parse(N + 1))
                SaveMap.Tile(X, Y).Anim = Val(Parse(N + 2))
                SaveMap.Tile(X, Y).Fringe = Val(Parse(N + 3))
                SaveMap.Tile(X, Y).Type = Val(Parse(N + 4))
                SaveMap.Tile(X, Y).Data1 = Val(Parse(N + 5))
                SaveMap.Tile(X, Y).Data2 = Val(Parse(N + 6))
                SaveMap.Tile(X, Y).Data3 = Val(Parse(N + 7))
                
                N = N + 8
            Next X
        Next Y
        
        For X = 1 To MAX_MAP_NPCS
            SaveMap.Npc(X) = Val(Parse(N))
            N = N + 1
        Next X
                
        ' Save the map
        Call SaveLocalMap(Val(Parse(1)))
        
        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmMirage.picMapEditor.Visible = False
            
            If frmMapWarp.Visible Then
                Unload frmMapWarp
            End If
            
            If frmMapProperties.Visible Then
                Unload frmMapProperties
            End If
        End If
        
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapitemdata" Then
        N = 1
        
        For I = 1 To MAX_MAP_ITEMS
            SaveMapItem(I).Num = Val(Parse(N))
            SaveMapItem(I).Value = Val(Parse(N + 1))
            SaveMapItem(I).Dur = Val(Parse(N + 2))
            SaveMapItem(I).X = Val(Parse(N + 3))
            SaveMapItem(I).Y = Val(Parse(N + 4))
            
            N = N + 5
        Next I
        
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapnpcdata" Then
        N = 1
        
        For I = 1 To MAX_MAP_NPCS
            SaveMapNpc(I).Num = Val(Parse(N))
            SaveMapNpc(I).X = Val(Parse(N + 1))
            SaveMapNpc(I).Y = Val(Parse(N + 2))
            SaveMapNpc(I).Dir = Val(Parse(N + 3))
            
            N = N + 4
        Next I
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapdone" Then
        Map = SaveMap
        
        For I = 1 To MAX_MAP_ITEMS
            MapItem(I) = SaveMapItem(I)
        Next I
        
        For I = 1 To MAX_MAP_NPCS
            MapNpc(I) = SaveMapNpc(I)
        Next I
        
        GettingMap = False
        
        ' Play music
        Call StopMidi
        If Map.Music > 0 Then
            If FileExist("music\music" & Trim(str(Map.Music)) & ".mid") Then
                Call PlayMidi("music\music" & Trim(str(Map.Music)) & ".mid")
            End If
        End If
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "saymsg") Or (LCase(Parse(0)) = "broadcastmsg") Or (LCase(Parse(0)) = "globalmsg") Or (LCase(Parse(0)) = "playermsg") Or (LCase(Parse(0)) = "mapmsg") Or (LCase(Parse(0)) = "adminmsg") Then
        Call AddText(Parse(1), Val(Parse(2)))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnitem" Then
        N = Val(Parse(1))
        
        MapItem(N).Num = Val(Parse(2))
        MapItem(N).Value = Val(Parse(3))
        MapItem(N).Dur = Val(Parse(4))
        MapItem(N).X = Val(Parse(5))
        MapItem(N).Y = Val(Parse(6))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "itemeditor") Then
        InItemsEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For I = 1 To MAX_ITEMS
            frmIndex.lstIndex.AddItem I & ": " & Trim(Item(I).Name)
        Next I
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateitem") Then
        N = Val(Parse(1))
        
        ' Update the item
        Item(N).Name = Parse(2)
        Item(N).Pic = Val(Parse(3))
        Item(N).Type = Val(Parse(4))
        Item(N).Data1 = 0
        Item(N).Data2 = 0
        Item(N).Data3 = 0
        Exit Sub
    End If
       
    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "edititem") Then
        N = Val(Parse(1))
        
        ' Update the item
        Item(N).Name = Parse(2)
        Item(N).Pic = Val(Parse(3))
        Item(N).Type = Val(Parse(4))
        Item(N).Data1 = Val(Parse(5))
        Item(N).Data2 = Val(Parse(6))
        Item(N).Data3 = Val(Parse(7))
        
        Item(N).UnBreakable = Val(Parse(8))
        Item(N).Locked = Val(Parse(9))
        Item(N).Disabled = Val(Parse(10))
        Item(N).Assigned = Parse(11)
        
        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnnpc" Then
        N = Val(Parse(1))
        
        MapNpc(N).Num = Val(Parse(2))
        MapNpc(N).X = Val(Parse(3))
        MapNpc(N).Y = Val(Parse(4))
        MapNpc(N).Dir = Val(Parse(5))
        
        ' Client use only
        MapNpc(N).XOffset = 0
        MapNpc(N).YOffset = 0
        MapNpc(N).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "npcdead" Then
        N = Val(Parse(1))
        
        MapNpc(N).Num = 0
        MapNpc(N).X = 0
        MapNpc(N).Y = 0
        MapNpc(N).Dir = 0
        
        ' Client use only
        MapNpc(N).XOffset = 0
        MapNpc(N).YOffset = 0
        MapNpc(N).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "npceditor") Then
        InNpcEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For I = 1 To MAX_NPCS
            frmIndex.lstIndex.AddItem I & ": " & Trim(Npc(I).Name)
        Next I
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatenpc") Then
        N = Val(Parse(1))
        
        ' Update the item
        Npc(N).Name = Parse(2)
        Npc(N).AttackSay = ""
        Npc(N).Sprite = Val(Parse(3))
        Npc(N).SpawnSecs = 0
        Npc(N).Behavior = 0
        Npc(N).Range = 0
        Npc(N).DropChance = 0
        Npc(N).DropItem = 0
        Npc(N).DropItemValue = 0
        Npc(N).str = 0
        Npc(N).DEF = 0
        Npc(N).SPEED = 0
        Npc(N).MAGI = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet :: <- Used for item editor admins only
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "editnpc") Then
        N = Val(Parse(1))
        
        ' Update the npc
        Npc(N).Name = Parse(2)
        Npc(N).AttackSay = Parse(3)
        Npc(N).Sprite = Val(Parse(4))
        Npc(N).SpawnSecs = Val(Parse(5))
        Npc(N).Behavior = Val(Parse(6))
        Npc(N).Range = Val(Parse(7))
        Npc(N).DropChance = Val(Parse(8))
        Npc(N).DropItem = Val(Parse(9))
        Npc(N).DropItemValue = Val(Parse(10))
        Npc(N).str = Val(Parse(11))
        Npc(N).DEF = Val(Parse(12))
        Npc(N).SPEED = Val(Parse(13))
        Npc(N).MAGI = Val(Parse(14))
        
        ' Initialize the npc editor
        Call NpcEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "mapkey") Then
        X = Val(Parse(1))
        Y = Val(Parse(2))
        N = Val(Parse(3))
        
        TempTile(X, Y).DoorOpen = N
        
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "editmap") Then
        Call EditorInit
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "shopeditor") Then
        InShopEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For I = 1 To MAX_SHOPS
            frmIndex.lstIndex.AddItem I & ": " & Trim(Shop(I).Name)
        Next I
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateshop") Then
        N = Val(Parse(1))
        
        ' Update the shop name
        Shop(N).Name = Parse(2)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "editshop") Then
        ShopNum = Val(Parse(1))
        
        ' Update the shop
        Shop(ShopNum).Name = Parse(2)
        Shop(ShopNum).JoinSay = Parse(3)
        Shop(ShopNum).LeaveSay = Parse(4)
        Shop(ShopNum).FixesItems = Val(Parse(5))
        
        N = 6
        For I = 1 To MAX_TRADES
            
            GiveItem = Val(Parse(N))
            GiveValue = Val(Parse(N + 1))
            GetItem = Val(Parse(N + 2))
            GetValue = Val(Parse(N + 3))
            
            Shop(ShopNum).TradeItem(I).GiveItem = GiveItem
            Shop(ShopNum).TradeItem(I).GiveValue = GiveValue
            Shop(ShopNum).TradeItem(I).GetItem = GetItem
            Shop(ShopNum).TradeItem(I).GetValue = GetValue
            
            N = N + 4
        Next I
        
        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "spelleditor") Then
        InSpellEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For I = 1 To MAX_SPELLS
            frmIndex.lstIndex.AddItem I & ": " & Trim(Spell(I).Name)
        Next I
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatespell") Then
        N = Val(Parse(1))
        
        ' Update the spell name
        Spell(N).Name = Parse(2)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "editspell") Then
        N = Val(Parse(1))
        
        ' Update the spell
        Spell(N).Name = Parse(2)
        Spell(N).ClassReq = Val(Parse(3))
        Spell(N).LevelReq = Val(Parse(4))
        Spell(N).Type = Val(Parse(5))
        Spell(N).Data1 = Val(Parse(6))
        Spell(N).Data2 = Val(Parse(7))
        Spell(N).Data3 = Val(Parse(8))
                        
        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (LCase(Parse(0)) = "trade") Then
        ShopNum = Val(Parse(1))
        If Val(Parse(2)) = 1 Then
            frmTrade.picFixItems.Visible = True
        Else
            frmTrade.picFixItems.Visible = False
        End If
        
        N = 3
        For I = 1 To MAX_TRADES
            GiveItem = Val(Parse(N))
            GiveValue = Val(Parse(N + 1))
            GetItem = Val(Parse(N + 2))
            GetValue = Val(Parse(N + 3))
            
            If GiveItem > 0 And GetItem > 0 Then
                frmTrade.lstTrade.AddItem "Give " & Trim(Shop(ShopNum).Name) & " " & GiveValue & " " & Trim(Item(GiveItem).Name) & " for " & GetValue & " " & Trim(Item(GetItem).Name)
            Else
                frmTrade.lstTrade.AddItem "<<Empty Slot>>"
            End If
            N = N + 4
        Next I
        
        If frmTrade.lstTrade.ListCount > 0 Then
            frmTrade.lstTrade.ListIndex = 0
        End If
        frmTrade.Show vbModal
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (LCase(Parse(0)) = "spells") Then
        
        frmMirage.picPlayerSpells.Visible = True
        frmMirage.lstSpells.Clear
        
        ' Put spells known in player record
        For I = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(I) = Val(Parse(I))
            If Player(MyIndex).Spell(I) <> 0 Then
                frmMirage.lstSpells.AddItem I & ": " & Trim(Spell(Player(MyIndex).Spell(I)).Name)
            Else
                frmMirage.lstSpells.AddItem "<free spells slot>"
            End If
        Next I
        
        frmMirage.lstSpells.ListIndex = 0
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "weather") Then
        GameWeather = Val(Parse(1))
    End If

    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (LCase(Parse(0)) = "time") Then
        GameTime = Val(Parse(1))
        Exit Sub
    End If
End Sub
