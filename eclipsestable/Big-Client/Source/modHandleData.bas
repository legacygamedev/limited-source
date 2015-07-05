Attribute VB_Name = "modHandleData"
Option Explicit

Sub HandleData(ByVal data As String)
    Dim Parse() As String
    Dim Dir As Long
    Dim i As Long, n As Long, X As Long, Y As Long, p As Long
    Dim shopNum As Long
    Dim z As Long
    Dim Strfilename As String
    Dim CustomX As Long
    Dim CustomY As Long
    Dim CustomIndex As Long
    Dim Customcolour As Long
    Dim Customsize As Long
    Dim Customtext As String
    Dim Casestring As String
    Dim Packet As String
    Dim M As Long
    Dim J As Long

    ' Handle Data
    Parse = Split(data, SEP_CHAR)

    ' Add packet info to debugger
    If frmDebug.Visible = True Then
        Call TextAdd(frmDebug.txtDebug, time & " - ( " & Parse(0) & " )", True)
    End If

    'Determine whats send
    Casestring = LCase$(Parse(0))
    
    'Switching to cases instead of if's
    Select Case LCase$(Parse(0))
    
        'Request for the script editor
        Case "maineditor"
            Call Packet_MainEditor(Parse(1), Parse(2))
            Exit Sub
            
        Case "totalonline"
            Call Packet_TotalOnline(Parse(1))
            Exit Sub
            
        Case "edithouse"
            Call Packet_EditHouse
            Exit Sub
            
        Case "leaveparty211"
            Call Packet_LeaveParty
            Exit Sub
        
        Case "playerhpreturn"
            Call Packet_PlayerHpreturn(Parse(1), Parse(2), Parse(3))
            Exit Sub
            
        Case "maxinfo"
            Call Packet_MaxInfo(Parse)
            Exit Sub
            
        Case "npchp"
            Call Packet_NpcHP(Parse(1), Parse(2), Parse(3))
            Exit Sub
            
        Case "alertmsg"
            Call Packet_AlertMsg(Parse(1))
            Exit Sub
            
        Case "plainmsg"
            Call Packet_PlainMsg(Parse(1), Parse(2))
            Exit Sub
            
        Case "allchars"
            Call Packet_AllChars(Parse)
            Exit Sub
            
        Case "loginok"
            Call Packet_LoginOk(Parse(1))
            Exit Sub
            
        Case "news"
            Call Packet_News(Parse(1), Parse(5), Parse(2), Parse(3), Parse(4))
            Exit Sub
            
        Case "newcharclasses"
            Call Packet_NewCharClasses(Parse)
            Exit Sub
            
        Case "classesdata"
            Call Packet_ClassData(Parse)
            Exit Sub
            
        Case "gameclock"
            Call Packet_GameClock(Parse(1), Parse(2), Parse(3), Parse(4))
            Exit Sub
            
        Case "ingame"
            Call Packet_Ingame
            Exit Sub
            
        Case "playerinv"
            Call Packet_PlayerInv(Parse)
            Exit Sub
            
        Case "playerinvupdate"
            Call Packet_PlayerInvUpdate(Parse(1), Parse(2), Parse(3), Parse(4), Parse(5))
            Exit Sub
        
        Case "playerbank"
            Call Packet_PlayerBank(Parse)
            Exit Sub
            
        Case "playerbankupdate"
            Call Packet_PlayerbankUpdate(Parse(1), Parse(2), Parse(3), Parse(4))
            Exit Sub
        
        Case "openbank"
            Call Packet_OpenBank
            Exit Sub
            
        Case "bankmsg"
            Call Packet_BankMsg(Parse(1))
            Exit Sub
            
        Case "playerworneq"
            Call Packet_PlayerWornEQ(Parse)
            Exit Sub
            
        Case "playerpoints"
            Call Packet_PlayerPoints(Parse(1))
            Exit Sub
            
        Case "cussprite"
            Call Packet_CusSprite(Parse(1), Parse(2), Parse(3), Parse(4))
            Exit Sub
            
        Case "playerhp"
            Call Packet_PlayerHp(Parse(1), Parse(2), Parse(3))
            Exit Sub
            
        Case "playerexp"
            Call Packet_PlayerExp(Parse(1), Parse(2))
            Exit Sub
            
        Case "playermp"
            Call Packet_PlayerMp(Parse(1), Parse(2))
            Exit Sub
            
        Case "playersp"
            Call Packet_PlayerSp(Parse(1), Parse(2))
            Exit Sub
        
        Case "mapmsg2"
            Call Packet_MapMsg2(Parse(1), Parse(2))
            Exit Sub
            
        Case "scriptbubble"
            Call Packet_Scriptbubble(Parse(1), Parse(2), Parse(3), Parse(4), Parse(5), Parse(6))
            Exit Sub
            
        Case "playerstatspacket"
            Call Packet_PlayerStatsPacket(Parse(1), Parse(2), Parse(3), Parse(4), Parse(5), Parse(6), Parse(7))
            Exit Sub
            
        Case "playerdata"
            Call Packet_PlayerData(Parse)
            Exit Sub
            
        Case "leave"
            Call Packet_Leave(Parse(1))
            Exit Sub
            
        Case "left"
            Call Packet_Left(Parse(1))
            Exit Sub
            
        Case "playerlevel"
            Call Packet_PlayerLevel(Parse(1), Parse(2))
            Exit Sub
            
        Case "updatesprite"
            Call Packet_UpdateSprite(Parse(1), Parse(2))
            Exit Sub
            
        Case "playermove"
            Call Packet_PlayerMove(Parse(1), Parse(2), Parse(3), Parse(4), Parse(5))
            Exit Sub
            
        Case "npcmove"
            Call Packet_NpcMove(Parse(1), Parse(2), Parse(3), Parse(4), Parse(5))
            Exit Sub
            
        Case "playerdir"
            Call Packet_PlayerDir(Parse(1), Parse(2))
            Exit Sub
            
        Case "npcdir"
            Call Packet_NpcDir(Parse(1), Parse(2))
            Exit Sub
            
        Case "playerxy"
            Call Packet_PlayerXY(Parse(1), Parse(2), Parse(3))
            Exit Sub
            
        Case "removemembers"
            Call Packet_RemoveMembers
            Exit Sub
            
        Case "updatemembers"
            Call Packet_UpdateMembers(Parse(1), Parse(2))
            Exit Sub
            
        Case "attack"
            Call Packet_PlayerAttack(Parse(1))
            Exit Sub
            
        Case "npcattack"
            Call Packet_NpcAttack(Parse(1))
            Exit Sub
        
    End Select


    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (Casestring = "checkformap") Then
        GettingMap = True
    
        ' Erase all players except self
        For i = 1 To MAX_PLAYERS
            If i <> MyIndex Then
                Call SetPlayerMap(i, 0)
            End If
        Next i

        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        X = Val(Parse(1))

        ' Get revision
        Y = Val(Parse(2))
        
        ' Close map editor if player leaves current map
        If InEditor Then
            ScreenMode = 0
            NightMode = 0
            GridMode = 0
            InEditor = False
            Unload frmMapEditor
            frmMapEditor.MousePointer = 1
            frmStable.MousePointer = 1
        End If
        

        If FileExists("maps\map" & X & ".dat") Then
            ' Check to see if the revisions match
            If GetMapRevision(X) = Y Then
            ' We do so we dont need the map

                ' Load the map
                Call LoadMap(X)

                Call SendData("needmap" & SEP_CHAR & "no" & END_CHAR)
                Exit Sub
            End If
        End If

        ' Either the revisions didn't match or we dont have the map, so we need it
        Call SendData("needmap" & SEP_CHAR & "yes" & END_CHAR)
        Exit Sub
    End If

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::

    If Casestring = "mapdata" Then
        n = 1

        Map(Val(Parse(1))).name = Parse(n + 1)
        Map(Val(Parse(1))).Revision = Val(Parse(n + 2))
        Map(Val(Parse(1))).Moral = Val(Parse(n + 3))
        Map(Val(Parse(1))).Up = Val(Parse(n + 4))
        Map(Val(Parse(1))).Down = Val(Parse(n + 5))
        Map(Val(Parse(1))).Left = Val(Parse(n + 6))
        Map(Val(Parse(1))).Right = Val(Parse(n + 7))
        Map(Val(Parse(1))).music = Parse(n + 8)
        Map(Val(Parse(1))).BootMap = Val(Parse(n + 9))
        Map(Val(Parse(1))).BootX = Val(Parse(n + 10))
        Map(Val(Parse(1))).BootY = Val(Parse(n + 11))
        Map(Val(Parse(1))).Indoors = Val(Parse(n + 12))
        Map(Val(Parse(1))).Weather = Val(Parse(n + 13))

        n = n + 14

        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(Val(Parse(1))).Tile(X, Y).Ground = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, Y).mask = Val(Parse(n + 1))
                Map(Val(Parse(1))).Tile(X, Y).Anim = Val(Parse(n + 2))
                Map(Val(Parse(1))).Tile(X, Y).Mask2 = Val(Parse(n + 3))
                Map(Val(Parse(1))).Tile(X, Y).M2Anim = Val(Parse(n + 4))
                Map(Val(Parse(1))).Tile(X, Y).Fringe = Val(Parse(n + 5))
                Map(Val(Parse(1))).Tile(X, Y).FAnim = Val(Parse(n + 6))
                Map(Val(Parse(1))).Tile(X, Y).Fringe2 = Val(Parse(n + 7))
                Map(Val(Parse(1))).Tile(X, Y).F2Anim = Val(Parse(n + 8))
                Map(Val(Parse(1))).Tile(X, Y).Type = Val(Parse(n + 9))
                Map(Val(Parse(1))).Tile(X, Y).Data1 = Val(Parse(n + 10))
                Map(Val(Parse(1))).Tile(X, Y).Data2 = Val(Parse(n + 11))
                Map(Val(Parse(1))).Tile(X, Y).Data3 = Val(Parse(n + 12))
                Map(Val(Parse(1))).Tile(X, Y).String1 = Parse(n + 13)
                Map(Val(Parse(1))).Tile(X, Y).String2 = Parse(n + 14)
                Map(Val(Parse(1))).Tile(X, Y).String3 = Parse(n + 15)
                Map(Val(Parse(1))).Tile(X, Y).light = Val(Parse(n + 16))
                Map(Val(Parse(1))).Tile(X, Y).GroundSet = Val(Parse(n + 17))
                Map(Val(Parse(1))).Tile(X, Y).MaskSet = Val(Parse(n + 18))
                Map(Val(Parse(1))).Tile(X, Y).AnimSet = Val(Parse(n + 19))
                Map(Val(Parse(1))).Tile(X, Y).Mask2Set = Val(Parse(n + 20))
                Map(Val(Parse(1))).Tile(X, Y).M2AnimSet = Val(Parse(n + 21))
                Map(Val(Parse(1))).Tile(X, Y).FringeSet = Val(Parse(n + 22))
                Map(Val(Parse(1))).Tile(X, Y).FAnimSet = Val(Parse(n + 23))
                Map(Val(Parse(1))).Tile(X, Y).Fringe2Set = Val(Parse(n + 24))
                Map(Val(Parse(1))).Tile(X, Y).F2AnimSet = Val(Parse(n + 25))
                n = n + 26
            Next X
        Next Y

        For X = 1 To 15
            Map(Val(Parse(1))).Npc(X) = Val(Parse(n))
            Map(Val(Parse(1))).SpawnX(X) = Val(Parse(n + 1))
            Map(Val(Parse(1))).SpawnY(X) = Val(Parse(n + 2))
            n = n + 3
        Next X

        ' Save the map
        Call SaveLocalMap(Val(Parse(1)))

        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmMapEditor.Visible = False
            frmMapEditor.Visible = False
            frmStable.Show
' frmMirage.picMapEditor.Visible = False

            If frmMapWarp.Visible Then
                Unload frmMapWarp
            End If

            If frmMapProperties.Visible Then
                Unload frmMapProperties
            End If
        End If

        Exit Sub
    End If

    If Casestring = "tilecheck" Then
        n = 5
        X = Val(Parse(2))
        Y = Val(Parse(3))

        Select Case Val(Parse(4))
            Case 0
                Map(Val(Parse(1))).Tile(X, Y).Ground = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, Y).GroundSet = Val(Parse(n + 1))
            Case 1
                Map(Val(Parse(1))).Tile(X, Y).mask = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, Y).MaskSet = Val(Parse(n + 1))
            Case 2
                Map(Val(Parse(1))).Tile(X, Y).Anim = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, Y).AnimSet = Val(Parse(n + 1))
            Case 3
                Map(Val(Parse(1))).Tile(X, Y).Mask2 = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, Y).Mask2Set = Val(Parse(n + 1))
            Case 4
                Map(Val(Parse(1))).Tile(X, Y).M2Anim = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, Y).M2AnimSet = Val(Parse(n + 1))
            Case 5
                Map(Val(Parse(1))).Tile(X, Y).Fringe = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, Y).FringeSet = Val(Parse(n + 1))
            Case 6
                Map(Val(Parse(1))).Tile(X, Y).FAnim = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, Y).FAnimSet = Val(Parse(n + 1))
            Case 7
                Map(Val(Parse(1))).Tile(X, Y).Fringe2 = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, Y).Fringe2Set = Val(Parse(n + 1))
            Case 8
                Map(Val(Parse(1))).Tile(X, Y).F2Anim = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, Y).F2AnimSet = Val(Parse(n + 1))
        End Select
        Call SaveLocalMap(Val(Parse(1)))
    End If

    If Casestring = "tilecheckattribute" Then
        n = 5
        X = Val(Parse(2))
        Y = Val(Parse(3))

        Map(Val(Parse(1))).Tile(X, Y).Type = Val(Parse(n - 1))
        Map(Val(Parse(1))).Tile(X, Y).Data1 = Val(Parse(n))
        Map(Val(Parse(1))).Tile(X, Y).Data2 = Val(Parse(n + 1))
        Map(Val(Parse(1))).Tile(X, Y).Data3 = Val(Parse(n + 2))
        Map(Val(Parse(1))).Tile(X, Y).String1 = Parse(n + 3)
        Map(Val(Parse(1))).Tile(X, Y).String2 = Parse(n + 4)
        Map(Val(Parse(1))).Tile(X, Y).String3 = Parse(n + 5)
        Call SaveLocalMap(Val(Parse(1)))
    End If

    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If Casestring = "mapitemdata" Then
        n = 1

        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).Num = Val(Parse(n))
            SaveMapItem(i).value = Val(Parse(n + 1))
            SaveMapItem(i).Dur = Val(Parse(n + 2))
            SaveMapItem(i).X = Val(Parse(n + 3))
            SaveMapItem(i).Y = Val(Parse(n + 4))

            n = n + 5
        Next i

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If Casestring = "mapnpcdata" Then
        n = 1

        For i = 1 To 15
            SaveMapNpc(i).Num = Val(Parse(n))
            SaveMapNpc(i).X = Val(Parse(n + 1))
            SaveMapNpc(i).Y = Val(Parse(n + 2))
            SaveMapNpc(i).Dir = Val(Parse(n + 3))

            n = n + 4
        Next i

        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If Casestring = "mapdone" Then
        ' Map = SaveMap

        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i

        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i

        GettingMap = False

        ' Play music
        If Trim$(Map(GetPlayerMap(MyIndex)).music) <> "None" Then
            Call MapMusic(Map(GetPlayerMap(MyIndex)).music)
        End If

        If GameWeather = WEATHER_RAINING Then
            Call PlayBGS("rain.wav")
        End If
        If GameWeather = WEATHER_THUNDER Then
            Call PlayBGS("thunder.wav")
        End If

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (Casestring = "saymsg") Or (Casestring = "broadcastmsg") Or (Casestring = "globalmsg") Or (Casestring = "playermsg") Or (Casestring = "mapmsg") Or (Casestring = "adminmsg") Then
        Call AddText(Parse(1), Val(Parse(2)))
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If Casestring = "spawnitem" Then
        n = Val(Parse(1))

        MapItem(n).Num = Val(Parse(2))
        MapItem(n).value = Val(Parse(3))
        MapItem(n).Dur = Val(Parse(4))
        MapItem(n).X = Val(Parse(5))
        MapItem(n).Y = Val(Parse(6))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    If (Casestring = "itemeditor") Then
        InItemsEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Item(i).name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (Casestring = "updateitem") Then
        n = Val(Parse(1))

        ' Update the item
        Item(n).name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Item(n).StrReq = Val(Parse(8))
        Item(n).DefReq = Val(Parse(9))
        Item(n).SpeedReq = Val(Parse(10))
        Item(n).MagicReq = Val(Parse(11))
        Item(n).ClassReq = Val(Parse(12))
        Item(n).AccessReq = Val(Parse(13))

        Item(n).AddHP = Val(Parse(14))
        Item(n).AddMP = Val(Parse(15))
        Item(n).AddSP = Val(Parse(16))
        Item(n).AddSTR = Val(Parse(17))
        Item(n).AddDEF = Val(Parse(18))
        Item(n).AddMAGI = Val(Parse(19))
        Item(n).AddSpeed = Val(Parse(20))
        Item(n).AddEXP = Val(Parse(21))
        Item(n).desc = Parse(22)
        Item(n).AttackSpeed = Val(Parse(23))
        Item(n).Price = Val(Parse(24))
        Item(n).Stackable = Val(Parse(25))
        Item(n).Bound = Val(Parse(26))
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (Casestring = "edititem") Then
        n = Val(Parse(1))

        ' Update the item
        Item(n).name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Item(n).StrReq = Val(Parse(8))
        Item(n).DefReq = Val(Parse(9))
        Item(n).SpeedReq = Val(Parse(10))
        Item(n).MagicReq = Val(Parse(11))
        Item(n).ClassReq = Val(Parse(12))
        Item(n).AccessReq = Val(Parse(13))

        Item(n).AddHP = Val(Parse(14))
        Item(n).AddMP = Val(Parse(15))
        Item(n).AddSP = Val(Parse(16))
        Item(n).AddSTR = Val(Parse(17))
        Item(n).AddDEF = Val(Parse(18))
        Item(n).AddMAGI = Val(Parse(19))
        Item(n).AddSpeed = Val(Parse(20))
        Item(n).AddEXP = Val(Parse(21))
        Item(n).desc = Parse(22)
        Item(n).AttackSpeed = Val(Parse(23))
        Item(n).Price = Val(Parse(24))
        Item(n).Stackable = Val(Parse(25))
        Item(n).Bound = Val(Parse(26))

        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: mouse packet  ::
    ' :::::::::::::::::::
    If (Casestring = "mouse") Then
        Player(MyIndex).input = 1
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' ::Weather Packet::
    ' ::::::::::::::::::
    If (Casestring = "mapweather") Then
        If 0 + Val(Parse(1)) <> 0 Then
            Map(Val(Parse(1))).Weather = Val(Parse(2))
            If Val(Parse(1)) = 2 Then
                frmStable.tmrSnowDrop.Interval = Val(Parse(3))
            ElseIf Val(Parse(1)) = 1 Then
                frmStable.tmrRainDrop.Interval = Val(Parse(3))
            End If
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If Casestring = "spawnnpc" Then
        n = Val(Parse(1))

        MapNpc(n).Num = Val(Parse(2))
        MapNpc(n).X = Val(Parse(3))
        MapNpc(n).Y = Val(Parse(4))
        MapNpc(n).Dir = Val(Parse(5))
        MapNpc(n).Big = Val(Parse(6))

        ' Client use only
        MapNpc(n).xOffset = 0
        MapNpc(n).yOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If Casestring = "npcdead" Then
        n = Val(Parse(1))

        MapNpc(n).Num = 0
        MapNpc(n).X = 0
        MapNpc(n).Y = 0
        MapNpc(n).Dir = 0

        ' Client use only
        MapNpc(n).xOffset = 0
        MapNpc(n).yOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    If (Casestring = "npceditor") Then
        InNpcEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Npc(i).name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (Casestring = "updatenpc") Then
        n = Val(Parse(1))

        ' Update the item
        Npc(n).name = Parse(2)
        Npc(n).AttackSay = vbNullString
        Npc(n).Sprite = Val(Parse(3))
        Npc(n).SpriteSize = Val(Parse(4))
        
        ' That's all well and good but it also resets our NPC - Pickle
        ' Npc(n).SpawnSecs = 0
        ' Npc(n).Behavior = 0
        ' Npc(n).Range = 0
        ' For i = 1 To MAX_NPC_DROPS
        ' Npc(n).ItemNPC(i).chance = 0
        ' Npc(n).ItemNPC(i).ItemNum = 0
        ' Npc(n).ItemNPC(i).ItemValue = 0
        ' Next i
        ' Npc(n).STR = 0
        ' Npc(n).DEF = 0
        ' Npc(n).speed = 0
        ' Npc(n).MAGI = 0
        
        Npc(n).Big = Val(Parse(5))
        Npc(n).MaxHp = Val(Parse(6))
        ' Npc(n).Exp = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet ::
    ' :::::::::::::::::::::
    If (Casestring = "editnpc") Then
        n = Val(Parse(1))

        ' Update the npc
        Npc(n).name = Parse(2)
        Npc(n).AttackSay = Parse(3)
        Npc(n).Sprite = Val(Parse(4))
        Npc(n).SpawnSecs = Val(Parse(5))
        Npc(n).Behavior = Val(Parse(6))
        Npc(n).Range = Val(Parse(7))
        Npc(n).STR = Val(Parse(8))
        Npc(n).DEF = Val(Parse(9))
        Npc(n).speed = Val(Parse(10))
        Npc(n).MAGI = Val(Parse(11))
        Npc(n).Big = Val(Parse(12))
        Npc(n).MaxHp = Val(Parse(13))
        Npc(n).Exp = Val(Parse(14))
        Npc(n).SpawnTime = Val(Parse(15))
        Npc(n).Element = Val(Parse(16))
        Npc(n).SpriteSize = Val(Parse(17))

        ' Call GlobalMsg("At editnpc..." & Npc(n).Element)
        z = 18
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).chance = Val(Parse(z))
            Npc(n).ItemNPC(i).ItemNum = Val(Parse(z + 1))
            Npc(n).ItemNPC(i).ItemValue = Val(Parse(z + 2))
            z = z + 3
        Next i

        ' Initialize the npc editor
        Call NpcEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (Casestring = "mapkey") Then
        X = Val(Parse(1))
        Y = Val(Parse(2))
        n = Val(Parse(3))

        TempTile(X, Y).DoorOpen = n

        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
    If (Casestring = "editmap") Then
        Call EditorInit
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    If (Casestring = "shopeditor") Then
        InShopEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Shop(i).name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (Casestring = "updateshop") Then
        n = Val(Parse(1))

        ' Update the shop name
        Shop(n).name = Parse(2)
        Shop(n).FixesItems = Val(Parse(3))
        Shop(n).BuysItems = Val(Parse(4))
        Shop(n).ShowInfo = Val(Parse(5))
        Shop(n).currencyItem = Val(Parse(6))

        M = 7
        ' Get shop items
        For i = 1 To MAX_SHOP_ITEMS
            Shop(n).ShopItem(i).ItemNum = Val(Parse(M))
            Shop(n).ShopItem(i).Amount = Val(Parse(M + 1))
            Shop(n).ShopItem(i).Price = Val(Parse(M + 2))
            M = M + 3
        Next i

        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    If (Casestring = "editshop") Then

        shopNum = Val(Parse(1))

        ' Update the shop
        Shop(shopNum).name = Parse(2)
        Shop(shopNum).FixesItems = Val(Parse(3))
        Shop(shopNum).BuysItems = Val(Parse(4))
        Shop(shopNum).ShowInfo = Val(Parse(5))
        Shop(shopNum).currencyItem = Val(Parse(6))

        M = 7
        For i = 1 To 25
            Shop(shopNum).ShopItem(i).ItemNum = Val(Parse(M))
            Shop(shopNum).ShopItem(i).Amount = Val(Parse(M + 1))
            Shop(shopNum).ShopItem(i).Price = Val(Parse(M + 2))
            M = M + 3
        Next i

        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    If (Casestring = "spelleditor") Then
        InSpellEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Spell(i).name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (Casestring = "updatespell") Then
        n = Val(Parse(1))

        ' Update the spell name
        Spell(n).name = Parse(2)
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (Casestring = "editspell") Then
        n = Val(Parse(1))

        ' Update the spell
        Spell(n).name = Parse(2)
        Spell(n).ClassReq = Val(Parse(3))
        Spell(n).LevelReq = Val(Parse(4))
        Spell(n).Type = Val(Parse(5))
        Spell(n).Data1 = Val(Parse(6))
        Spell(n).Data2 = Val(Parse(7))
        Spell(n).Data3 = Val(Parse(8))
        Spell(n).MPCost = Val(Parse(9))
        Spell(n).Sound = Val(Parse(10))
        Spell(n).Range = Val(Parse(11))
        Spell(n).SpellAnim = Val(Parse(12))
        Spell(n).SpellTime = Val(Parse(13))
        Spell(n).SpellDone = Val(Parse(14))
        Spell(n).AE = Val(Parse(15))
        Spell(n).Big = Val(Parse(16))
        Spell(n).Element = Val(Parse(17))


        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (Casestring = "goshop") Then
        shopNum = Val(Parse(1))
        ' Show the shop
        Call GoShop(shopNum)
    End If

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (Casestring = "spells") Then

        frmStable.picPlayerSpells.Visible = True
        frmStable.lstSpells.Clear

        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(i) = Val(Parse(i))
            If Player(MyIndex).Spell(i) <> 0 Then
                frmStable.lstSpells.addItem i & ": " & Trim$(Spell(Player(MyIndex).Spell(i)).name)
            Else
                frmStable.lstSpells.addItem "--- Slot Free ---"
            End If
        Next i

        frmStable.lstSpells.ListIndex = 0
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (Casestring = "weather") Then
        If Val(Parse(1)) = WEATHER_RAINING And GameWeather <> WEATHER_RAINING Then
            Call AddText("You see drops of rain falling from the sky above!", BRIGHTGREEN)
            Call PlayBGS("rain.mp3")
        End If
        If Val(Parse(1)) = WEATHER_THUNDER And GameWeather <> WEATHER_THUNDER Then
            Call AddText("You see thunder in the sky above!", BRIGHTGREEN)
            Call PlayBGS("thunder.mp3")
        End If
        If Val(Parse(1)) = WEATHER_SNOWING And GameWeather <> WEATHER_SNOWING Then
            Call AddText("You see snow falling from the sky above!", BRIGHTGREEN)
        End If

        If Val(Parse(1)) = WEATHER_NONE Then
            If GameWeather = WEATHER_RAINING Then
                Call AddText("The rain beings to calm.", BRIGHTGREEN)
                Call StopSound
            ElseIf GameWeather = WEATHER_SNOWING Then
                Call AddText("The snow is melting away.", BRIGHTGREEN)
            ElseIf GameWeather = WEATHER_THUNDER Then
                Call AddText("The thunder begins to disapear.", BRIGHTGREEN)
                Call StopSound
            End If
        End If
        GameWeather = Val(Parse(1))
        RainIntensity = Val(Parse(2))
        If MAX_RAINDROPS <> RainIntensity Then
            MAX_RAINDROPS = RainIntensity
            ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
            ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
        End If
    End If

    ' ::::::::::::::::::::::::::::::::
    ' :: playername coloring packet ::
    ' ::::::::::::::::::::::::::::::::
    If (Casestring = "namecolor") Then
        Player(MyIndex).color = Val(Parse(1))
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: image packet      ::
    ' :::::::::::::::::::::::
    If (LCase$(Parse(0)) = "fog") Then
        rec.top = Int(Val(Parse(4)))
        rec.Bottom = Int(Val(Parse(5)))
        rec.Left = Int(Val(Parse(6)))
        rec.Right = Int(Val(Parse(7)))
        Call DD_BackBuffer.BltFast(Val(Parse(1)), Val(Parse(2)), DD_TileSurf(Val(Parse(3))), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Get Online List ::
    ' ::::::::::::::::::::::::::
    If Casestring = "onlinelist" Then
        frmStable.lstOnline.Clear

        n = 2
        z = Val(Parse(1))
        For X = n To (z + 1)
            frmStable.lstOnline.addItem Trim$(Parse(n))
            n = n + 2
        Next X
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
    If Casestring = "blitplayerdmg" Then
        DmgDamage = Val(Parse(1))
        NPCWho = Val(Parse(2))
        DmgTime = GetTickCount
        iii = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
    If Casestring = "blitnpcdmg" Then
        NPCDmgDamage = Val(Parse(1))
        NPCDmgTime = GetTickCount
        ii = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::::::
    ' :: Retrieve the player's inventory ::
    ' :::::::::::::::::::::::::::::::::::::
    If Casestring = "pptrading" Then
        frmPlayerTrade.Items1.Clear
        frmPlayerTrade.Items2.Clear
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = vbNullString
            Trading(i).InvVal = 0
            Trading2(i).InvNum = 0
            Trading2(i).InvName = vbNullString
            Trading2(i).InvVal = 0
            frmPlayerTrade.Items1.addItem i & ": <Nothing>"
            frmPlayerTrade.Items2.addItem i & ": <Nothing>"
        Next i

        frmPlayerTrade.Items1.ListIndex = 0

        Call UpdateTradeInventory
        frmPlayerTrade.Show vbModeless, frmStable
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
        If Casestring = "qtrade" Then
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = vbNullString
            Trading(i).InvVal = 0
            Trading2(i).InvNum = 0
            Trading2(i).InvName = vbNullString
            Trading2(i).InvVal = 0
        Next i

        frmPlayerTrade.Command1.ForeColor = &H0&
        frmPlayerTrade.Command2.ForeColor = &H0&

        frmPlayerTrade.Visible = False
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Disable Time ::
    ' ::::::::::::::::::
    If Casestring = "dtime" Then
        If Parse(1) = "True" Then
            frmStable.lblGameTime.Caption = vbNullString
            frmStable.lblGameClock.Caption = vbNullString
            frmStable.lblGameTime.Visible = False
            frmStable.lblGameClock.Visible = False
            frmStable.tmrGameClock.Enabled = False
        Else
            frmStable.lblGameTime.Caption = "It is now:"
            frmStable.lblGameClock.Caption = vbNullString
            frmStable.lblGameTime.Visible = True
            frmStable.lblGameClock.Visible = True
            frmStable.tmrGameClock.Enabled = True
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If Casestring = "updatetradeitem" Then
        n = Val(Parse(1))

        Trading2(n).InvNum = Val(Parse(2))
        Trading2(n).InvName = Parse(3)
        Trading2(n).InvVal = Val(Parse(4))

        If Trading2(n).InvNum <= 0 Then
            frmPlayerTrade.Items2.List(n - 1) = n & ": <Nothing>"
            Exit Sub
            End If
            
             If Item(Trading2(n).InvNum).Type = ITEM_TYPE_CURRENCY Or Item(Trading2(n).InvNum).Stackable = 1 Then
                       frmPlayerTrade.Items2.List(n - 1) = n & ": " & Trim$(Trading2(n).InvName) & " / " & Int(Trading2(n).InvVal) & ""
            Else
                       frmPlayerTrade.Items2.List(n - 1) = n & ": " & Trim$(Trading2(n).InvName) & ""
            End If
            Exit Sub
        End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If Casestring = "trading" Then
        n = Val(Parse(1))
        If n = 0 Then
            frmPlayerTrade.Command2.ForeColor = &H0&
        End If
        If n = 1 Then
            frmPlayerTrade.Command2.ForeColor = &HFF00&
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Chat System Packets ::
    ' :::::::::::::::::::::::::
    If Casestring = "ppchatting" Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Label1.Caption = "Chatting With: " & Trim$(Player(Val(Parse(1))).name)

        frmPlayerChat.Show vbModeless, frmStable
        Exit Sub
    End If

    If Casestring = "qchat" Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Visible = False
        frmPlayerTrade.Command2.ForeColor = &H8000000F
        frmPlayerTrade.Command1.ForeColor = &H8000000F
        Exit Sub
    End If

    If Casestring = "sendchat" Then
        Dim s As String

        s = vbNewLine & GetPlayerName(Val(Parse(2))) & "> " & Trim$(Parse(1))
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text)
        frmPlayerChat.txtChat.SelColor = QBColor(BROWN)
        frmPlayerChat.txtChat.SelText = s
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text) - 1
        Exit Sub
    End If
' :::::::::::::::::::::::::::::
' :: END Chat System Packets ::
' :::::::::::::::::::::::::::::

    ' :::::::::::::::::::::::
    ' :: Play Sound Packet ::
    ' :::::::::::::::::::::::
    If Casestring = "sound" Then
        s = LCase$(Parse(1))
        Select Case Trim$(s)
            Case "attack"
                Call PlaySound("sword.wav")
            Case "critical"
                Call PlaySound("critical.wav")
            Case "miss"
                Call PlaySound("miss.wav")
            Case "key"
                Call PlaySound("key.wav")
            Case "magic"
                Call PlaySound("magic" & Val(Parse(2)) & ".wav")
            Case "warp"
                If FileExists("SFX\warp.wav") Then
                    Call PlaySound("warp.wav")
                End If
            Case "pain"
                Call PlaySound("pain.wav")
            Case "newmsg"
                Call PlaySound("magic18.wav")
            Case "soundattribute"
                Call PlaySound(Trim$(Parse(2)))
        End Select
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Sprite Change Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If Casestring = "spritechange" Then
        If Val(Parse(1)) = 1 Then
            i = MsgBox("Would you like to buy this sprite?", 4, "Buying Sprite")
            If i = 6 Then
                Call SendData("buysprite" & END_CHAR)
            End If
        Else
            Call SendData("buysprite" & END_CHAR)
        End If
        Exit Sub
    End If
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: House Buy Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If Casestring = "housebuy" Then
        If Val(Parse(1)) = 1 Then
            i = MsgBox("Would you like to buy this house?", 4, "Buying House")
            If i = 6 Then
                Call SendData("buyhouse" & END_CHAR)
            End If
        Else
            Call SendData("buyhouse" & END_CHAR)
        End If
        Exit Sub
    End If
    
    If Casestring = "housesell" Then
        i = MsgBox("Would you like to sell this house?", 4, "Selling House")
        If i = 6 Then
            Call SendData("sellhouse" & END_CHAR)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::::::::
    ' :: Change Player Direction Packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If Casestring = "changedir" Then
        Player(Val(Parse(2))).Dir = Val(Parse(1))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Flash Movie Event Packet ::
    ' ::::::::::::::::::::::::::::::
    If Casestring = "flashevent" Then
        If LCase$(Mid$(Trim$(Parse(1)), 1, 7)) = "http://" Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopBGM
            Call StopSound
            frmFlash.Flash.LoadMovie 0, Trim$(Parse(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmStable
        ElseIf FileExists("Flashs\" & Trim$(Parse(1))) = True Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopBGM
            Call StopSound
            frmFlash.Flash.LoadMovie 0, App.Path & "\Flashs\" & Trim$(Parse(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmStable
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If Casestring = "prompt" Then
        i = MsgBox(Trim$(Parse(1)), vbYesNo)
        Call SendData("prompt" & SEP_CHAR & i & SEP_CHAR & Val(Parse(2)) & END_CHAR)
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If Casestring = "querybox" Then
        frmQuery.Label1.Caption = Trim$(Parse(1))
        frmQuery.Label2.Caption = Parse(2)
        frmQuery.Show vbModal
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Emoticon editor packet ::
    ' ::::::::::::::::::::::::::::
    If (Casestring = "emoticoneditor") Then
        InEmoticonEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For i = 0 To MAX_EMOTICONS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Emoticons(i).Command)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Element editor packet ::
    ' ::::::::::::::::::::::::::::
    If (Casestring = "elementeditor") Then
        InElementEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For i = 0 To MAX_ELEMENTS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Element(i).name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    If (Casestring = "editelement") Then
        n = Val(Parse(1))

        Element(n).name = Parse(2)
        Element(n).Strong = Val(Parse(3))
        Element(n).Weak = Val(Parse(4))

        Call ElementEditorInit
        Exit Sub
    End If

    If (Casestring = "updateelement") Then
        n = Val(Parse(1))

        Element(n).name = Parse(2)
        Element(n).Strong = Val(Parse(3))
        Element(n).Weak = Val(Parse(4))
        Exit Sub
    End If

    If (Casestring = "editemoticon") Then
        n = Val(Parse(1))

        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))

        Call EmoticonEditorInit
        Exit Sub
    End If

    If (Casestring = "updateemoticon") Then
        n = Val(Parse(1))

        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Arrow editor packet ::
    ' ::::::::::::::::::::::::::::
    If (Casestring = "arroweditor") Then
        InArrowEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For i = 1 To MAX_ARROWS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Arrows(i).name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    If (Casestring = "updatearrow") Then
        n = Val(Parse(1))

        Arrows(n).name = Parse(2)
        Arrows(n).Pic = Val(Parse(3))
        Arrows(n).Range = Val(Parse(4))
        Arrows(n).Amount = Val(Parse(5))
        Exit Sub
    End If

    If (Casestring = "editarrow") Then
        n = Val(Parse(1))

        Arrows(n).name = Parse(2)

        Call ArrowEditorInit
        Exit Sub
    End If

    If (Casestring = "updatearrow") Then
        n = Val(Parse(1))

        Arrows(n).name = Parse(2)
        Arrows(n).Pic = Val(Parse(3))
        Arrows(n).Range = Val(Parse(4))
        Arrows(n).Amount = Val(Parse(5))
        Exit Sub
    End If

    If (Casestring = "hookshot") Then
        n = Val(Parse(1))
        i = Val(Parse(3))

        Player(n).HookShotAnim = Arrows(Val(Parse(2))).Pic
        Player(n).HookShotTime = GetTickCount
        Player(n).HookShotToX = Val(Parse(4))
        Player(n).HookShotToY = Val(Parse(5))
        Player(n).HookShotX = GetPlayerX(n)
        Player(n).HookShotY = GetPlayerY(n)
        Player(n).HookShotSucces = Val(Parse(6))
        Player(n).HookShotDir = Val(Parse(3))

        Call PlaySound("grapple.wav")
        Call PlaySound("grapple-fire.wav")

        If i = DIR_DOWN Then
            Player(n).HookShotX = GetPlayerX(n)
            Player(n).HookShotY = GetPlayerY(n) + 1
            If Player(n).HookShotX - 1 > MAX_MAPY Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If i = DIR_UP Then
            Player(n).HookShotX = GetPlayerX(n)
            Player(n).HookShotY = GetPlayerY(n) - 1
            If Player(n).HookShotY + 1 < 0 Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If i = DIR_RIGHT Then
            Player(n).HookShotX = GetPlayerX(n) + 1
            Player(n).HookShotY = GetPlayerY(n)
            If Player(n).HookShotX - 1 > MAX_MAPX Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If i = DIR_LEFT Then
            Player(n).HookShotX = GetPlayerX(n) - 1
            Player(n).HookShotY = GetPlayerY(n)
            If Player(n).HookShotX + 1 < 0 Then
                Player(n).Arrow(X).Arrow = 0
                Exit Sub
            End If
        End If
        Exit Sub
    End If

    If (Casestring = "checkarrows") Then
        n = Val(Parse(1))
        z = Val(Parse(2))
        i = Val(Parse(3))

        For X = 1 To MAX_PLAYER_ARROWS
            If Player(n).Arrow(X).Arrow = 0 Then
                Player(n).Arrow(X).Arrow = 1
                Player(n).Arrow(X).ArrowNum = z
                Player(n).Arrow(X).ArrowAnim = Arrows(z).Pic
                Player(n).Arrow(X).ArrowTime = GetTickCount
                Player(n).Arrow(X).ArrowVarX = 0
                Player(n).Arrow(X).ArrowVarY = 0
                Player(n).Arrow(X).ArrowY = GetPlayerY(n)
                Player(n).Arrow(X).ArrowX = GetPlayerX(n)
                Player(n).Arrow(X).ArrowAmount = p

                If i = DIR_DOWN Then
                    Player(n).Arrow(X).ArrowY = GetPlayerY(n) + 1
                    Player(n).Arrow(X).ArrowPosition = 0
                    If Player(n).Arrow(X).ArrowY - 1 > MAX_MAPY Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If i = DIR_UP Then
                    Player(n).Arrow(X).ArrowY = GetPlayerY(n) - 1
                    Player(n).Arrow(X).ArrowPosition = 1
                    If Player(n).Arrow(X).ArrowY + 1 < 0 Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If i = DIR_RIGHT Then
                    Player(n).Arrow(X).ArrowX = GetPlayerX(n) + 1
                    Player(n).Arrow(X).ArrowPosition = 2
                    If Player(n).Arrow(X).ArrowX - 1 > MAX_MAPX Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If i = DIR_LEFT Then
                    Player(n).Arrow(X).ArrowX = GetPlayerX(n) - 1
                    Player(n).Arrow(X).ArrowPosition = 3
                    If Player(n).Arrow(X).ArrowX + 1 < 0 Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                Exit For
            End If
        Next X
        Exit Sub
    End If

    If (Casestring = "checksprite") Then
        n = Val(Parse(1))

        Player(n).Sprite = Val(Parse(2))
        Exit Sub
    End If

    If (Casestring = "mapreport") Then
        n = 1

        frmMapReport.lstIndex.Clear
        For i = 1 To MAX_MAPS
            frmMapReport.lstIndex.addItem i & ": " & Trim$(Parse(n))
            n = n + 1
        Next i

        frmMapReport.Show vbModeless, frmStable
        Exit Sub
    End If

    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (Casestring = "time") Then
        'Call AddText("GameTime Is: " & Parse(1), GREEN)
        GameTime = Val(Parse(1))
        If GameTime = TIME_DAY Then
            Call AddText("Day has dawned in this realm.", WHITE)
        Else
            Call AddText("Night has fallen upon the weary eyed nightowls.", WHITE)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Spell anim packet ::
    ' :::::::::::::::::::::::
    If (Casestring = "spellanim") Then
        Dim SpellNum As Long
        SpellNum = Val(Parse(1))

        Spell(SpellNum).SpellAnim = Val(Parse(2))
        Spell(SpellNum).SpellTime = Val(Parse(3))
        Spell(SpellNum).SpellDone = Val(Parse(4))
        Spell(SpellNum).Big = Val(Parse(9))

        Player(Val(Parse(5))).SpellNum = SpellNum

        For i = 1 To MAX_SPELL_ANIM
            If Player(Val(Parse(5))).SpellAnim(i).CastedSpell = NO Then
                Player(Val(Parse(5))).SpellAnim(i).SpellDone = 0
                Player(Val(Parse(5))).SpellAnim(i).SpellVar = 0
                Player(Val(Parse(5))).SpellAnim(i).SpellTime = GetTickCount
                Player(Val(Parse(5))).SpellAnim(i).TargetType = Val(Parse(6))
                Player(Val(Parse(5))).SpellAnim(i).Target = Val(Parse(7))
                Player(Val(Parse(5))).SpellAnim(i).CastedSpell = YES
                Exit For
            End If
        Next i
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Script Spell anim packet ::
    ' :::::::::::::::::::::::
    If (Casestring = "scriptspellanim") Then
        Spell(Val(Parse(1))).SpellAnim = Val(Parse(2))
        Spell(Val(Parse(1))).SpellTime = Val(Parse(3))
        Spell(Val(Parse(1))).SpellDone = Val(Parse(4))
        Spell(Val(Parse(1))).Big = Val(Parse(7))


        For i = 1 To MAX_SCRIPTSPELLS
            If ScriptSpell(i).CastedSpell = NO Then
                ScriptSpell(i).SpellNum = Val(Parse(1))
                ScriptSpell(i).SpellDone = 0
                ScriptSpell(i).SpellVar = 0
                ScriptSpell(i).SpellTime = GetTickCount
                ScriptSpell(i).X = Val(Parse(5))
                ScriptSpell(i).Y = Val(Parse(6))
                ScriptSpell(i).CastedSpell = YES
                Exit For
            End If
        Next i
        Exit Sub
    End If

    If (Casestring = "checkemoticons") Then
        n = Val(Parse(1))

        Player(n).EmoticonNum = Val(Parse(2))
        Player(n).EmoticonTime = GetTickCount
        Player(n).EmoticonVar = 0
        Exit Sub
    End If


    If Casestring = "levelup" Then
        Player(Val(Parse(1))).LevelUpT = GetTickCount
        Player(Val(Parse(1))).LevelUp = 1
        Exit Sub
    End If

    If Casestring = "damagedisplay" Then
        For i = 1 To MAX_BLT_LINE
            If Val(Parse(1)) = 0 Then
                If BattlePMsg(i).index <= 0 Then
                    BattlePMsg(i).index = 1
                    BattlePMsg(i).Msg = Parse(2)
                    BattlePMsg(i).color = Val(Parse(3))
                    BattlePMsg(i).time = GetTickCount
                    BattlePMsg(i).Done = 1
                    BattlePMsg(i).Y = 0
                    Exit Sub
                Else
                    BattlePMsg(i).Y = BattlePMsg(i).Y - 15
                End If
            Else
                If BattleMMsg(i).index <= 0 Then
                    BattleMMsg(i).index = 1
                    BattleMMsg(i).Msg = Parse(2)
                    BattleMMsg(i).color = Val(Parse(3))
                    BattleMMsg(i).time = GetTickCount
                    BattleMMsg(i).Done = 1
                    BattleMMsg(i).Y = 0
                    Exit Sub
                Else
                    BattleMMsg(i).Y = BattleMMsg(i).Y - 15
                End If
            End If
        Next i

        z = 1
        If Val(Parse(1)) = 0 Then
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then
                    If BattlePMsg(i).Y < BattlePMsg(i + 1).Y Then
                        z = i
                    End If
                Else
                    If BattlePMsg(i).Y < BattlePMsg(1).Y Then
                        z = i
                    End If
                End If
            Next i

            BattlePMsg(z).index = 1
            BattlePMsg(z).Msg = Parse(2)
            BattlePMsg(z).color = Val(Parse(3))
            BattlePMsg(z).time = GetTickCount
            BattlePMsg(z).Done = 1
            BattlePMsg(z).Y = 0
        Else
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then
                    If BattleMMsg(i).Y < BattleMMsg(i + 1).Y Then
                        z = i
                    End If
                Else
                    If BattleMMsg(i).Y < BattleMMsg(1).Y Then
                        z = i
                    End If
                End If
            Next i

            BattleMMsg(z).index = 1
            BattleMMsg(z).Msg = Parse(2)
            BattleMMsg(z).color = Val(Parse(3))
            BattleMMsg(z).time = GetTickCount
            BattleMMsg(z).Done = 1
            BattleMMsg(z).Y = 0
        End If
        Exit Sub
    End If

    If Casestring = "itembreak" Then
        ItemDur(Val(Parse(1))).Item = Val(Parse(2))
        ItemDur(Val(Parse(1))).Dur = Val(Parse(3))
        ItemDur(Val(Parse(1))).Done = 1
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::::::::::::
    ' :: Index player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::::::::
    If Casestring = "itemworn" Then
        Player(Val(Parse(1))).Armor = Val(Parse(2))
        Player(Val(Parse(1))).Weapon = Val(Parse(3))
        Player(Val(Parse(1))).Helmet = Val(Parse(4))
        Player(Val(Parse(1))).Shield = Val(Parse(5))
        Player(Val(Parse(1))).legs = Val(Parse(6))
        Player(Val(Parse(1))).Ring = Val(Parse(7))
        Player(Val(Parse(1))).Necklace = Val(Parse(8))
        Exit Sub
    End If

    If Casestring = "scripttile" Then
        frmScript.lblScript.Caption = Parse(1)
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Set player speed ::
    ' ::::::::::::::::::::::
    If Casestring = "setspeed" Then
        SetSpeed Parse(1), Val(Parse(2))
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Custom Menu  ::
    ' ::::::::::::::::::
    If (Casestring = "showcustommenu") Then
        ' Error handling
        If Not FileExists(Parse(2)) Then
            Call MsgBox(Parse(2) & " not found. Menu loading aborted. Please contact a GM to fix this problem.", vbExclamation)
            Exit Sub
        End If

        CUSTOM_TITLE = Parse(1)
        CUSTOM_IS_CLOSABLE = Val(Parse(3))

        frmCustom1.picBackground.top = 0
        frmCustom1.picBackground.Left = 0
        frmCustom1.picBackground = LoadPicture(App.Path & Parse(2))
        frmCustom1.Height = PixelsToTwips(24 + frmCustom1.picBackground.Height, 1)
        frmCustom1.Width = PixelsToTwips(6 + frmCustom1.picBackground.Width, 0)
        frmCustom1.Visible = True

        Exit Sub
    End If

    If (Casestring = "closecustommenu") Then

        CUSTOM_TITLE = "CLOSED"
        Unload frmCustom1

        Exit Sub
    End If

    If (Casestring = "loadpiccustommenu") Then

        CustomIndex = Parse(1)
        Strfilename = Parse(2)
        CustomX = Val(Parse(3))
        CustomY = Val(Parse(4))
        
        If Not IsInArray(frmCustom1.picCustom, CInt(CustomIndex)) Then
            Load frmCustom1.picCustom(CustomIndex)
        End If

        If Strfilename = vbNullString Then
            Strfilename = "MEGAUBERBLANKNESSOFUNHOLYPOWER" 'smooth :\    -Pickle
        End If

        If FileExists(Strfilename) = True Then
            frmCustom1.picCustom(CustomIndex) = LoadPicture(App.Path & Strfilename)
            frmCustom1.picCustom(CustomIndex).top = CustomY
            frmCustom1.picCustom(CustomIndex).Left = CustomX
            frmCustom1.picCustom(CustomIndex).Visible = True
        Else
            frmCustom1.picCustom(CustomIndex).Picture = LoadPicture()
            frmCustom1.picCustom(CustomIndex).Visible = False
        End If

        Exit Sub
    End If

    If (Casestring = "loadlabelcustommenu") Then

        CustomIndex = Parse(1)
        Strfilename = Parse(2)
        CustomX = Val(Parse(3))
        CustomY = Val(Parse(4))
        Customsize = Val(Parse(5))
        Customcolour = Val(Parse(6))
        
        If Not IsInArray(frmCustom1.BtnCustom, CInt(CustomIndex)) Then
            Load frmCustom1.BtnCustom(CustomIndex)
        End If

        frmCustom1.BtnCustom(CustomIndex).Caption = Strfilename
        frmCustom1.BtnCustom(CustomIndex).top = CustomY
        frmCustom1.BtnCustom(CustomIndex).Left = CustomX
        frmCustom1.BtnCustom(CustomIndex).Font.Bold = True
        frmCustom1.BtnCustom(CustomIndex).Font.Size = Customsize
        frmCustom1.BtnCustom(CustomIndex).ForeColor = QBColor(Customcolour)
        frmCustom1.BtnCustom(CustomIndex).Visible = True
        frmCustom1.BtnCustom(CustomIndex).Alignment = Parse(7)

        If Parse(8) <= 0 Or Parse(9) <= 0 Then
            frmCustom1.BtnCustom(CustomIndex).AutoSize = True
        Else
            frmCustom1.BtnCustom(CustomIndex).AutoSize = False
            frmCustom1.BtnCustom(CustomIndex).Width = Parse(8)
            frmCustom1.BtnCustom(CustomIndex).Height = Parse(9)
        End If

        Exit Sub
    End If

    If (Casestring = "loadtextboxcustommenu") Then

        CustomIndex = Parse(1)
        Strfilename = Parse(2)
        CustomX = Val(Parse(3))
        CustomY = Val(Parse(4))
        Customtext = Parse(5)
        
        If Not IsInArray(frmCustom1.txtCustom, CInt(CustomIndex)) Then
            Load frmCustom1.txtCustom(CustomIndex)
            Load frmCustom1.txtcustomOK(CustomIndex)
        End If

        frmCustom1.txtCustom(CustomIndex).Text = Customtext
        frmCustom1.txtCustom(CustomIndex).top = CustomY
        frmCustom1.txtCustom(CustomIndex).Left = Strfilename
        frmCustom1.txtCustom(CustomIndex).Width = CustomX - 32
        frmCustom1.txtcustomOK(CustomIndex).top = CustomY
        frmCustom1.txtcustomOK(CustomIndex).Left = frmCustom1.txtCustom(CustomIndex).Left + frmCustom1.txtCustom(CustomIndex).Width
        frmCustom1.txtcustomOK(CustomIndex).Visible = True
        frmCustom1.txtCustom(CustomIndex).Visible = True

        Exit Sub
    End If

    If (Casestring = "loadinternetwindow") Then
        Customtext = Parse(1)
        ' DEBUG STRING
        ' Call AddText(customtext, 15)
        ShellExecute 1, "open", Trim(Customtext), vbNullString, vbNullString, 1
        Exit Sub
    End If

    If (Casestring = "returncustomboxmsg") Then
        Customsize = Parse(1)

        Packet = "returningcustomboxmsg" & SEP_CHAR & frmCustom1.txtCustom(Customsize).Text & END_CHAR
        Call SendData(Packet)

        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Sound Stuff              ::
    ' ::::::::::::::::::::::::::::::
    
    'Play mapmusic
    If Casestring = "playbgm" Then
        Call PlayBGM(Trim$(Map(GetPlayerMap(MyIndex)).music))
        Exit Sub
    End If
    
    'Stop mapmusic
    If Casestring = "stopbgm" Then
        Call StopBGM
        Exit Sub
    End If
    
    'Play Background Sound
    If Casestring = "bkgsound" Then
        Call PlayBGS(Parse(1))
        Exit Sub
    End If
    
    'Stop Sound
    If Casestring = "stopsound" Then
        Call StopSound
        Exit Sub
    End If
    
    If (Casestring = "playernewxy") Then
        X = Val(Parse(1))
        Y = Val(Parse(2))

        If Not GetPlayerX(MyIndex) = X Then Call SetPlayerX(MyIndex, X)
        If Not GetPlayerY(MyIndex) = Y Then Call SetPlayerY(MyIndex, Y)

        Exit Sub
    End If
    
    'Email system
    If (Casestring = "setmsgbody") Then
        If Parse(4) = 1 Then
            frmInbox.txtSender.Text = Parse(1)
            frmInbox.txtSubject.Text = Parse(2)
            frmInbox.txtBody.Text = Parse(3)
            frmInbox.txtBody2.Text = Parse(3)
        Else
            frmInbox.txtReceiver2.Text = Parse(1)
            frmInbox.txtSubject.Text = Parse(2)
            frmInbox.txtBody.Text = Parse(3)
            frmInbox.txtBody3.Text = Parse(3)
        End If
        Exit Sub
    End If
    
    If (Casestring = "xobniym") Then
        If Val(Parse(2)) <= 9 Then
            frmInbox.lstMail.addItem "[000" & Parse(2) & "]" & " " & Parse(1) & " - " & Parse(3)
        ElseIf Val(Parse(2)) >= 9 And Val(Parse(2)) <= 99 Then
            frmInbox.lstMail.addItem "[00" & Parse(2) & "]" & " " & Parse(1) & " - " & Parse(3)
        ElseIf Val(Parse(2)) > 99 Then
            frmInbox.lstMail.addItem "[0" & Parse(2) & "]" & " " & Parse(1) & " - " & Parse(3)
        Else
            frmInbox.lstMail.addItem "[" & Parse(2) & "]" & " " & Parse(1) & " - " & Parse(3)
        End If
        Exit Sub
    End If
    
    If (Casestring = "myoutbox") Then
        If Val(Parse(2)) <= 9 Then
            frmInbox.lstOutbox.addItem "[000" & Parse(2) & "]" & " " & Parse(1) & " - " & Parse(3)
        ElseIf Val(Parse(2)) >= 9 And Val(Parse(2)) <= 99 Then
            frmInbox.lstOutbox.addItem "[00" & Parse(2) & "]" & " " & Parse(1) & " - " & Parse(3)
        ElseIf Val(Parse(2)) > 99 Then
            frmInbox.lstOutbox.addItem "[0" & Parse(2) & "]" & " " & Parse(1) & " - " & Parse(3)
        Else
            frmInbox.lstOutbox.addItem "[" & Parse(2) & "]" & " " & Parse(1) & " - " & Parse(3)
        End If
        Exit Sub
    End If
    If (Casestring = "unreadmsg") Then
        Dim E As Long
        Dim Ending1 As String
        If Parse(1) > 0 Then
            For E = 1 To 3
                If E = 1 Then Ending1 = ".gif"
                If E = 2 Then Ending1 = ".jpg"
                If E = 3 Then Ending1 = ".bmp"

                If FileExists("GUI\Inbox" & Ending1) Then
                    frmStable.CmdMail.Picture = LoadPicture(App.Path & "\GUI\Inbox2" & Ending1)
                End If
            Next E
        Else
            For E = 1 To 3
                If E = 1 Then Ending1 = ".gif"
                If E = 2 Then Ending1 = ".jpg"
                If E = 3 Then Ending1 = ".bmp"

                If FileExists("GUI\Inbox2" & Ending1) Then
                    frmStable.CmdMail.Picture = LoadPicture(App.Path & "\GUI\Inbox" & Ending1)
                End If
            Next E
        End If
    End If

End Sub

Public Sub Packet_MainEditor(ByVal filename As String, ByVal Text As String)
     'Enable this piece of code to create a clientsided script folder when it not exists, its needed to work with it,
     'but could post a security problem...
    'If LCase(Dir(App.Path & "\Scripts", vbDirectory)) <> "scripts" Then
        'Call MkDir(App.Path & "\Scripts")
    'End If
    
    AFileName = filename
         
    Dim f
    f = FreeFile
    Open App.Path & "\Scripts\" & AFileName For Output As #f
        Print #f, Text
    Close #f
    
    Unload frmEditor
    frmEditor.Show
End Sub

Public Sub Packet_TotalOnline(ByVal total As Long)
    frmMainMenu.LblTotalOnline.Caption = "Total Players Online: " & Trim(total)
End Sub

Public Sub Packet_LeaveParty()
    Dim i As Long
    For i = 1 To MAX_PARTY_MEMBERS
            Player(MyIndex).Party.Member(i) = 0
    Next
End Sub

Public Sub Packet_PlayerHpreturn(ByVal Player, ByVal HP, ByVal MaxHp)
    Player(Val(Player)).HP = Val(HP)
    Player(Val(Player)).MaxHp = Val(MaxHp)
    ' Call MsgBox("player(" & Val(Player) & ").hp = " & Val(HP))
    ' Call BltPlayerBars(val(Player))
End Sub

Public Sub Packet_MaxInfo(ByRef Parse() As String)
    Dim i As Long
    
    ' Set the global configuration values.
    GAME_NAME = Trim$(Parse(1))
    MAX_PLAYERS = Val(Parse(2))
    MAX_ITEMS = Val(Parse(3))
    MAX_NPCS = Val(Parse(4))
    MAX_SHOPS = Val(Parse(5))
    MAX_SPELLS = Val(Parse(6))
    MAX_MAPS = Val(Parse(7))
    MAX_MAP_ITEMS = Val(Parse(8))
    MAX_MAPX = Val(Parse(9))
    MAX_MAPY = Val(Parse(10))
    MAX_EMOTICONS = Val(Parse(11))
    MAX_ELEMENTS = Val(Parse(12))
    paperdoll = Val(Parse(13))
    SpriteSize = Val(Parse(14))
    MAX_SCRIPTSPELLS = Val(Parse(15))
    CustomPlayers = Val(Parse(16))
    lvl = Val(Parse(17))
    MAX_PARTY_MEMBERS = Val(Parse(18))
    STAT1 = Parse(19)
    STAT2 = Parse(20)
    STAT3 = Parse(21)
    STAT4 = Parse(22)
    WalkFix = Parse(23)
    
    frmMainMenu.lblVersion.Caption = "Version: " & Parse(24)

    If 0 + CustomPlayers > 0 Then
        frmNewChar.Picture4.Visible = False
        frmNewChar.HScroll1.Visible = True
        frmNewChar.HScroll2.Visible = True
        frmNewChar.HScroll3.Visible = True
        frmNewChar.Label14.Visible = True
        frmNewChar.Label11.Visible = True
        frmNewChar.Label12.Visible = True
        frmNewChar.Picture1.Visible = True

        If FileExists("GFX\Heads.bmp") Then
            frmNewChar.iconn(0).Picture = LoadPicture(App.Path & "\GFX\Heads.bmp")
        End If
        If FileExists("GFX\Bodys.bmp") Then
            frmNewChar.iconn(1).Picture = LoadPicture(App.Path & "\GFX\Bodys.bmp")
        End If
        If FileExists("GFX\Legs.bmp") Then
            frmNewChar.iconn(2).Picture = LoadPicture(App.Path & "\GFX\Legs.bmp")
        End If


        If SpriteSize = 1 Then
            frmNewChar.iconn(0).Left = -Val(5 * PIC_X)
            frmNewChar.iconn(0).top = -Val(PIC_Y - 15)

            frmNewChar.iconn(1).Left = -Val(5 * PIC_X)
            frmNewChar.iconn(1).top = -Val(PIC_Y - 7)

            frmNewChar.iconn(2).Left = -Val(5 * PIC_X)
            frmNewChar.iconn(2).top = -Val(PIC_Y + 3)
        Else
            frmNewChar.iconn(0).Left = -Val(5 * PIC_X)
            frmNewChar.iconn(0).top = -Val(PIC_Y)

            frmNewChar.iconn(1).Left = -Val(5 * PIC_X)
            frmNewChar.iconn(1).top = -Val(PIC_Y)

            frmNewChar.iconn(2).Left = -Val(5 * PIC_X)
            frmNewChar.iconn(2).top = -Val(PIC_Y)
        End If
    End If

    ReDim Map(0 To MAX_MAPS) As MapRec
    ReDim Player(1 To MAX_PLAYERS) As PlayerRec
    ReDim Item(1 To MAX_ITEMS) As ItemRec
    ReDim Npc(1 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Element(0 To MAX_ELEMENTS) As ElementRec
    ReDim Bubble(1 To MAX_PLAYERS) As ChatBubble
    ReDim ScriptBubble(1 To MAX_BUBBLES) As ScriptBubble
    ReDim SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim ScriptSpell(1 To MAX_SCRIPTSPELLS) As ScriptSpellAnimRec

    For i = 1 To MAX_MAPS
        ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Next
    
    ReDim TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
    ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
    ReDim MapReport(1 To MAX_MAPS) As MapRec
    
    MAX_SPELL_ANIM = MAX_MAPX * MAX_MAPY

    MAX_BLT_LINE = 6
    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec

    For i = 1 To MAX_PLAYERS
        ReDim Player(i).SpellAnim(1 To MAX_SPELL_ANIM) As SpellAnimRec
    Next

    For i = 0 To MAX_EMOTICONS
        Emoticons(i).Pic = 0
        Emoticons(i).Command = vbNullString
    Next

    Call ClearTempTile

    ' Clear out players
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
    Next

    For i = 1 To MAX_MAPS
        Call LoadMap(i)
    Next

    frmStable.Caption = Trim$(GAME_NAME)
    App.Title = GAME_NAME

    AllDataReceived = True

End Sub

Public Sub Packet_NpcHP(ByVal NpcNum As Long, ByVal HP As Long, ByVal MaxHp As Long)
    MapNpc(NpcNum).HP = Val(HP)
    MapNpc(NpcNum).MaxHp = Val(MaxHp)
End Sub

Public Sub Packet_AlertMsg(ByVal Msg As String)
    frmStable.Visible = False
    frmSendGetData.Visible = False
    frmMainMenu.Visible = True

    Call MsgBox(Msg, vbOKOnly, GAME_NAME)
End Sub

Public Sub Packet_PlainMsg(ByVal Msg As String, ByVal Form As Long)
    frmSendGetData.Visible = False

    If Form = 0 Then frmMainMenu.Show
    
    If Form = 1 Then frmNewAccount.Show
    
    If Form = 2 Then frmDeleteAccount.Show

    If Form = 3 Then frmLogin.Show

    If Form = 4 Then frmNewChar.Show

    If Form = 5 Then frmChars.Show

    Call MsgBox(Msg, vbOKOnly, GAME_NAME)
End Sub

Public Sub Packet_AllChars(ByRef Parse() As String)
    Dim name As String
    Dim Class As String
    Dim Level As Long
    Dim Loops As Long
    Dim Number As Long

    ' Hide and show forms
    frmSendGetData.Visible = False
    frmChars.Visible = True

    ' Clear the character list.
    frmChars.lstChars.Clear

    ' Start the index as the first packet.
    Number = 1

    ' Loop through all of the characters.
    For Loops = 1 To MAX_CHARS

        ' Get the character data from the packet.
        name = Parse(Number)
        Class = Parse(Number + 1)
        Level = CLng(Parse(Number + 2))

        ' Display the character information to the user.
        If Trim$(name) = vbNullString Then
            frmChars.lstChars.addItem ("Free Character Slot")
        Else
            frmChars.lstChars.addItem (name & " a level " & Level & " " & Class)
        End If

        ' Start the index at the next character.
        Number = Number + 3
    Next

    ' Set the first available item.
    frmChars.lstChars.ListIndex = 0

End Sub

Public Sub Packet_LoginOk(ByVal index As Long)
    ' Now we can receive game data
    MyIndex = Val(index)

    frmSendGetData.Visible = True
    frmChars.Visible = False

    ReDim Player(MyIndex).Party.Member(1 To MAX_PARTY_MEMBERS)

    Call SetStatus("Receiving game data...")
End Sub

Public Sub Packet_News(ByVal News As String, ByVal desc As String, ByVal RED As Long, ByVal GREEN As Long, ByVal BLUE As Long)
    Call WriteINI("DATA", "News", News, (App.Path & "\News.ini"))
    Call WriteINI("DATA", "Desc", desc, (App.Path & "\News.ini"))
    Call WriteINI("COLOR", "Red", CInt(RED), (App.Path & "\News.ini"))
    Call WriteINI("COLOR", "Green", CInt(GREEN), (App.Path & "\News.ini"))
    Call WriteINI("COLOR", "Blue", CInt(BLUE), (App.Path & "\News.ini"))

    ' We just gots the news, so change the news label
    Call ParseNews
End Sub

Public Sub Packet_NewCharClasses(ByRef Parse() As String)
    Dim Loops As Long
    Dim Number As Long

    ' Start the index at the first packet.
    Number = 1

    ' Get the maximum amount of classes.
    Max_Classes = CLng(Parse(1))

    ' Get the toggle if we're using classes or not.
    ClassesOn = CLng(Parse(2))

    ' ReDim the class array based on the maximum amount of classes.
    ReDim Class(0 To Max_Classes) As ClassRec

    ' Start the index at the third packet.
    Number = 3

    ' Loop through all of the classes in the packet.
    For Loops = 0 To Max_Classes

        ' Get the class name.
        Class(Loops).name = Parse(Number)

        ' Get the class vitals.
        Class(Loops).HP = CLng(Parse(Number + 1))
        Class(Loops).MP = CLng(Parse(Number + 2))
        Class(Loops).SP = CLng(Parse(Number + 3))

        ' Get the class status points.
        Class(Loops).STR = CLng(Parse(Number + 4))
        Class(Loops).DEF = CLng(Parse(Number + 5))
        Class(Loops).speed = CLng(Parse(Number + 6))
        Class(Loops).MAGI = CLng(Parse(Number + 7))

        ' Get the class gender sprites.
        Class(Loops).MaleSprite = CLng(Parse(Number + 8))
        Class(Loops).FemaleSprite = CLng(Parse(Number + 9))

        ' Get the class usable state.
        Class(Loops).Locked = CLng(Parse(Number + 10))

        ' Get the class description.
        Class(Loops).desc = Parse(Number + 11)

        ' Start the index at the next class.
        Number = Number + 12
    Next

    ' Hide the status form.
    frmSendGetData.Visible = False

    ' Show the new character form.
    frmNewChar.Visible = True
    ' Clear the class combo box.
    frmNewChar.cmbClass.Clear

    ' Add the class names to the combo box.
    For Loops = 0 To Max_Classes
        If Class(Loops).Locked = 0 Then
            frmNewChar.cmbClass.addItem Trim$(Class(Loops).name)
        End If
    Next

    ' Select the top-most class name.
    frmNewChar.cmbClass.ListIndex = 0

    ' Check if classes are enabled, and show the combo box.
    If ClassesOn = 0 Then
        frmNewChar.cmbClass.Visible = False
        frmNewChar.lblClassDesc.Visible = False
    Else
        frmNewChar.cmbClass.Visible = True
        frmNewChar.lblClassDesc.Visible = True
    End If

    ' Display the class vitals to the user.
    frmNewChar.lblHP.Caption = CStr(Class(0).HP)
    frmNewChar.lblMP.Caption = CStr(Class(0).MP)
    frmNewChar.lblSP.Caption = CStr(Class(0).SP)

    ' Display the class status points to the user.
    frmNewChar.lblSTR.Caption = CStr(Class(0).STR)
    frmNewChar.lblDEF.Caption = CStr(Class(0).DEF)
    frmNewChar.lblSPEED.Caption = CStr(Class(0).speed)
    frmNewChar.lblMAGI.Caption = CStr(Class(0).MAGI)

    ' Display the class description to the user.
    frmNewChar.lblClassDesc.Caption = Class(0).desc
End Sub

Public Sub Packet_ClassData(ByRef Parse() As String)
    Dim Loops As Long
    Dim Number As Long

    ' Get the maximum amount of classes.
    Max_Classes = CLng(Parse(1))

    ' ReDim the class array based on the maximum amount of classes.
    ReDim Class(0 To Max_Classes) As ClassRec

    ' Start the index at the second packet.
    Number = 2

    For Loops = 0 To Max_Classes

        ' Get the class name.
        Class(Loops).name = Parse(Number)

        ' Get the class vitals.
        Class(Loops).HP = CLng(Parse(Number + 1))
        Class(Loops).MP = CLng(Parse(Number + 2))
        Class(Loops).SP = CLng(Parse(Number + 3))

        ' Get the class status points.
        Class(Loops).STR = CLng(Parse(Number + 4))
        Class(Loops).DEF = CLng(Parse(Number + 5))
        Class(Loops).speed = CLng(Parse(Number + 6))
        Class(Loops).MAGI = CLng(Parse(Number + 7))

        ' Get the class usable state.
        Class(Loops).Locked = CLng(Parse(Number + 8))

        ' Get the class description.
        Class(Loops).desc = Parse(Number + 9)

        ' Start the index at the next class.
        Number = Number + 10
    Next
End Sub

Public Sub Packet_GameClock(ByVal Second As Long, ByVal Minute As Long, ByVal Hour As Long, ByVal speed As Long)
    Seconds = Val(Second)
    Minutes = Val(Minute)
    Hours = Val(Hour)
    
    Gamespeed = Val(speed)
    
    frmStable.lblGameTime.Caption = "It is now:"
    frmStable.lblGameTime.Visible = True
End Sub

Public Sub Packet_Ingame()
    Call GameInit
    Call GameLoop
End Sub

Public Sub Packet_PlayerInv(ByRef Parse() As String)
    Dim n As Long
    Dim index As Long
    Dim i As Long
    
    n = 2
    index = Val(Parse(1))

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(index, i, Val(Parse(n)))
        Call SetPlayerInvItemValue(index, i, Val(Parse(n + 1)))
        Call SetPlayerInvItemDur(index, i, Val(Parse(n + 2)))

        n = n + 3
    Next

    If index = MyIndex Then
        Call UpdateVisInv
    End If
End Sub

Public Sub Packet_PlayerInvUpdate(ByVal slot As Long, ByVal index As Long, ByVal Item As Long, ByVal value As Long, ByVal Dur As Long)

    Call SetPlayerInvItemNum(index, slot, Val(Item))
    Call SetPlayerInvItemValue(index, slot, Val(value))
    Call SetPlayerInvItemDur(index, slot, Val(Dur))
    If index = MyIndex Then
        Call UpdateVisInv
    End If
End Sub

Public Sub Packet_PlayerBank(ByRef Parse() As String)
    Dim n As Long
    Dim slot As Long
    
    n = 1
    For slot = 1 To MAX_BANK
        Call SetPlayerBankItemNum(MyIndex, slot, Val(Parse(n)))
        Call SetPlayerBankItemValue(MyIndex, slot, Val(Parse(n + 1)))
        Call SetPlayerBankItemDur(MyIndex, slot, Val(Parse(n + 2)))

        n = n + 3
    Next

    If frmBank.Visible = True Then
        Call UpdateBank
    End If
End Sub

Public Sub Packet_PlayerbankUpdate(ByVal slot As Long, ByVal ItemNum As Long, ByVal ItemValue As Long, ByVal ItemDur As Long)

    Call SetPlayerBankItemNum(MyIndex, slot, Val(ItemNum))
    Call SetPlayerBankItemValue(MyIndex, slot, Val(ItemValue))
    Call SetPlayerBankItemDur(MyIndex, slot, Val(ItemDur))
    
    If frmBank.Visible = True Then
        Call UpdateBank
    End If
End Sub

Public Sub Packet_OpenBank()
    Dim slot As Long
    
    'Clear lists
    frmBank.lstInventory.Clear
    frmBank.lstBank.Clear
    
    'Get all items in Inventory
    For slot = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, slot) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, slot)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, slot)).Stackable = 1 Then
                frmBank.lstInventory.addItem slot & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, slot)).name) & " (" & GetPlayerInvItemValue(MyIndex, slot) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = slot Or GetPlayerArmorSlot(MyIndex) = slot Or GetPlayerHelmetSlot(MyIndex) = slot Or GetPlayerShieldSlot(MyIndex) = slot Or GetPlayerLegsSlot(MyIndex) = slot Or GetPlayerRingSlot(MyIndex) = slot Or GetPlayerNecklaceSlot(MyIndex) = slot Then
                    frmBank.lstInventory.addItem slot & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, slot)).name) & " (worn)"
                Else
                    frmBank.lstInventory.addItem slot & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, slot)).name)
                End If
            End If
        Else
            frmBank.lstInventory.addItem slot & "> Empty"
        End If

    Next
    
    'Get all items in Bank
    For slot = 1 To MAX_BANK
        If GetPlayerBankItemNum(MyIndex, slot) > 0 Then
            If Item(GetPlayerBankItemNum(MyIndex, slot)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(MyIndex, slot)).Stackable = 1 Then
                frmBank.lstBank.addItem slot & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, slot)).name) & " (" & GetPlayerBankItemValue(MyIndex, slot) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = slot Or GetPlayerArmorSlot(MyIndex) = slot Or GetPlayerHelmetSlot(MyIndex) = slot Or GetPlayerShieldSlot(MyIndex) = slot Or GetPlayerLegsSlot(MyIndex) = slot Or GetPlayerRingSlot(MyIndex) = slot Or GetPlayerNecklaceSlot(MyIndex) = slot Then
                    frmBank.lstBank.addItem slot & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, slot)).name) & " (worn)"
                Else
                    frmBank.lstBank.addItem slot & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, slot)).name)
                End If
            End If
        Else
            frmBank.lstBank.addItem slot & "> Empty"
        End If

    Next
    
    'Select first items on list
    frmBank.lstBank.ListIndex = 0
    frmBank.lstInventory.ListIndex = 0
    
    'Now open bank form
    frmBank.Show vbModal
End Sub

Public Sub Packet_BankMsg(ByVal Msg As String)
    frmBank.lblMsg.Caption = Trim$(Msg)
End Sub

Public Sub Packet_PlayerWornEQ(ByRef Parse() As String)
    Dim index As Long

    'Set index.
    index = Parse(1)

    Call SetPlayerArmorSlot(index, Val(Parse(2)))
    Call SetPlayerWeaponSlot(index, Val(Parse(3)))
    Call SetPlayerHelmetSlot(index, Val(Parse(4)))
    Call SetPlayerShieldSlot(index, Val(Parse(5)))
    Call SetPlayerLegsSlot(index, Val(Parse(6)))
    Call SetPlayerRingSlot(index, Val(Parse(7)))
    Call SetPlayerNecklaceSlot(index, Val(Parse(8)))

    'Show the update
    If index = MyIndex Then Call UpdateVisInv
End Sub

Public Sub Packet_PlayerPoints(ByVal POINTS As Long)
    Player(MyIndex).POINTS = Val(POINTS)

        If GetPlayerPOINTS(MyIndex) > 0 Then
            frmStable.AddSTR.Visible = True
            frmStable.AddDEF.Visible = True
            frmStable.AddSPD.Visible = True
            frmStable.AddMAGI.Visible = True
        Else
            frmStable.AddSTR.Visible = False
            frmStable.AddDEF.Visible = False
            frmStable.AddSPD.Visible = False
            frmStable.AddMAGI.Visible = False
        End If

        frmStable.lblPoints.Caption = Val(POINTS)
End Sub

Public Sub Packet_CusSprite(ByVal index As Long, ByVal head As Long, ByVal body As Long, ByVal legs As Long)
    'Set new sprite pieces
    Player(Val(index)).head = Val(head)
    Player(Val(index)).body = Val(body)
    Player(Val(index)).leg = Val(legs)
End Sub

Public Sub Packet_PlayerHp(ByVal index As Long, ByVal MaxHp As Long, ByVal HP As Long)
    Player(index).MaxHp = Val(MaxHp)
    Call SetPlayerHP(index, Val(HP))
    If index = MyIndex And GetPlayerMaxHP(MyIndex) > 0 Then
        ' frmMirage.shpHP.FillColor = RGB(208, 11, 0)
        frmStable.shpHP.Width = ((GetPlayerHP(MyIndex)) / (GetPlayerMaxHP(MyIndex))) * 150
        frmStable.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
    End If
End Sub

Public Sub Packet_PlayerExp(ByVal Exp As Long, ByVal MaxExp As Long)
    Call SetPlayerExp(MyIndex, Val(Exp))
    frmStable.lblEXP.Caption = Val(Exp) & " / " & Val(MaxExp)
    frmStable.shpTNL.Width = (((Val(Exp)) / (Val(MaxExp))) * 150)
End Sub

Public Sub Packet_PlayerMp(ByVal MaxMP As Long, ByVal MP As Long)
    Player(MyIndex).MaxMP = Val(MaxMP)
    Call SetPlayerMP(MyIndex, Val(MP))
    If GetPlayerMaxMP(MyIndex) > 0 Then
        ' frmMirage.shpMP.FillColor = RGB(208, 11, 0)
        frmStable.shpMP.Width = ((GetPlayerMP(MyIndex)) / (GetPlayerMaxMP(MyIndex))) * 150
        frmStable.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
    End If
End Sub

Public Sub Packet_PlayerSp(ByVal MaxSP As Long, ByVal SP As Long)
    Player(MyIndex).MaxSP = Val(MaxSP)
    Call SetPlayerSP(MyIndex, Val(SP))
    If GetPlayerMaxSP(MyIndex) > 0 Then
        frmStable.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
    End If
End Sub

Public Sub Packet_MapMsg2(ByVal Text As String, ByVal index As Long)
    Bubble(Val(index)).Text = Text
    Bubble(Val(index)).Created = GetTickCount()
End Sub

Public Sub Packet_Scriptbubble(ByVal index As Long, ByVal Text As String, ByVal Map As Long, ByVal X As Long, ByVal Y As Long, ByVal color As Long)
    ScriptBubble(Val(index)).Text = Trim$(Text)
    ScriptBubble(Val(index)).Map = Val(Map)
    ScriptBubble(Val(index)).X = Val(X)
    ScriptBubble(Val(index)).Y = Val(Y)
    ScriptBubble(Val(index)).Colour = Val(color)
    ScriptBubble(Val(index)).Created = GetTickCount()
End Sub

Public Sub Packet_PlayerStatsPacket(ByVal STR As Long, ByVal DEF As Long, ByVal Mag As Long, ByVal Spe As Long, ByVal MaxExp As Long, ByVal Exp As Long, ByVal lvl As Long)
    Dim SubDef As Long, SubMagi As Long, SubSpeed As Long, SubStr As Long
    SubStr = 0
    SubDef = 0
    SubMagi = 0
    SubSpeed = 0

    If GetPlayerWeaponSlot(MyIndex) > 0 Then
        SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSTR
        SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddDEF
        SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddMAGI
        SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSpeed
    End If
    If GetPlayerArmorSlot(MyIndex) > 0 Then
        SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSTR
        SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddDEF
        SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddMAGI
        SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSpeed
    End If
    If GetPlayerShieldSlot(MyIndex) > 0 Then
        SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddSTR
        SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddDEF
        SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddMAGI
        SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddSpeed
    End If
    If GetPlayerHelmetSlot(MyIndex) > 0 Then
        SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddSTR
        SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddDEF
        SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddMAGI
        SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddSpeed
    End If
    If GetPlayerLegsSlot(MyIndex) > 0 Then
        SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddSTR
        SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddDEF
        SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddMAGI
        SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddSpeed
    End If
    If GetPlayerRingSlot(MyIndex) > 0 Then
        SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddSTR
        SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddDEF
        SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddMAGI
        SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddSpeed
    End If
    If GetPlayerNecklaceSlot(MyIndex) > 0 Then
        SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddSTR
        SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddDEF
        SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddMAGI
        SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddSpeed
    End If

    If SubStr > 0 Then
        frmStable.lblSTR.Caption = Val(STR) - SubStr & " (+" & SubStr & ")"
    Else
        frmStable.lblSTR.Caption = Val(STR)
    End If
    If SubDef > 0 Then
        frmStable.lblDEF.Caption = Val(DEF) - SubDef & " (+" & SubDef & ")"
    Else
        frmStable.lblDEF.Caption = Val(DEF)
    End If
    If SubMagi > 0 Then
        frmStable.lblMAGI.Caption = Val(Mag) - SubMagi & " (+" & SubMagi & ")"
    Else
        frmStable.lblMAGI.Caption = Val(Mag)
    End If
    If SubSpeed > 0 Then
        frmStable.lblSPEED.Caption = Val(Spe) - SubSpeed & " (+" & SubSpeed & ")"
    Else
        frmStable.lblSPEED.Caption = Val(Spe)
    End If
    frmStable.lblEXP.Caption = Val(Exp) & " / " & Val(MaxExp)

    frmStable.shpTNL.Width = (((Val(Exp)) / (Val(MaxExp))) * 150)
    frmStable.lblLevel.Caption = Val(lvl)
    Player(MyIndex).Level = Val(lvl)
End Sub

Public Sub Packet_PlayerData(ByRef Parse() As String)
    Dim index As Long
    
    index = Val(Parse(1))
    Call SetPlayerName(index, Parse(2))
    Call SetPlayerSprite(index, Val(Parse(3)))
    Call SetPlayerMap(index, Val(Parse(4)))
    Call SetPlayerX(index, Val(Parse(5)))
    Call SetPlayerY(index, Val(Parse(6)))
    Call SetPlayerDir(index, Val(Parse(7)))
    Call SetPlayerAccess(index, Val(Parse(8)))
    Call SetPlayerPK(index, Val(Parse(9)))
    Call SetPlayerGuild(index, Parse(10))
    Call SetPlayerGuildAccess(index, Val(Parse(11)))
    Call SetPlayerClass(index, Val(Parse(12)))
    Call SetPlayerHead(index, Val(Parse(13)))
    Call SetPlayerBody(index, Val(Parse(14)))
    Call SetPlayerLeg(index, Val(Parse(15)))
    Call SetPlayerPaperdoll(index, Val(Parse(16)))
    Call SetPlayerLevel(index, Val(Parse(17)))

    ' Make sure they aren't walking
    Player(index).Moving = 0
    Player(index).xOffset = 0
    Player(index).yOffset = 0

    ' Check if the player is the client player, and if so reset directions
    If index = MyIndex Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
    End If
End Sub

Public Sub Packet_Leave(ByVal Mapnumber As Long)
    Call SetPlayerMap(CLng(Mapnumber), 0)
End Sub

Public Sub Packet_Left(ByVal index As Long)
    Call ClearPlayer(index)
End Sub

Public Sub Packet_PlayerLevel(ByVal index As Long, Level As String)
    Player(Val(index)).Level = Val(Level)
End Sub

Public Sub Packet_UpdateSprite(ByVal index As Long, ByVal Sprite As Long)
    Call SetPlayerSprite(index, Val(Sprite))
End Sub

Public Sub Packet_PlayerMove(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal Dir As String, ByVal Move As Long)

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerX(index, X)
    Call SetPlayerY(index, Y)
    Call SetPlayerDir(index, Dir)

    Player(index).xOffset = 0
    Player(index).yOffset = 0
    Player(index).Moving = Move
    
    ' Replaced with the one from TE.
    Select Case GetPlayerDir(index)
        Case DIR_UP
            Player(index).yOffset = PIC_Y
        Case DIR_DOWN
            Player(index).yOffset = PIC_Y * -1
        Case DIR_LEFT
            Player(index).xOffset = PIC_X
        Case DIR_RIGHT
            Player(index).xOffset = PIC_X * -1
    End Select
End Sub

Public Sub Packet_NpcMove(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal Dir As String, ByVal Move As Long)

    MapNpc(index).X = X
    MapNpc(index).Y = Y
    MapNpc(index).Dir = Dir
    MapNpc(index).xOffset = 0
    MapNpc(index).yOffset = 0
    MapNpc(index).Moving = 1

    If Move <> 1 Then
        Select Case MapNpc(index).Dir
            Case DIR_UP
                MapNpc(index).yOffset = PIC_Y * Val(Move - 1)
            Case DIR_DOWN
                MapNpc(index).yOffset = PIC_Y * -Move
            Case DIR_LEFT
                MapNpc(index).xOffset = PIC_X * Val(Move - 1)
            Case DIR_RIGHT
                MapNpc(index).xOffset = PIC_X * -Move
        End Select
    Else
        Select Case MapNpc(index).Dir
            Case DIR_UP
                MapNpc(index).yOffset = PIC_Y
            Case DIR_DOWN
                MapNpc(index).yOffset = PIC_Y * -1
            Case DIR_LEFT
                MapNpc(index).xOffset = PIC_X
            Case DIR_RIGHT
                MapNpc(index).xOffset = PIC_X * -1
        End Select
    End If
End Sub

Public Sub Packet_PlayerDir(ByVal index As Long, ByVal Dir As Long)

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)

    Player(index).xOffset = 0
    Player(index).yOffset = 0
    Player(index).MovingH = 0
    Player(index).MovingV = 0
    Player(index).Moving = 0
End Sub

Public Sub Packet_NpcDir(ByVal index As Long, ByVal Dir As Long)
    MapNpc(index).Dir = Dir

    MapNpc(index).xOffset = 0
    MapNpc(index).yOffset = 0
    MapNpc(index).Moving = 0
End Sub

Public Sub Packet_PlayerXY(ByVal index As Long, ByVal X As Long, ByVal Y As Long)

    Call SetPlayerX(index, X)
    Call SetPlayerY(index, Y)

    ' Make sure they aren't walking
    Player(index).Moving = 0
    Player(index).xOffset = 0
    Player(index).yOffset = 0
End Sub

Public Sub Packet_RemoveMembers()
    Dim n As Long
    For n = 1 To MAX_PARTY_MEMBERS
        Player(MyIndex).Party.Member(n) = 0
    Next
End Sub

Public Sub Packet_UpdateMembers(ByVal i As Long, ByVal index As Long)
    Player(MyIndex).Party.Member(Val(i)) = Val(index)
End Sub

Public Sub Packet_PlayerAttack(ByVal index As Long)
    ' Set player to attacking
    Player(index).Attacking = 1
    Player(index).AttackTimer = GetTickCount
End Sub

Public Sub Packet_NpcAttack(ByVal index As Long)
    ' Set Npc to attacking
    MapNpc(index).Attacking = 1
    MapNpc(index).AttackTimer = GetTickCount
End Sub

Public Sub Packet_EditHouse()
    Call HouseEditorInit
End Sub

