Attribute VB_Name = "modHandleData"
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Created module.
'****************************************************************

Option Explicit

Public Sub HandleData(ByVal Data As String)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Added map constants.
'****************************************************************

Dim Parse() As String
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
Dim i As Long, n As Long, X As Long, Y As Long, p As Integer, f As Integer
Dim ShopNum As Long, GiveItem As Long, GiveValue As Long, GetItem As Long, GetValue As Long, Stock As Long
Dim Weapon As Long, Armor As Long, Helmet As Long, Shield As Long
Dim WeaponDur As Long, ArmorDur As Long, HelmetDur As Long, ShieldDur As Long
    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
    
    ' Add the data to the debug window if we are in debug mode
    If Trim$(Command) = "-debug" Then
        If frmDebug.Visible = False Then frmDebug.Visible = True
        Call TextAdd(frmDebug.txtDebug, "((( Processed Packet " & Parse(0) & " )))", True)
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "alertmsg" Then
        frmSendGetData.Visible = False
        Call LoadMenu
        
        Msg = Parse(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "allchars" Then
        n = 1
        
        frmChars.Visible = True
        frmSendGetData.Visible = False
        
        frmChars.lstChars.Clear
        
        For i = 1 To MAX_CHARS
            Name = Parse(n)
            Msg = Parse(n + 1)
            Level = Val(Parse(n + 2))
            
            If Trim$(Name) = "" Then
                frmChars.lstChars.AddItem "Free Character Slot"
            Else
                frmChars.lstChars.AddItem Name & " a level " & Level & " " & Msg
            End If
            
            n = n + 3
        Next i
        
        frmChars.lstChars.ListIndex = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "loginok" Then
        ' Now we can receive game data
        MyIndex = Val(Parse(1))
        
        frmSendGetData.Visible = True
        frmChars.Visible = False
        
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
        Max_Visible_Classes = Val(Parse(n + 1))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 2
        
        For i = 0 To Max_Classes
            Class(i).Name = Parse(n)
            
            Class(i).HP = Val(Parse(n + 1))
            Class(i).MP = Val(Parse(n + 2))
            Class(i).SP = Val(Parse(n + 3))
            
            Class(i).STR = Val(Parse(n + 4))
            Class(i).DEF = Val(Parse(n + 5))
            Class(i).Speed = Val(Parse(n + 6))
            Class(i).MAGI = Val(Parse(n + 7))
            
            Class(i).Sprite = Val(Parse(n + 8))
            
            n = n + 9
        Next i
        
        ' Used for if the player is creating a new character
        frmNewChar.Visible = True
        frmSendGetData.Visible = False

        frmNewChar.cmbClass.Clear

        For i = 0 To Max_Visible_Classes
            frmNewChar.cmbClass.AddItem Trim$(Class(i).Name)
        Next i
            
        frmNewChar.cmbClass.ListIndex = 0
        frmNewChar.lblHP.Caption = STR(Class(0).HP)
        frmNewChar.lblMP.Caption = STR(Class(0).MP)
        frmNewChar.lblSP.Caption = STR(Class(0).SP)
    
        frmNewChar.lblSTR.Caption = STR(Class(0).STR)
        frmNewChar.lblDEF.Caption = STR(Class(0).DEF)
        frmNewChar.lblSPEED.Caption = STR(Class(0).Speed)
        frmNewChar.lblMAGI.Caption = STR(Class(0).MAGI)
        
        'Set picturebox size
        Call SetPicSize(App.Path + GFX_PATH + "sprites" + GFX_EXT, frmNewChar.picCurrentSprite)
        
        'Load sprite picture
        frmNewChar.picCurrentSprite.Picture = LoadPicture(App.Path & GFX_PATH & "sprites" & GFX_EXT)
        
        'Set the picturebox top and left to show current sprite
        frmNewChar.picCurrentSprite.Left = ((4 * PIC_X) * -1)
        frmNewChar.picCurrentSprite.top = ((Class(frmNewChar.cmbClass.ListIndex).Sprite * PIC_Y) * -1)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "classesdata" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse(n))
        
        ' Redim the Class array
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For i = 0 To Max_Classes
            Class(i).Name = Parse(n)
            
            Class(i).HP = Val(Parse(n + 1))
            Class(i).MP = Val(Parse(n + 2))
            Class(i).SP = Val(Parse(n + 3))
            
            Class(i).STR = Val(Parse(n + 4))
            Class(i).DEF = Val(Parse(n + 5))
            Class(i).Speed = Val(Parse(n + 6))
            Class(i).MAGI = Val(Parse(n + 7))
            
            'Update the data info if it seems to be in use
            If frmSendGetData.Visible = True Then Call SetStatus("Receiving class data!")
            
            n = n + 8
        Next i
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "ingame" Then
        InGame = True
        frmDualSolace.Visible = True
        frmSendGetData.Visible = False
        'Call ResizeGUI
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
        Call DrawInventory
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
        Call DrawInventory
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerworneq" Then
        Call SetPlayerArmorSlot(MyIndex, Val(Parse(1)))
        Call SetPlayerWeaponSlot(MyIndex, Val(Parse(2)))
        Call SetPlayerShieldSlot(MyIndex, Val(Parse(3)))
        Call SetPlayerShieldSlot(MyIndex, Val(Parse(4)))
        Call UpdateInventory
        Call DrawInventory
        Call SendData("GETEQUIPDATA" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playerhp" Then
        Player(MyIndex).MaxHP = Val(Parse(1))
        Call SetPlayerHP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then
            frmDualSolace.lblHPN.Caption = GetPlayerHP(MyIndex) & "/" & GetPlayerMaxHP(MyIndex)
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
            frmDualSolace.lblMPN.Caption = GetPlayerMP(MyIndex) & "/" & GetPlayerMaxMP(MyIndex)
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
            frmDualSolace.lblSPN.Caption = GetPlayerSP(MyIndex) & "/" & GetPlayerMaxSP(MyIndex)
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
        frmDualSolace.lblSTR.Caption = Val(Parse(1))
        frmDualSolace.lblDEF.Caption = Val(Parse(2))
        frmDualSolace.lblSpd.Caption = Val(Parse(3))
        frmDualSolace.lblMAGI.Caption = Val(Parse(4))
        frmDualSolace.lblPName.Caption = Parse(5)
        frmDualSolace.lblLevel.Caption = Val(Parse(6))
        frmDualSolace.lblEXP.Caption = Val(Parse(7))
        frmDualSolace.lblTNL.Caption = Val(Parse(8))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player player packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerplayers" Then
        frmDualSolace.lstPlayers.Clear
        For n = 1 To UBound(Parse)
            frmDualSolace.lstPlayers.AddItem Parse(n)
        Next n
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Player friends packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerfriends" Then
        For n = 1 To MAX_FRIENDS
            Player(MyIndex).Friends(n) = Parse(n)
            If IsPlaying(FindPlayer(Parse(n))) Then
                frmDualSolace.lstFriends.AddItem Parse(n)
            End If
        Next n
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
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playermove") Then
        i = Val(Parse(1))
        X = Val(Parse(2))
        Y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        Call SetPlayerX(i, X)
        Call SetPlayerY(i, Y)
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
        X = Val(Parse(2))
        Y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        MapNpc(i).X = X
        MapNpc(i).Y = Y
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
        i = Val(Parse(1))
        
        ' Set player to attacking
        Player(i).Attacking = 1
        Player(i).AttackTimer = GetTickCount
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
        
        If FileExist(MAP_PATH & "map" & X & MAP_EXT, True) Then
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
        If frmSendGetData.Visible = True Then Call SetStatus("Receiving map data...")
        DoEvents
        n = 1
        
        SaveMap.Name = Parse(n + 1)
        SaveMap.Revision = Val(Parse(n + 2))
        SaveMap.Moral = Val(Parse(n + 3))
        SaveMap.Up = Val(Parse(n + 4))
        SaveMap.Down = Val(Parse(n + 5))
        SaveMap.Left = Val(Parse(n + 6))
        SaveMap.Right = Val(Parse(n + 7))
        SaveMap.Music = Parse(n + 8)
        SaveMap.BootMap = Val(Parse(n + 9))
        SaveMap.BootX = Val(Parse(n + 10))
        SaveMap.BootY = Val(Parse(n + 11))
        SaveMap.Shop = Val(Parse(n + 12))
        
        n = n + 13
        
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                SaveMap.Tile(X, Y).Ground = Val(Parse(n))
                SaveMap.Tile(X, Y).Mask = Val(Parse(n + 1))
                SaveMap.Tile(X, Y).Mask2 = Val(Parse(n + 2))
                SaveMap.Tile(X, Y).Anim = Val(Parse(n + 3))
                SaveMap.Tile(X, Y).Anim2 = Val(Parse(n + 4))
                SaveMap.Tile(X, Y).Fringe = Val(Parse(n + 5))
                SaveMap.Tile(X, Y).FringeAnim = Val(Parse(n + 6))
                SaveMap.Tile(X, Y).Fringe2 = Val(Parse(n + 7))
                n = n + 8
            Next X
        Next Y
        
        For X = 1 To MAX_MAP_NPCS
            SaveMap.Npc(X) = Val(Parse(n))
            n = n + 1
        Next X
                
        ' Save the map
        Call SaveLocalMap(Val(Parse(1)))
        
        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmDualSolace.picMapEditor.Visible = False
            
            If frmMapWarp.Visible Then
                Unload frmMapWarp
            End If
            
            If frmMapProperties.Visible Then
                Unload frmMapProperties
            End If
        End If
        
        Exit Sub
    End If
    
    If LCase$(Parse(0)) = "mapattributes" Then
        If frmSendGetData.Visible = True Then Call SetStatus("Receiving Map Attributes...")
        DoEvents
        
        n = 2
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                ' I've received errors that make walkable false.
                ' This will check if walkable is false, if so,
                ' it will make sure blocked is true or all side
                ' blocks are true.
                If CBool(Val(Parse(n))) = True Then
                    SaveMap.Tile(X, Y).Walkable = Val(Parse(n))
                ElseIf CBool(Val(Parse(n))) = False Then
                    If CBool(Val(Parse(n + 1))) = True Then
                        SaveMap.Tile(X, Y).Walkable = Val(Parse(n))
                    Else
                        If CBool(Val(Parse(n + 16))) = True Then
                            If CBool(Val(Parse(n + 17))) = True Then
                                If CBool(Val(Parse(n + 18))) = True Then
                                    If CBool(Val(Parse(n + 19))) = True Then
                                        SaveMap.Tile(X, Y).Walkable = Val(Parse(n))
                                    Else
                                        ' There was an error, set it right
                                        SaveMap.Tile(X, Y).Walkable = True
                                    End If
                                Else
                                    ' There was an error, set it right
                                    SaveMap.Tile(X, Y).Walkable = True
                                End If
                            Else
                                ' There was an error, set it right
                                SaveMap.Tile(X, Y).Walkable = True
                            End If
                        Else
                            ' There was an error, set it right
                            SaveMap.Tile(X, Y).Walkable = True
                        End If
                    End If
                End If
                SaveMap.Tile(X, Y).Blocked = Val(Parse(n + 1))
                SaveMap.Tile(X, Y).Warp = Val(Parse(n + 2))
                SaveMap.Tile(X, Y).WarpMap = Val(Parse(n + 3))
                SaveMap.Tile(X, Y).WarpX = Val(Parse(n + 4))
                SaveMap.Tile(X, Y).WarpY = Val(Parse(n + 5))
                SaveMap.Tile(X, Y).Item = Val(Parse(n + 6))
                SaveMap.Tile(X, Y).ItemNum = Val(Parse(n + 7))
                SaveMap.Tile(X, Y).ItemValue = Val(Parse(n + 8))
                SaveMap.Tile(X, Y).NpcAvoid = Val(Parse(n + 9))
                SaveMap.Tile(X, Y).Key = Val(Parse(n + 10))
                SaveMap.Tile(X, Y).KeyNum = Val(Parse(n + 11))
                SaveMap.Tile(X, Y).KeyTake = Val(Parse(n + 12))
                SaveMap.Tile(X, Y).KeyOpen = Val(Parse(n + 13))
                SaveMap.Tile(X, Y).KeyOpenX = Val(Parse(n + 14))
                SaveMap.Tile(X, Y).KeyOpenY = Val(Parse(n + 15))
                SaveMap.Tile(X, Y).North = Val(Parse(n + 16))
                SaveMap.Tile(X, Y).West = Val(Parse(n + 17))
                SaveMap.Tile(X, Y).East = Val(Parse(n + 18))
                SaveMap.Tile(X, Y).South = Val(Parse(n + 19))
                SaveMap.Tile(X, Y).Shop = Val(Parse(n + 20))
                SaveMap.Tile(X, Y).ShopNum = Val(Parse(n + 21))
                SaveMap.Tile(X, Y).Bank = Val(Parse(n + 22))
                SaveMap.Tile(X, Y).Heal = Val(Parse(n + 23))
                SaveMap.Tile(X, Y).HealValue = Val(Parse(n + 24))
                SaveMap.Tile(X, Y).Damage = Val(Parse(n + 25))
                SaveMap.Tile(X, Y).DamageValue = Val(Parse(n + 26))
                n = n + 27
            Next X
        Next Y
        
        Call SaveLocalMap(Val(Parse(1)))
    End If
        
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapitemdata" Then
        n = 1
        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).num = Val(Parse(n))
            SaveMapItem(i).Value = Val(Parse(n + 1))
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
    If LCase(Parse(0)) = "mapnpcdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_NPCS
            SaveMapNpc(i).num = Val(Parse(n))
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
    If LCase(Parse(0)) = "mapdone" Then
        Map = SaveMap
        
        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i
        
        GettingMap = False
        
        ' Play music
        Call DirectMusic.StopMusic
        If Map.Music <> "No Music" And Map.Music <> "" Then
            'See if the client should play music
            If frmGameSettings.chkMusic.Value = 1 Then
                Call DirectMusic.PlayMusic(Map.Music & ".mid")
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
        n = Val(Parse(1))
        
        MapItem(n).num = Val(Parse(2))
        MapItem(n).Value = Val(Parse(3))
        MapItem(n).Dur = Val(Parse(4))
        MapItem(n).X = Val(Parse(5))
        MapItem(n).Y = Val(Parse(6))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "itemeditor") Then

        frmEditor.lstItemEditor.Clear
        
        ' Add the names
        For i = 1 To MAX_ITEMS
            frmEditor.lstItemEditor.AddItem i & ": " & Trim$(Item(i).Name)
        Next i
        
        frmEditor.lstItemEditor.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateitem") Then
        n = Val(Parse(1))
        
        ' Update the item
        Item(n).Name = Parse(2)
        Item(n).Description = Parse(3)
        Item(n).Pic = Val(Parse(4))
        Item(n).Type = Val(Parse(5))
        Item(n).Data1 = Val(Parse(6))
        Item(n).Data2 = Val(Parse(7))
        Item(n).Data3 = Val(Parse(8))
        Item(n).Data4 = Val(Parse(9))
        Item(n).Data5 = Val(Parse(10))
        Item(n).Sound = Trim$(Parse(11))
        
        'Update the data info if it seems to be in use
        If frmSendGetData.Visible = True Then Call SetStatus("Receiving item data!")
        DoEvents
        Exit Sub
    End If
       
    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "edititem") Then
        n = Val(Parse(1))
        
        ' Update the item
        Item(n).Name = Parse(2)
        Item(n).Description = Parse(3)
        Item(n).Pic = Val(Parse(4))
        Item(n).Type = Val(Parse(5))
        Item(n).Data1 = Val(Parse(6))
        Item(n).Data2 = Val(Parse(7))
        Item(n).Data3 = Val(Parse(8))
        Item(n).Data4 = Val(Parse(9))
        Item(n).Data5 = Val(Parse(10))
        Item(n).Sound = Trim$(Parse(11))
        
        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnnpc" Then
        n = Val(Parse(1))
        
        MapNpc(n).num = Val(Parse(2))
        MapNpc(n).X = Val(Parse(3))
        MapNpc(n).Y = Val(Parse(4))
        MapNpc(n).Dir = Val(Parse(5))
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        
        'Update NPC's HP
        MapNpc(n).HP = MapNpc(n).MaxHP
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "npcdead" Then
        n = Val(Parse(1))
        
        MapNpc(n).Death.DeathX = MapNpc(n).X
        MapNpc(n).Death.DeathY = MapNpc(n).Y
        MapNpc(n).Death.DeathDir = MapNpc(n).Dir
        
        MapNpc(n).num = 0
        MapNpc(n).X = 0
        MapNpc(n).Y = 0
        MapNpc(n).Dir = 0
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        
        'Stops the bar from being blitted when the npc is dead
        MapNpc(n).HP = 0
        
        ' Set the map npc death animation
        MapNpc(n).Death.Display = True
        MapNpc(n).Death.AnimCount = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "npceditor") Then

        frmEditor.lstNpcEditor.Clear
        
        ' Add the names
        For i = 1 To MAX_NPCS
            frmEditor.lstNpcEditor.AddItem i & ": " & Trim$(Npc(i).Name)
        Next i
        
        frmEditor.lstNpcEditor.ListIndex = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatenpc") Then
        n = Val(Parse(1))
        
        ' Update the item
        Npc(n).Name = Parse(2)
        Npc(n).AttackSay = ""
        Npc(n).Sprite = Val(Parse(3))
        Npc(n).SpawnSecs = 0
        Npc(n).Behavior = 0
        Npc(n).Range = 0
        Npc(n).DropChance = 0
        Npc(n).DropItem = 0
        Npc(n).DropItemValue = 0
        Npc(n).STR = 0
        Npc(n).DEF = 0
        Npc(n).Speed = 0
        Npc(n).MAGI = 0
        
        'Update the data info if it seems to be in use
        If frmSendGetData.Visible = True Then Call SetStatus("Receiving npc data!")
        
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet :: <- Used for item editor admins only
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "editnpc") Then
        n = Val(Parse(1))
        
        ' Update the npc
        Npc(n).Name = Parse(2)
        Npc(n).AttackSay = Parse(3)
        Npc(n).Sprite = Val(Parse(4))
        Npc(n).SpawnSecs = Val(Parse(5))
        Npc(n).Behavior = Val(Parse(6))
        Npc(n).Range = Val(Parse(7))
        Npc(n).DropChance = Val(Parse(8))
        Npc(n).DropItem = Val(Parse(9))
        Npc(n).DropItemValue = Val(Parse(10))
        Npc(n).HP = CLng(Parse(11))
        Npc(n).STR = Val(Parse(12))
        Npc(n).DEF = Val(Parse(13))
        Npc(n).Speed = Val(Parse(14))
        Npc(n).MAGI = Val(Parse(15))
        Npc(n).EXP = Val(Parse(16))
        Npc(n).Fear = CBool(Parse(17))
        
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
        n = Val(Parse(3))
        
        TempTile(X, Y).DoorOpen = n
        
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

        frmEditor.lstShopEditor.Clear
        
        ' Add the names
        For i = 1 To MAX_SHOPS
            frmEditor.lstShopEditor.AddItem i & ": " & Trim$(Shop(i).Name)
        Next i
        
        frmEditor.lstShopEditor.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateshop") Then
        n = Val(Parse(1))
        
        ' Update the shop name
        Shop(n).Name = Parse(2)
        
        'Update the data info if it seems to be in use
        If frmSendGetData.Visible = True Then Call SetStatus("Receiving shop data!")
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
        Shop(ShopNum).Restock = Val(Parse(6))
        
        n = 7
        For i = 1 To MAX_TRADES
            
            GiveItem = Val(Parse(n))
            GiveValue = Val(Parse(n + 1))
            GetItem = Val(Parse(n + 2))
            GetValue = Val(Parse(n + 3))
            Stock = Val(Parse(n + 4))
            
            Shop(ShopNum).TradeItem(i).GiveItem = GiveItem
            Shop(ShopNum).TradeItem(i).GiveValue = GiveValue
            Shop(ShopNum).TradeItem(i).GetItem = GetItem
            Shop(ShopNum).TradeItem(i).GetValue = GetValue
            Shop(ShopNum).TradeItem(i).MaxStock = Stock
            
            n = n + 5
        Next i
        
        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "spelleditor") Then
        
        frmEditor.lstSpellEditor.Clear
        
        ' Add the names
        For i = 1 To MAX_SPELLS
            frmEditor.lstSpellEditor.AddItem i & ": " & Trim$(Spell(i).Name)
        Next i
        
        frmEditor.lstSpellEditor.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatespell") Then
        n = Val(Parse(1))
        
        ' Update the spell name
        Spell(n).Name = Parse(2)
        
        'Update the data info if it seems to be in use
        If frmSendGetData.Visible = True Then Call SetStatus("Receiving spell data!")
        
        'Update the data info if it seems to be in use
        '-delete if want to be more specific
        If frmSendGetData.Visible = True Then Call SetStatus("Receiving user data!")
            
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "editspell") Then
        n = Val(Parse(1))
        
        ' Update the spell
        Spell(n).Name = Parse(2)
        Spell(n).ClassReq = Val(Parse(3))
        Spell(n).LevelReq = Val(Parse(4))
        Spell(n).Type = Val(Parse(5))
        Spell(n).Data1 = Val(Parse(6))
        Spell(n).Data2 = Val(Parse(7))
        Spell(n).Data3 = Val(Parse(8))
                        
        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Class editor packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "classeditor") Then
        
        frmEditor.lstClassEditor.Clear
        
        ' Add the names
        For i = 0 To Max_Classes
            frmEditor.lstClassEditor.AddItem i & ": " & Trim$(Class(i).Name)
        Next i
        
        frmEditor.lstClassEditor.ListIndex = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Update class packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateclass") Then
        n = Val(Parse(1))
        
        ' Update the class name
        Class(n).Name = Parse(2)
        Class(n).Sprite = Val(Parse(3))
        Class(n).HP = Val(Parse(4))
        Class(n).MP = Val(Parse(5))
        Class(n).SP = Val(Parse(6))
        Class(n).STR = Val(Parse(7))
        Class(n).DEF = Val(Parse(8))
        Class(n).MAGI = Val(Parse(9))
        Class(n).Speed = Val(Parse(10))
        Class(n).Map = Val(Parse(11))
        Class(n).X = Val(Parse(12))
        Class(n).Y = Val(Parse(13))
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit class packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "editclass") Then
        n = Val(Parse(1))
        
        ' Update the Class
        Class(n).Name = Parse(2)
        Class(n).Sprite = Val(Parse(3))
        Class(n).HP = Val(Parse(4))
        Class(n).MP = Val(Parse(5))
        Class(n).SP = Val(Parse(6))
        Class(n).STR = Val(Parse(7))
        Class(n).DEF = Val(Parse(8))
        Class(n).MAGI = Val(Parse(9))
        Class(n).Speed = Val(Parse(10))
        Class(n).Map = Val(Parse(11))
        Class(n).X = Val(Parse(12))
        Class(n).Y = Val(Parse(13))
                        
        Debug.Print "Class edit received!"
                        
        ' Initialize the Class editor
        Call ClassEditorInit

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
        
        frmTrade.lblRestock.Caption = Parse(3)
        
        n = 4
        For i = 1 To MAX_TRADES
            GiveItem = Val(Parse(n))
            GiveValue = Val(Parse(n + 1))
            GetItem = Val(Parse(n + 2))
            GetValue = Val(Parse(n + 3))
            Stock = Val(Parse(n + 4))
            
            'Set the values
            Shop(ShopNum).TradeItem(i).GiveItem = GiveItem
            Shop(ShopNum).TradeItem(i).GetItem = GetItem
            
            If GiveItem > 0 And GetItem > 0 And Stock > -1 Then
                frmTrade.lstTrade.AddItem (Trim$(Item(GetItem).Name) & "(" & Stock & ")")
            ElseIf GiveItem > 0 And GetItem > 0 And Stock = -1 Then
                frmTrade.lstTrade.AddItem (Trim$(Item(GetItem).Name) & "(Infinite)")
            End If
            n = n + 5
        Next i
        
        If frmTrade.lstTrade.ListCount > 0 Then
            frmTrade.lstTrade.ListIndex = 0
        End If
            If Shop(ShopNum).TradeItem(1).GetItem > 0 Then
                frmTrade.lblName.Caption = Trim$(Item(Shop(ShopNum).TradeItem(1).GetItem).Name)
                frmTrade.lblDescription.Caption = Trim$(Item(Shop(ShopNum).TradeItem(1).GetItem).Description)
            End If
            If Shop(ShopNum).TradeItem(1).GiveItem > 0 Then
                frmTrade.lblCost.Caption = Trim$(Item(Shop(ShopNum).TradeItem(1).GiveItem).Name) & ": " & CStr(Shop(ShopNum).TradeItem(1).GiveValue)
            End If
                frmTrade.lblStock.Caption = CStr(Shop(ShopNum).TradeItem(1).Stock)
                
        frmTrade.Show vbModal
        Exit Sub
    End If
    
    'Get Item Data in shop
    '-smchronos
    If (LCase(Parse(0)) = "tradegetitem") Then
        ShopNum = Val(Parse(1))
        frmTrade.lblName.Caption = Trim$(Parse(2))
        frmTrade.lblDescription.Caption = Trim$(Parse(3))
        frmTrade.lblCost.Caption = Trim$(Parse(4)) & ": " & Parse(5)
        
        If Val(Parse(6)) = -1 Then
            frmTrade.lblStock.Caption = "Inifinite"
        Else
            frmTrade.lblStock.Caption = Parse(6)
        End If
        
        'Not using any of these...clean later
        '-smchronos
        'If GetPlayerWeaponSlot(MyIndex) > 0 Then Weapon = Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).Data2
        'If GetPlayerArmorSlot(MyIndex) > 0 Then Armor = Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).Data2
        'If GetPlayershieldSlot(MyIndex) > 0 Then shield = Item(GetPlayerInvItemNum(MyIndex, GetPlayershieldSlot(MyIndex))).Data2
        'If GetPlayerShieldSlot(MyIndex) > 0 Then Shield = Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).Data2
        'If GetPlayerWeaponSlot(MyIndex) > 0 Then Weapon = WornItems.Weapon.Strength
        'If GetPlayerArmorSlot(MyIndex) > 0 Then Armor = WornItems.Armor.Strength
        'If GetPlayershieldSlot(MyIndex) > 0 Then shield = WornItems.shield.Strength
        'If GetPlayerShieldSlot(MyIndex) > 0 Then Shield = WornItems.Shield.Strength

        'Set the Variables
        Weapon = Val(Parse(9))
        Armor = Val(Parse(10))
        Shield = Val(Parse(11))
        Shield = Val(Parse(12))
        
        'Set high labels
        frmTrade.lblHSTR.Caption = CStr(GetPlayerSTR(MyIndex))
        frmTrade.lblHDEF.Caption = CStr(GetPlayerDEF(MyIndex))
        frmTrade.lblHMAG.Caption = CStr(GetPlayerMAGI(MyIndex))
        frmTrade.lblHSPD.Caption = CStr(GetPlayerSPEED(MyIndex))
        
        'Set the original labal colors for the new stats
        frmTrade.lblHSTR.ForeColor = &H0&
        frmTrade.lblHDEF.ForeColor = &H0&
        frmTrade.lblHMAG.ForeColor = &H0&
        frmTrade.lblHSPD.ForeColor = &H0&
        
        If frmTrade.lstTrade.ListIndex < 0 Then frmTrade.lstTrade.ListIndex = 0
        
        If Shop(ShopNum).TradeItem(frmTrade.lstTrade.ListIndex + 1).GetItem < 1 Then Exit Sub
        If Shop(ShopNum).TradeItem(frmTrade.lstTrade.ListIndex + 1).GiveItem < 1 Then Exit Sub
        
        If CInt(Parse(7)) = 1 Then
            If CLng(Parse(8)) > Weapon Then
                frmTrade.lblHSTR.ForeColor = &HC000&
                frmTrade.lblHSTR.Caption = CStr(CLng(Parse(8)) + GetPlayerSTR(MyIndex))
            ElseIf CLng(Parse(8)) < Weapon Then
                frmTrade.lblHSTR.ForeColor = &HC0&
                frmTrade.lblHSTR.Caption = CStr(CLng(Parse(8)) + GetPlayerSTR(MyIndex))
            End If
        ElseIf CInt(Parse(7)) = 2 Then
            If CLng(Parse(8)) > Armor Then
                frmTrade.lblHDEF.ForeColor = &HC000&
                frmTrade.lblHDEF.Caption = CStr(CLng(Parse(8)) + GetPlayerDEF(MyIndex))
            ElseIf CLng(Parse(8)) < Armor Then
                frmTrade.lblHDEF.ForeColor = &HC0&
                frmTrade.lblHDEF.Caption = CStr(CLng(Parse(8)) + GetPlayerDEF(MyIndex))
            End If
        ElseIf CInt(Parse(7)) = 3 Then
            If CLng(Parse(8)) > Shield Then
                frmTrade.lblHDEF.ForeColor = &HC000&
                frmTrade.lblHDEF.Caption = CStr(CLng(Parse(8)) + GetPlayerDEF(MyIndex))
            ElseIf CLng(Parse(8)) < Shield Then
                frmTrade.lblHDEF.ForeColor = &HC0&
                frmTrade.lblHDEF.Caption = CStr(CLng(Parse(8)) + GetPlayerDEF(MyIndex))
            End If
        ElseIf CInt(Parse(7)) = 4 Then
            If CLng(Parse(8)) > Shield Then
                frmTrade.lblHDEF.ForeColor = &HC000&
                frmTrade.lblHDEF.Caption = CStr(CLng(Parse(8)) + GetPlayerDEF(MyIndex))
            ElseIf CLng(Parse(8)) < Shield Then
                frmTrade.lblHDEF.ForeColor = &HC0&
                frmTrade.lblHDEF.Caption = CStr(CLng(Parse(8)) + GetPlayerDEF(MyIndex))
            End If
        End If
    End If
    
    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (LCase(Parse(0)) = "spells") Then
        
        frmDualSolace.picPlayerSpells.Visible = True
        frmDualSolace.lstSpells.Clear
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(i) = Val(Parse(i))
            If Player(MyIndex).Spell(i) <> 0 Then
                frmDualSolace.lstSpells.AddItem i & ": " & Trim$(Spell(Player(MyIndex).Spell(i)).Name)
            Else
                frmDualSolace.lstSpells.AddItem "<free spells slot>"
            End If
        Next i
        
        frmDualSolace.lstSpells.ListIndex = 0
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
    End If
    
    ' :::::::::::::::::
    ' :: Exp. packet :: -smchronos
    ' :::::::::::::::::
    If (LCase(Parse(0)) = "experience") Then
        frmDualSolace.lblEXP.Caption = Val(Parse(1))
        frmDualSolace.lblTNL.Caption = Val(Parse(2))
    End If
    
    ' :::::::::::::::::
    ' :: Stat packet :: -smchronos
    ' :::::::::::::::::
    If (LCase(Parse(0)) = "pstats") Then
        frmDualSolace.lblSTR.Caption = Val(Parse(1))
        frmDualSolace.lblDEF.Caption = Val(Parse(2))
        frmDualSolace.lblMAGI.Caption = Val(Parse(3))
        frmDualSolace.lblSpd.Caption = Val(Parse(4))
        frmDualSolace.lblPName.Caption = Parse(5)
        frmDualSolace.lblLevel.Caption = Val(Parse(6))
        frmDualSolace.lblEXP.Caption = Val(Parse(7))
        frmDualSolace.lblTNL.Caption = Val(Parse(8))
    End If
    
    ' ::::::::::::::::::::::
    ' :: OnLevelUp packet :: -smchronos
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "levelup") Then
        frmDualSolace.lblSTR.Caption = Val(Parse(1))
        frmDualSolace.lblDEF.Caption = Val(Parse(2))
        frmDualSolace.lblMAGI.Caption = Val(Parse(3))
        frmDualSolace.lblSpd.Caption = Val(Parse(4))
        frmDualSolace.lblLevel.Caption = Val(Parse(5))
        frmDualSolace.lblEXP.Caption = Val(Parse(6))
        frmDualSolace.lblTNL.Caption = Val(Parse(7))
    End If
    
    ' :::::::::::::::::::::::
    ' :: PlayerJoin packet :: -smchronos
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerjoin") Then
        frmDualSolace.lstPlayers.AddItem Parse(1)
        If IsFriend(MyIndex, Parse(1)) Then
            frmDualSolace.lstFriends.AddItem Parse(1)
            If frmDualSolace.lblFNumber.Caption <> "NUM" Then
                frmDualSolace.lblFNumber.Caption = CStr(CInt(frmDualSolace.lblFNumber.Caption) + 1)
            Else
                frmDualSolace.lblFNumber.Caption = "0"
            End If
        End If
        frmDualSolace.lblPNumber.Caption = CStr(CInt(frmDualSolace.lblPNumber.Caption) + 1)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: PlayerLeave packet :: -smchronos
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerleave") Then
        n = Parse(1)
        frmDualSolace.lblPNumber.Caption = n
        f = 0
        
        If n = "0" Then
            Exit Sub
        End If
        
        frmDualSolace.lstPlayers.Clear
        frmDualSolace.lstFriends.Clear
        
        For p = 2 To (n + 1)
            frmDualSolace.lstPlayers.AddItem Parse(p)
            If IsFriend(MyIndex, Parse(p)) Then
                frmDualSolace.lstFriends.AddItem Parse(p)
                f = f + 1
            End If
        Next p
        frmDualSolace.lblFNumber.Caption = CStr(f)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: PlayerList packet :: -smchronos
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "plist") Then
        n = Parse(1)
        frmDualSolace.lblPNumber.Caption = n
        frmDualSolace.lblFNumber.Caption = "0"
        f = 0
        
        If n = "0" Then
            Exit Sub
        End If
        
        For p = 2 To (n + 1)
            frmDualSolace.lstPlayers.AddItem Parse(p)
            If IsFriend(MyIndex, Parse(p)) Then
                frmDualSolace.lstFriends.AddItem Parse(p)
                f = f + 1
            End If
        Next p
        frmDualSolace.lblFNumber.Caption = CStr(f)
        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Friend packets :: -smchronos
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "addfriend") Then
        frmDualSolace.lstFriends.AddItem Parse(1)
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "remfriend") Then
        For p = 0 To frmDualSolace.lstFriends.ListCount
            If frmDualSolace.lstFriends.List(p) = Parse(1) Then
                frmDualSolace.lstFriends.RemoveItem (p)
                Exit Sub
            End If
        Next p
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Get Info Packet :: -smchronos
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "serverinfo") Then
        frmMainMenu.txtInfo.Text = Parse(1)
        frmDualSolace.Socket.Close
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' ::  Get MAX DATA  :: -smchronos
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "maxdata") Then
    GAME_NAME = Parse(1)
    MAX_PLAYERS = Val(Parse(2))
    MAX_MAPS = Val(Parse(3))
    MAX_ITEMS = Val(Parse(4))
    MAX_NPCS = Val(Parse(5))
    MAX_SHOPS = Val(Parse(6))
    MAX_SPELLS = Val(Parse(7))
    
    ReDim Player(1 To MAX_PLAYERS) As PlayerRec
    ReDim Item(1 To MAX_ITEMS) As ItemRec
    ReDim Npc(1 To MAX_NPCS) As NpcRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    frmMapItem.scrlItem.Max = MAX_ITEMS
    frmMapKey.scrlItem.Max = MAX_ITEMS
    Call SetName(GAME_NAME)
    Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' ::  Get MAX DATA  :: -smchronos
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "equipdata" Then
    If InGame = True Then
        WeaponDur = 0
        ArmorDur = 0
        HelmetDur = 0
        ShieldDur = 0
        Weapon = Val(Parse(1))
        Armor = Val(Parse(3))
        Helmet = Val(Parse(5))
        Shield = Val(Parse(7))
        If Weapon > 0 Then WeaponDur = Val(Parse(2))
        If Armor > 0 Then ArmorDur = Val(Parse(4))
        If Helmet > 0 Then HelmetDur = Val(Parse(6))
        If Shield > 0 Then ShieldDur = Val(Parse(8))
        Call DrawEquipment(Weapon, WeaponDur, Armor, ArmorDur, Helmet, HelmetDur, Shield, ShieldDur)
    End If
    Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' ::    BANK DATA   :: -smchronos
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "playerbank" Then
        n = 1
        For i = 1 To MAX_BANK_ITEMS
            Call SetPlayerBankItemNum(MyIndex, i, Val(Parse(n)))
            Call SetPlayerBankItemValue(MyIndex, i, Val(Parse(n + 1)))
            Call SetPlayerBankItemDur(MyIndex, i, Val(Parse(n + 2)))
            n = n + 3
            'If i >= 1 And i <= 20 And GetPlayerBankItemNum(MyIndex, i) <> 0 Then frmDualSolace.lblBValue(i - 1).Caption = CStr(GetPlayerBankItemValue(MyIndex, i))
        Next i
        'set the bank default settings
        frmDualSolace.txtDeposit.Text = "1"
        frmDualSolace.txtWithdraw.Text = "1"
        InvSelected = 1
        BankSelected = 1
        frmDualSolace.lblItemName.Caption = Trim$(Item(GetPlayerBankItemNum(MyIndex, 1)).Name)
        If Item(GetPlayerBankItemNum(MyIndex, 1)).Type = ITEM_TYPE_CURRENCY Then
            frmDualSolace.lblItemData.Caption = GetPlayerBankItemValue(MyIndex, 1)
        Else
            frmDualSolace.lblItemData.Caption = "1"
        End If
        frmDualSolace.picBank.Visible = True
        Call DrawInventory
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update bank packet :: -smchronos
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatebankitem") Then
        n = CByte(Parse(1))
        Call SetPlayerBankItemNum(MyIndex, n, Val(Parse(2)))
        Call SetPlayerBankItemValue(MyIndex, n, Val(Parse(3)))
        Call SetPlayerBankItemDur(MyIndex, n, Val(Parse(4)))
        Call DrawInventory
        'If frmDualSolace.lblBValue(n - 1).Caption = "0" Then frmDualSolace.lblBValue(n - 1).Caption = ""
        'Debug.Print "ItemValue: " & ItemVal & " - BankItemValue: " & GetPlayerBankItemValue(Index, i)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::: This is only here to fix an odd bug.
    ' :: Update bank packet :: Items that should be deleted aren't.
    ' :::::::::::::::::::::::: -smchronos
    If (LCase(Parse(0)) = "killdata") Then
        Select Case Trim$(LCase$(Parse(1)))
        Case "bank":
        n = Val(Parse(2))
            Call SetPlayerBankItemNum(MyIndex, n, 0)
            Call SetPlayerBankItemValue(MyIndex, n, 0)
            Call SetPlayerBankItemDur(MyIndex, n, 0)
        
        Case "item":
        n = Val(Parse(2))
            Call SetPlayerInvItemNum(MyIndex, n, 0)
            Call SetPlayerInvItemValue(MyIndex, n, 0)
            Call SetPlayerInvItemDur(MyIndex, n, 0)
        
        End Select
    End If
    
    ' ::::::::::::::::::::
    ' :: Tracker packet ::
    ' ::::::::::::::::::::
    If LCase$(Parse(0)) = "trackerupdate" Then
        If frmTrack.Visible = True Then
        frmTrack.SetFocus
            Select Case Parse(1)
            Case "MAP":
                Call TextAdd(frmTrack.txtMapChat, Parse(2), True)
                Call TextAdd(frmTrack.txtFullChat, "(Regular) " & Parse(2), True)
                frmTrack.sbChat.SimpleText = "New map message!"
            Case "EMOTE":
                Call TextAdd(frmTrack.txtMapChat, Parse(2), True)
                Call TextAdd(frmTrack.txtFullChat, "(Emote) " & Parse(2), True)
                frmTrack.sbChat.SimpleText = "New emote message!"
            Case "BROADCAST":
                Call TextAdd(frmTrack.txtBroadcastChat, Parse(2), True)
                Call TextAdd(frmTrack.txtFullChat, "(Broadcast) " & Parse(2), True)
                frmTrack.sbChat.SimpleText = "New broadcast message!"
            Case "GLOBAL":
                Call TextAdd(frmTrack.txtGlobalChat, Parse(2), True)
                Call TextAdd(frmTrack.txtFullChat, "(Global) " & Parse(2), True)
                frmTrack.sbChat.SimpleText = "New global message!"
            Case "PRIVATE":
                Call TextAdd(frmTrack.txtPrivateChat, Parse(2), True)
                Call TextAdd(frmTrack.txtFullChat, "(Private) " & Parse(2), True)
                frmTrack.sbChat.SimpleText = "New private message!"
            Case "ADMIN":
                Call TextAdd(frmTrack.txtAdminChat, Parse(2), True)
                Call TextAdd(frmTrack.txtFullChat, "(Admin) " & Parse(2), True)
                frmTrack.sbChat.SimpleText = "New admin message!"
            Case "TRACKER":
                Call TextAdd(frmTrack.txtTrackerChat, Parse(2), True)
                Call TextAdd(frmTrack.txtFullChat, "(Tracker) " & Parse(2), True)
                frmTrack.sbChat.SimpleText = "New tracker message!"
            End Select
        End If
    End If
    
    'admin panel
    If LCase(Parse(0)) = "adminpanel" Then
        If frmDualSolace.picAdminPanel.Visible = True Then
            frmDualSolace.picAdminPanel.Visible = False
            frmDualSolace.cmdBan.Enabled = False
            frmDualSolace.cmdKick.Enabled = False
            frmDualSolace.cmdListen.Enabled = False
            frmDualSolace.cmdTrack.Enabled = False
            frmDualSolace.cmdMapEditor.Enabled = False
            frmDualSolace.cmdWarpTo.Enabled = False
            frmDualSolace.cmdWarpMeTo.Enabled = False
            frmDualSolace.cmdWarpToMe.Enabled = False
            frmDualSolace.lstMapSelection.Clear
            frmDualSolace.cmdMapWarp.Enabled = False
            frmDualSolace.cmdEditor.Enabled = False
            frmDualSolace.cmdSetAccess.Enabled = False
            frmDualSolace.txtAccessName.Enabled = False
            frmDualSolace.cmbAccess.Enabled = False
            frmDualSolace.cmbAccess.Clear
            Exit Sub
        End If
    
        'set mod options
        If Val(Parse(1)) >= 1 Then
            frmDualSolace.cmdBan.Enabled = True
            frmDualSolace.cmdKick.Enabled = True
            frmDualSolace.cmdListen.Enabled = True
            frmDualSolace.cmdTrack.Enabled = True
        End If
        
        'set mapper options
        If Val(Parse(1)) >= 2 Then
            frmDualSolace.cmdMapEditor.Enabled = True
            frmDualSolace.cmdWarpTo.Enabled = True
            frmDualSolace.cmdWarpMeTo.Enabled = True
            frmDualSolace.cmdWarpToMe.Enabled = True
            For n = 1 To MAX_MAPS
                Call LoadMap(n)
                If Trim$(SaveMap.Name) <> "" Then
                    frmDualSolace.lstMapSelection.AddItem n & ". " & Trim$(SaveMap.Name)
                Else
                    frmDualSolace.lstMapSelection.AddItem n & "."
                End If
            Next n
            Call LoadMap(GetPlayerMap(MyIndex))
            frmDualSolace.cmdMapWarp.Enabled = True
        End If
        
        'set developer options
        If Val(Parse(1)) >= 3 Then
            frmDualSolace.cmdEditor.Enabled = True
        End If
        
        'set administrator options
        If Val(Parse(1)) >= 4 Then
            frmDualSolace.cmdSetAccess.Enabled = True
            frmDualSolace.txtAccessName.Enabled = True
            frmDualSolace.cmbAccess.Enabled = True
            For n = 0 To 9
                frmDualSolace.cmbAccess.AddItem n
            Next n
            frmDualSolace.cmbAccess.ListIndex = 0
        Else
            For n = 0 To 9
                frmDualSolace.cmbAccess.AddItem n
            Next n
            frmDualSolace.cmbAccess.ListIndex = 0
        End If
        
        frmDualSolace.picAdminPanel.Visible = True
    End If
    
    If LCase$(Parse(0)) = "classdata" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse(n))
        Max_Visible_Classes = Val(Parse(n + 1))
        
        n = n + 2
        
        For i = 0 To Max_Classes
            Class(i).Name = Parse(n)
            
            Class(i).HP = Val(Parse(n + 1))
            Class(i).MP = Val(Parse(n + 2))
            Class(i).SP = Val(Parse(n + 3))
            
            Class(i).STR = Val(Parse(n + 4))
            Class(i).DEF = Val(Parse(n + 5))
            Class(i).Speed = Val(Parse(n + 6))
            Class(i).MAGI = Val(Parse(n + 7))
            
            Class(i).Sprite = Val(Parse(n + 8))
            
            n = n + 9
        Next i
    End If
    
    ' The play sound command
    If LCase$(Parse(0)) = "playsound" Then
        'Play the sound
        If LCase$(Trim$(Parse(1))) <> "no sound" Then
            If Trim$(Parse(1)) <> "" Then
                If Not PlayWAV(Trim$(Parse(1)) & ".wav", True) Then
                    MsgBox "Error playing wave file!"
                End If
                'Call DirectMusic.PlayMusic(Trim$(Parse(1)) & ".wav")
                Debug.Print "Sound " & Trim$(Parse(1)) & " played!"
            End If
        End If
    End If
    
    ' The pause map command
    If LCase$(Parse(0)) = "pausemap" Then
        'Most likely, we are receiving new data and don't want to
        'have an error from players moving on that map
        If LCase$(Trim$(Parse(1))) = "lock" Then
            PauseMap = True
            ' Check if we have a message to display, if so, set it
            If Parse(2) <> END_CHAR Then
                PauseMessage = Trim$(Parse(2))
            Else
                PauseMessage = vbNullString
            End If
        ElseIf LCase$(Trim$(Parse(1))) = "unlock" Then
            PauseMap = False
        End If
    End If
    
    ' The map extra notice
    If LCase$(Parse(0)) = "mapextra" Then
        Select Case LCase$(Trim$(Parse(1)))
        
        Case "allowmovement":
            AllowMovement = True
            Exit Sub
        
        Case "disallowmovement":
            AllowMovement = False
            Exit Sub
        
        End Select
    End If
    
    ' Records the damage received and displays it
    If LCase$(Parse(0)) = "attacknpc" Then
        ' Fill a variable!
        n = Val(Parse(1))
        
        ' First, find an open damage slot
        For i = 1 To 5
            ' Check if we've found an empty slot
            If MapNpc(n).Damage(i).Display = False Then
                ' Set the display to true
                MapNpc(n).Damage(i).Display = True
                
                ' Set the position
                MapNpc(n).Damage(i).TextX = MapNpc(n).X * PIC_X + PIC_X
                MapNpc(n).Damage(i).TextY = MapNpc(n).Y * PIC_Y + PIC_Y
                
                ' Readjust if needed
                If MapNpc(n).Damage(i).TextX + 20 >= frmDualSolace.picScreen.ScaleWidth Then
                    MapNpc(n).Damage(i).TextX = MapNpc(n).Damage(i).TextX - PIC_X
                End If
                
                ' Reset the count
                MapNpc(n).Damage(i).AnimCount = 0
                
                ' Set the damage value
                MapNpc(n).Damage(i).Value = Val(Parse(2))
                
                ' Exit the loop so we don't mark another damage set
                Exit For
            Else
                ' Check if this was the final loop, if so,
                ' overwrite the first one
                If i = 5 Then
                    ' Set the display to true
                    MapNpc(n).Damage(1).Display = True
                
                    ' Set the position
                    MapNpc(n).Damage(1).TextX = MapNpc(n).X * PIC_X + PIC_X
                    MapNpc(n).Damage(1).TextY = MapNpc(n).Y * PIC_Y + PIC_Y
                
                    ' Readjust if needed
                    If MapNpc(n).Damage(1).TextX + 20 >= frmDualSolace.picScreen.ScaleWidth Then
                        MapNpc(n).Damage(1).TextX = MapNpc(n).Damage(1).TextX - PIC_X
                    End If
                
                    ' Reset the count
                    MapNpc(n).Damage(1).AnimCount = 0
                
                    ' Set the damage value
                    MapNpc(n).Damage(1).Value = Val(Parse(2))
                
                    ' Exit the loop so we don't mark another damage set
                    Exit For
                End If
            End If
        Next i
    End If
    
    ' Records the damage received and displays it
    If LCase$(Parse(0)) = "attackplayer" Then
        ' Fill a variable!
        n = Val(Parse(1))
        
        ' First, find an open damage slot
        For i = 1 To 5
            ' Check if we've found an empty slot
            If Player(n).Damage(i).Display = False Then
                ' Set the display to true
                Player(n).Damage(i).Display = True

                ' Set the position
                Player(n).Damage(i).TextX = Player(n).X * PIC_X + PIC_X
                Player(n).Damage(i).TextY = Player(n).Y * PIC_Y + PIC_Y
                
                ' Readjust if needed
                If Player(n).Damage(i).TextX + 20 >= frmDualSolace.picScreen.ScaleWidth Then
                    Player(n).Damage(i).TextX = Player(n).Damage(i).TextX - PIC_X
                End If
                
                ' Reset the count
                Player(n).Damage(i).AnimCount = 0
                
                ' Set the damage value
                Player(n).Damage(i).Value = Val(Parse(2))
                
                ' Exit the loop so we don't mark another damage set
                Exit For
            Else
                ' Check if this was the final loop, if so,
                ' overwrite the first one
                If i = 5 Then
                    ' Set the display to true
                    Player(n).Damage(1).Display = True
                
                    ' Set the position
                    Player(n).Damage(1).TextX = Player(n).X * PIC_X + PIC_X
                    Player(n).Damage(1).TextY = Player(n).Y * PIC_Y + PIC_Y
                
                    ' Readjust if needed
                    If Player(n).Damage(1).TextX + 20 >= frmDualSolace.picScreen.ScaleWidth Then
                        Player(n).Damage(1).TextX = Player(n).Damage(1).TextX - PIC_X
                    End If
                
                    ' Reset the count
                    Player(n).Damage(1).AnimCount = 0
                
                    ' Set the damage value
                    Player(n).Damage(1).Value = Val(Parse(2))
                
                    ' Exit the loop so we don't mark another damage set
                    Exit For
                End If
            End If
        Next i
    End If
    
    ' Records the damage received and displays it
    If LCase$(Parse(0)) = "killplayer" Then
        n = Val(Parse(1))
        Player(n).Death.Display = True
        Player(n).Death.AnimCount = 0
        Player(n).Death.DeathMap = GetPlayerMap(MyIndex)
        Player(n).Death.DeathX = GetPlayerX(n)
        Player(n).Death.DeathY = GetPlayerY(n)
    End If
End Sub
