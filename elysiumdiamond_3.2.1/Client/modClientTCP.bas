Attribute VB_Name = "modClientTCP"

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean
Public TradePlayer As Long

Sub TcpInit()
    SEP_CHAR = Chr$(253)
    END_CHAR = Chr$(254)
    NEXT_CHAR = Chr$(255)
    GoDebug = NO
    PlayerBuffer = vbNullString
    
    Dim Filename As String
    Filename = App.Path & "\config.ini"

    frmMirage.Socket.RemoteHost = GetVar(Filename, "IPCONFIG", "IP")
    frmMirage.Socket.RemotePort = Val(GetVar(Filename, "IPCONFIG", "PORT"))
End Sub

Sub TcpDestroy()
    frmMirage.Socket.Close
    
    If frmChars.Visible Then frmChars.Visible = False
    If frmCredits.Visible Then frmCredits.Visible = False
    If frmDeleteAccount.Visible Then frmDeleteAccount.Visible = False
    If frmLogin.Visible Then frmLogin.Visible = False
    If frmNewAccount.Visible Then frmNewAccount.Visible = False
    If frmNewChar.Visible Then frmNewChar.Visible = False
End Sub

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim Top As String * 3
Dim Start As Long

    frmMirage.Socket.GetData Buffer, vbString, DataLength
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

Sub HandleData(ByVal Data As String)
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
Dim Direction As Long
Dim InvNum As Long
Dim Ammount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Level As Long
Dim I As Long, n As Long, x As Long, y As Long, f As Long
Dim GiveItem As Long, GiveValue As Long, GetItem As Long, GetValue As Long
Dim z As Long

    ' Handle Data
    Parse$ = Split(Data, SEP_CHAR)
    
    ' Add the data to the debug window if we are in debug mode
    'If Trim$(Command) = "-debug" Then
    If GoDebug = YES Then
        If frmDebug.Visible = False Then frmDebug.Visible = True
        Call TextAdd(frmDebug.txtDebug, "((( Processed Packet " & Parse$(0) & " )))", True)
    End If
    
    'Parse$(0) = LCase$(Parse$(0))
    
    ' :::::::::::::::::::::::
    ' :: Get players stats ::
    ' :::::::::::::::::::::::
    If Parse$(0) = MAXINFO_CHAR Then
        GAME_NAME = Trim$(Parse$(1))
        MAX_PLAYERS = Val(Parse$(2))
        MAX_ITEMS = Val(Parse$(3))
        MAX_NPCS = Val(Parse$(4))
        MAX_SHOPS = Val(Parse$(5))
        MAX_SPELLS = Val(Parse$(6))
        MAX_MAPS = Val(Parse$(7))
        MAX_MAP_ITEMS = Val(Parse$(8))
        MAX_MAPX = Val(Parse$(9))
        MAX_MAPY = Val(Parse$(10))
        MAX_EMOTICONS = Val(Parse$(11))
        MAX_SPEECH = Val(Parse$(12))
        
        ReDim Map(1 To MAX_MAPS) As MapRec
        ReDim Player(1 To MAX_PLAYERS) As PlayerRec
        ReDim Item(1 To MAX_ITEMS) As ItemRec
        ReDim Npc(1 To MAX_NPCS) As NpcRec
        ReDim MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim Shop(1 To MAX_SHOPS) As ShopRec
        ReDim Spell(1 To MAX_SPELLS) As SpellRec
        ReDim Bubble(1 To MAX_PLAYERS) As ChatBubble
        ReDim SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        For I = 1 To MAX_MAPS
            ReDim Map(I).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
            ReDim Map(I).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        Next I
        ReDim TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
        ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
        ReDim MapReport(1 To MAX_MAPS) As MapRec
        MAX_SPELL_ANIM = MAX_MAPX * MAX_MAPY
        
        ReDim Speech(1 To MAX_SPEECH) As SpeechRec
        
        MAX_BLT_LINE = 6
        ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
        ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
        
        For I = 1 To MAX_PLAYERS
            ReDim Player(I).SpellAnim(1 To MAX_SPELL_ANIM) As SpellAnimRec
        Next I
        
        For I = 0 To MAX_EMOTICONS
            Emoticons(I).Pic = 0
            Emoticons(I).Command = vbNullString
        Next I
        
        Call ClearMaps
        Call ClearTempTile
        Call ClearSpeeches
        
        ' Clear out players
        For I = 1 To MAX_PLAYERS
            Call ClearPlayer(I)
        Next I
        
        For I = 1 To MAX_MAPS
            Call LoadMap(I)
        Next I
    
        frmMirage.Caption = Trim$(GAME_NAME)
        App.Title = GAME_NAME
 
        Exit Sub
    End If
    
    If Parse$(0) = INFO_CHAR Then
        frmLogin.lblOnOff.Caption = "Online"
        frmLogin.lblPlayers.Caption = Parse$(1) & " Players Online"
        frmLogin.lblPlayers.Visible = True
        frmLogin.tmrInfo.Enabled = False
        Exit Sub
    End If
        
    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    If Parse$(0) = NPCHP_CHAR Then
        n = Val(Parse$(1))
 
        MapNpc(n).HP = Val(Parse$(2))
        MapNpc(n).MaxHP = Val(Parse$(3))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    If Parse$(0) = ALERTMSG_CHAR Then
        frmMirage.Visible = False
        frmSendGetData.Visible = False
        frmMainMenu.Visible = True
        DoEvents

        Msg = Parse$(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Plain message packet ::
    ' ::::::::::::::::::::::::::
    If Parse$(0) = PLAINMSG_CHAR Then
        frmSendGetData.Visible = False
        n = Val(Parse$(2))
        
        If n = 1 Then frmNewAccount.Show
        If n = 2 Then frmDeleteAccount.Show
        If n = 3 Then frmLogin.Show
        If n = 4 Then frmNewChar.Show
        If n = 5 Then frmChars.Show
        
        Msg = Parse$(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
    If Parse$(0) = ALLCHARS_CHAR Then
        n = 1
        
        frmChars.Visible = True
        frmSendGetData.Visible = False
        
        frmChars.lstChars.Clear
        
        For I = 1 To MAX_CHARS
            Name = Parse$(n)
            Msg = Parse$(n + 1)
            Level = Val(Parse$(n + 2))
            
            If Trim$(Name) = vbNullString Then
                frmChars.lstChars.AddItem "Free Character Slot"
            Else
                frmChars.lstChars.AddItem Name & " a level " & Level & " " & Msg
            End If
            
            n = n + 3
        Next I
        
        frmChars.lstChars.ListIndex = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    If Parse$(0) = LOGINOK_CHAR Then
        ' Now we can receive game data
        MyIndex = Val(Parse$(1))
        
        frmSendGetData.Visible = True
        frmChars.Visible = False
        
        Call SetStatus("Receiving game data...")
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If Parse$(0) = NEWCHARCLASSES_CHAR Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse$(n))
        ReDim Class(1 To Max_Classes) As ClassRec
        
        n = n + 1

        For I = 1 To Max_Classes
            Class(I).Name = Parse$(n)
            
            Class(I).HP = Val(Parse$(n + 1))
            Class(I).MP = Val(Parse$(n + 2))
            Class(I).SP = Val(Parse$(n + 3))
            
            Class(I).STR = Val(Parse$(n + 4))
            Class(I).DEF = Val(Parse$(n + 5))
            Class(I).Speed = Val(Parse$(n + 6))
            Class(I).MAGI = Val(Parse$(n + 7))
            'Class(i).INTEL = Val(Parse$(n + 8))
            Class(I).MaleSprite = Val(Parse$(n + 8))
            Class(I).FemaleSprite = Val(Parse$(n + 9))
            Class(I).Locked = Val(Parse$(n + 10))
        
        n = n + 11
        Next I
        
        ' Used for if the player is creating a new character
        frmNewChar.Visible = True
        frmSendGetData.Visible = False

        frmNewChar.cmbClass.Clear
        For I = 1 To Max_Classes
            If Class(I).Locked = 0 Then
                frmNewChar.cmbClass.AddItem Trim$(Class(I).Name)
            End If
        Next I
        
        frmNewChar.cmbClass.ListIndex = 0
        frmNewChar.lblHP.Caption = STR(Class(1).HP)
        frmNewChar.lblMP.Caption = STR(Class(1).MP)
        frmNewChar.lblSP.Caption = STR(Class(1).SP)
    
        frmNewChar.lblSTR.Caption = STR(Class(1).STR)
        frmNewChar.lblDEF.Caption = STR(Class(1).DEF)
        frmNewChar.lblSPEED.Caption = STR(Class(1).Speed)
        frmNewChar.lblMAGI.Caption = STR(Class(1).MAGI)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If Parse$(0) = CLASSESDATA_CHAR Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse$(n))
        ReDim Class(1 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For I = 1 To Max_Classes
            Class(I).Name = Parse$(n)
            
            Class(I).HP = Val(Parse$(n + 1))
            Class(I).MP = Val(Parse$(n + 2))
            Class(I).SP = Val(Parse$(n + 3))
            
            Class(I).STR = Val(Parse$(n + 4))
            Class(I).DEF = Val(Parse$(n + 5))
            Class(I).Speed = Val(Parse$(n + 6))
            Class(I).MAGI = Val(Parse$(n + 7))
            
            Class(I).Locked = Val(Parse$(n + 8))
            
            n = n + 9
        Next I
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If Parse$(0) = INGAME_CHAR Then
        InGame = True
        Call GameInit
        Call GameLoop
        If Parse$(1) = END_CHAR Then
            MsgBox ("here")
            End
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    If Parse$(0) = PLAYERINV_CHAR Then
        n = 2
        z = Val(Parse$(1))
        
        For I = 1 To MAX_INV
            Call SetPlayerInvItemNum(z, I, Val(Parse$(n)))
            Call SetPlayerInvItemValue(z, I, Val(Parse$(n + 1)))
            Call SetPlayerInvItemDur(z, I, Val(Parse$(n + 2)))
            
            n = n + 3
        Next I
        
        If z = MyIndex Then Call UpdateVisInv
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If Parse$(0) = PLAYERINVUPDATE_CHAR Then
        n = Val(Parse$(1))
        z = Val(Parse$(2))
        
        Call SetPlayerInvItemNum(z, n, Val(Parse$(3)))
        Call SetPlayerInvItemValue(z, n, Val(Parse$(4)))
        Call SetPlayerInvItemDur(z, n, Val(Parse$(5)))
        If z = MyIndex Then Call UpdateVisInv
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    If Parse$(0) = PLAYERWORNEQ_CHAR Then
        z = Val(Parse$(1))
        If z <= 0 Then Exit Sub
        
        Call SetPlayerArmorSlot(z, Val(Parse$(2)))
        Call SetPlayerWeaponSlot(z, Val(Parse$(3)))
        Call SetPlayerHelmetSlot(z, Val(Parse$(4)))
        Call SetPlayerShieldSlot(z, Val(Parse$(5)))
        
        If z = MyIndex Then
            Call UpdateVisInv
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player points packet ::
    ' ::::::::::::::::::::::::::
    If Parse$(0) = PLAYERPOINTS_CHAR Then
        Player(MyIndex).POINTS = Val(Parse$(1))
        frmMirage.lblPoints.Caption = Val(Parse$(1)) & " points"
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    If Parse$(0) = PLAYERHP_CHAR Then
        Player(MyIndex).MaxHP = Val(Parse$(1))
        Call SetPlayerHP(MyIndex, Val(Parse$(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then
            frmMirage.shpHP.FillColor = RGB(208, 11, 0)
            frmMirage.shpHP.Width = (((GetPlayerHP(MyIndex) / frmMirage.lblHP.Width) / (GetPlayerMaxHP(MyIndex) / frmMirage.lblHP.Width)) * frmMirage.lblHP.Width)
            frmMirage.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Pet hp packet ::
    ' :::::::::::::::::::
    If Parse$(0) = PETHP_CHAR Then
        Player(MyIndex).Pet.MaxHP = Val(Parse$(1))
        Player(MyIndex).Pet.HP = Val(Parse$(2))
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player mp packet ::
    ' ::::::::::::::::::::::
    If Parse$(0) = PLAYERMP_CHAR Then
        Player(MyIndex).MaxMP = Val(Parse$(1))
        Call SetPlayerMP(MyIndex, Val(Parse$(2)))
        If GetPlayerMaxMP(MyIndex) > 0 Then
            frmMirage.shpMP.FillColor = RGB(208, 11, 0)
            frmMirage.shpMP.Width = (((GetPlayerMP(MyIndex) / frmMirage.lblMP.Width) / (GetPlayerMaxMP(MyIndex) / frmMirage.lblMP.Width)) * frmMirage.lblMP.Width)
            frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
        End If
        Exit Sub
    End If
    
    ' speech bubble parse
    If (Parse$(0) = MAPMSG2_CHAR) Then
        Bubble(Val(Parse$(2))).Text = Parse$(1)
        Bubble(Val(Parse$(2))).Created = GetTickCount()
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    If Parse$(0) = PLAYERSP_CHAR Then
       ' Player(MyIndex).MaxSP = Val(Parse$(1))
        'Call SetPlayerSP(MyIndex, Val(Parse$(2)))
        'If GetPlayerMaxSP(MyIndex) > 0 Then
            'frmMirage.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
        'End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Player Stats Packet ::
    ' :::::::::::::::::::::::::
    If (Parse$(0) = PLAYERSTATSPACKET_CHAR) Then
        Dim SubStr As Long, SubDef As Long, SubMagi As Long, SubSpeed As Long
        SubStr = 0
        SubDef = 0
        SubMagi = 0
        SubSpeed = 0
        
        If GetPlayerWeaponSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerArmorSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerShieldSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerHelmetSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddSpeed
        End If
        
        If SubStr > 0 Then
            frmMirage.lblSTR.Caption = "Strength - " & Val(Parse$(1)) - SubStr & " (+" & SubStr & ")"
        Else
            frmMirage.lblSTR.Caption = "Strength - " & Val(Parse$(1))
        End If
        If SubDef > 0 Then
            frmMirage.lblDEF.Caption = "Defense - " & Val(Parse$(2)) - SubDef & " (+" & SubDef & ")"
        Else
            frmMirage.lblDEF.Caption = "Defense - " & Val(Parse$(2))
        End If
        If SubMagi > 0 Then
            frmMirage.lblMAGI.Caption = "Magic - " & Val(Parse$(4)) - SubMagi & " (+" & SubMagi & ")"
        Else
            frmMirage.lblMAGI.Caption = "Magic - " & Val(Parse$(4))
        End If
        If SubSpeed > 0 Then
            frmMirage.lblSPEED.Caption = "Speed - " & Val(Parse$(3)) - SubSpeed & " (+" & SubSpeed & ")"
        Else
            frmMirage.lblSPEED.Caption = "Speed - " & Val(Parse$(3))
        End If
        frmMirage.lblEXP.Caption = Val(Parse$(6)) & " / " & Val(Parse$(5))
        
        frmMirage.shpTNL.Width = (((Val(Parse$(6)) / frmMirage.lblEXP.Width) / (Val(Parse$(5)) / frmMirage.lblEXP.Width)) * frmMirage.lblEXP.Width)
        frmMirage.lblLevel.Caption = "Level " & Val(Parse$(7))
        
        Exit Sub
    End If
                

    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If Parse$(0) = PLAYERDATA_CHAR Then
        I = Val(Parse$(1))
        Call SetPlayerName(I, Parse$(2))
        Call SetPlayerSprite(I, Val(Parse$(3)))
        Call SetPlayerMap(I, Val(Parse$(4)))
        Call SetPlayerX(I, Val(Parse$(5)))
        Call SetPlayerY(I, Val(Parse$(6)))
        Call SetPlayerDir(I, Val(Parse$(7)))
        Call SetPlayerAccess(I, Val(Parse$(8)))
        Call SetPlayerPK(I, Val(Parse$(9)))
        Call SetPlayerGuild(I, Parse$(10))
        Call SetPlayerGuildAccess(I, Val(Parse$(11)))
        Call SetPlayerClass(I, Val(Parse$(12)))

        ' Make sure they aren't walking
        Player(I).Moving = 0
        Player(I).XOffset = 0
        Player(I).YOffset = 0
        
        ' Check if the player is the client player, and if so reset Directions
        If I = MyIndex Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Pet data packet ::
    ' :::::::::::::::::::::
    If Parse$(0) = PETDATA_CHAR Then
        I = Val(Parse$(1))
        
        Player(I).Pet.Alive = Val(Parse$(2))
        Player(I).Pet.Map = Val(Parse$(3))
        Player(I).Pet.x = Val(Parse$(4))
        Player(I).Pet.y = Val(Parse$(5))
        Player(I).Pet.Dir = Val(Parse$(6))
        Player(I).Pet.Sprite = Val(Parse$(7))
        Player(I).Pet.HP = Val(Parse$(8))
        Player(I).Pet.MaxHP = Val(Parse$(9))
        
        ' Make sure their pet isn't walking
        Player(I).Pet.Moving = 0
        Player(I).Pet.XOffset = 0
        Player(I).Pet.YOffset = 0
        
        ' Check if the player is the client player, and if so reset Directions
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
    If (Parse$(0) = PLAYERMOVE_CHAR) Then
        I = Val(Parse$(1))
        x = Val(Parse$(2))
        y = Val(Parse$(3))
        Direction = Val(Parse$(4))
        n = Val(Parse$(5))

        If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub

        Call SetPlayerX(I, x)
        Call SetPlayerY(I, y)
        Call SetPlayerDir(I, Direction)
        
        Player(I).XOffset = 0
        Player(I).YOffset = 0
        Player(I).Moving = n
        
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
    ' :: Pet movement packet ::
    ' :::::::::::::::::::::::::
    If (Parse$(0) = PETMOVE_CHAR) Then
        I = Val(Parse$(1))
        x = Val(Parse$(2))
        y = Val(Parse$(3))
        Direction = Val(Parse$(4))
        n = Val(Parse$(5))

        Player(I).Pet.x = x
        Player(I).Pet.y = y
        Player(I).Pet.Dir = Direction
        Player(I).Pet.XOffset = 0
        Player(I).Pet.YOffset = 0
        Player(I).Pet.Moving = MOVING_WALKING
        
        Select Case Player(I).Pet.Dir
            Case DIR_UP
                Player(I).Pet.YOffset = PIC_Y
            Case DIR_DOWN
                Player(I).Pet.YOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(I).Pet.XOffset = PIC_X
            Case DIR_RIGHT
                Player(I).Pet.XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (Parse$(0) = NPCMOVE_CHAR) Then
        I = Val(Parse$(1))
        x = Val(Parse$(2))
        y = Val(Parse$(3))
        Direction = Val(Parse$(4))
        n = Val(Parse$(5))

        MapNpc(I).x = x
        MapNpc(I).y = y
        MapNpc(I).Dir = Direction
        MapNpc(I).XOffset = 0
        MapNpc(I).YOffset = 0
        MapNpc(I).Moving = n
        
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
    ' :: Player Direction packet ::
    ' :::::::::::::::::::::::::::::
    If (Parse$(0) = PLAYERDIR_CHAR) Then
        I = Val(Parse$(1))
        Direction = Val(Parse$(2))
        
        If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub

        Call SetPlayerDir(I, Direction)
        
        Player(I).XOffset = 0
        Player(I).YOffset = 0
        Player(I).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC Direction packet ::
    ' ::::::::::::::::::::::::::
    If (Parse$(0) = NPCDIR_CHAR) Then
        I = Val(Parse$(1))
        Direction = Val(Parse$(2))
        MapNpc(I).Dir = Direction
        
        MapNpc(I).XOffset = 0
        MapNpc(I).YOffset = 0
        MapNpc(I).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (Parse$(0) = PLAYERXY_CHAR) Then
        x = Val(Parse$(1))
        y = Val(Parse$(2))
        
        Call SetPlayerX(MyIndex, x)
        Call SetPlayerY(MyIndex, y)
        
        ' Make sure they aren't walking
        Player(MyIndex).Moving = 0
        Player(MyIndex).XOffset = 0
        Player(MyIndex).YOffset = 0
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If (Parse$(0) = ATTACKPLAYER_CHAR) Then
        I = Val(Parse$(1))
        
        ' Set player to attacking
        Player(I).Attacking = 1
        Player(I).AttackTimer = GetTickCount
        Player(I).LastAttack = GetTickCount
        
        Exit Sub
    End If
    
    If (Parse$(0) = ATTACKNPC_CHAR) Then
        I = Val(Parse$(1))
        
        ' Set player to attacking
        Player(I).Attacking = 1
        Player(I).AttackTimer = GetTickCount
        Player(I).LastAttack = GetTickCount
        
        ' The server now also keeps track, just to let you know
        MapNpc(Val(Parse$(2))).LastAttack = GetTickCount
        Exit Sub
    End If
    
    If (Parse$(0) = PETATTACKNPC_CHAR) Then
        I = Val(Parse$(1))
        
        ' Set pet to attacking
        Player(I).Pet.Attacking = 1
        Player(I).Pet.AttackTimer = GetTickCount
        Player(I).Pet.LastAttack = GetTickCount
        
        ' The server now also keeps track, just to let you know
        MapNpc(Val(Parse$(2))).LastAttack = GetTickCount
        Exit Sub
    End If
    
    If (Parse$(0) = NPCATTACK_CHAR) Then
        I = Val(Parse$(1))
        
        ' Set npc to attacking
        MapNpc(I).Attacking = 1
        MapNpc(I).AttackTimer = GetTickCount
        MapNpc(I).LastAttack = GetTickCount
        
        Player(Val(Parse$(2))).LastAttack = GetTickCount
        Exit Sub
    End If
    
    If (Parse$(0) = NPCATTACKPET_CHAR) Then
        I = Val(Parse$(1))
        
        ' Set npc to attacking
        MapNpc(I).Attacking = 1
        MapNpc(I).AttackTimer = GetTickCount
        MapNpc(I).LastAttack = GetTickCount
        
        Player(Val(Parse$(2))).Pet.LastAttack = GetTickCount
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (Parse$(0) = CHECKFORMAP_CHAR) Then
        ' Erase all players except self
        For I = 1 To MAX_PLAYERS
            If I <> MyIndex Then
                Call SetPlayerMap(I, 0)
            End If
        Next I
        
        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        x = Val(Parse$(1))
        
        ' Get revision
        y = Val(Parse$(2))
        
        If FileExist("maps\map" & x & ".dat") Then
            ' Check to see if the revisions match
            If GetMapRevision(x) = y Then
                ' We do so we don't need the map
                
                ' Load the map
                Call LoadMap(x)
                
                Call SendData(NEEDMAP_CHAR & SEP_CHAR & "no" & END_CHAR)
                Exit Sub
            End If
        End If
        
        ' Either the revisions didn't match or we don't have the map, so we need it
        Call SendData(NEEDMAP_CHAR & SEP_CHAR & "yes" & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    If Parse$(0) = MAPDATA_CHAR Then
        n = 1
        I = Val(Parse$(1))
        
        Call ClearMap(I)
        
        Map(I).Name = Parse$(n + 1)
        Map(I).Revision = Val(Parse$(n + 2))
        Map(I).Moral = Val(Parse$(n + 3))
        Map(I).Up = Val(Parse$(n + 4))
        Map(I).Down = Val(Parse$(n + 5))
        Map(I).Left = Val(Parse$(n + 6))
        Map(I).Right = Val(Parse$(n + 7))
        Map(I).Music = Parse$(n + 8)
        Map(I).BootMap = Val(Parse$(n + 9))
        Map(I).BootX = Val(Parse$(n + 10))
        Map(I).BootY = Val(Parse$(n + 11))
        Map(I).Indoors = Val(Parse$(n + 12))
        
        n = n + 13
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Ground = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).GroundSet = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Mask = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).MaskSet = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Anim = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).AnimSet = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Fringe = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).FringeSet = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Type = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Data1 = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Data2 = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Data3 = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).String1 = Parse$(n)
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).String2 = Parse$(n)
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).String3 = Parse$(n)
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Mask2 = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Mask2Set = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).M2Anim = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).M2AnimSet = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).FAnim = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).FAnimSet = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Fringe2 = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Fringe2Set = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).Light = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).F2Anim = Val(Parse$(n))
                    n = n + 1
                End If
                If Parse$(n) <> NEXT_CHAR Then
                    Map(I).Tile(x, y).F2AnimSet = Val(Parse$(n))
                    n = n + 1
                End If
                n = n + 1
            Next x
        Next y
        
        For x = 1 To MAX_MAP_NPCS
            Map(I).Npc(x) = Val(Parse$(n))
            Map(I).NpcSpawn(x).Used = Val(Parse$(n + 1))
            Map(I).NpcSpawn(x).x = Val(Parse$(n + 2))
            Map(I).NpcSpawn(x).y = Val(Parse$(n + 3))
            n = n + 4
        Next x
                
        ' Save the map
        Call SaveLocalMap(I)
        
        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmMapEditor.Visible = False
            'frmAttributes.Visible = False
            frmMirage.Show
            'frmMirage.picMapEditor.Visible = False
            
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
    If Parse$(0) = MAPITEMDATA_CHAR Then
        n = 1
        
        For I = 1 To MAX_MAP_ITEMS
            SaveMapItem(I).Num = Val(Parse$(n))
            SaveMapItem(I).Value = Val(Parse$(n + 1))
            SaveMapItem(I).Dur = Val(Parse$(n + 2))
            SaveMapItem(I).x = Val(Parse$(n + 3))
            SaveMapItem(I).y = Val(Parse$(n + 4))
            
            n = n + 5
        Next I
        
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If Parse$(0) = MAPNPCDATA_CHAR Then
        n = 1
        
        For I = 1 To MAX_MAP_NPCS
            SaveMapNpc(I).Num = Val(Parse$(n))
            SaveMapNpc(I).x = Val(Parse$(n + 1))
            SaveMapNpc(I).y = Val(Parse$(n + 2))
            SaveMapNpc(I).Dir = Val(Parse$(n + 3))
            
            n = n + 4
        Next I
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If Parse$(0) = MAPDONE_CHAR Then
        'Map = SaveMap
        
        For I = 1 To MAX_MAP_ITEMS
            MapItem(I) = SaveMapItem(I)
        Next I
        
        For I = 1 To MAX_MAP_NPCS
            MapNpc(I) = SaveMapNpc(I)
        Next I
        
        GettingMap = False
        
        ' Play music
        If Trim$(Map(GetPlayerMap(MyIndex)).Music) <> "None" Then
            Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
        Else
            Call StopMidi
        End If
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (Parse$(0) = SAYMSG_CHAR) Or (Parse$(0) = BROADCASTMSG_CHAR) Or (Parse$(0) = GLOBALMSG_CHAR) Or (Parse$(0) = PLAYERMSG_CHAR) Or (Parse$(0) = MAPMSG_CHAR) Or (Parse$(0) = ADMINMSG_CHAR) Then
        Call AddText(Parse$(1), Val(Parse$(2)))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If Parse$(0) = SPAWNITEM_CHAR Then
        n = Val(Parse$(1))
        
        MapItem(n).Num = Val(Parse$(2))
        MapItem(n).Value = Val(Parse$(3))
        MapItem(n).Dur = Val(Parse$(4))
        MapItem(n).x = Val(Parse$(5))
        MapItem(n).y = Val(Parse$(6))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    If (Parse$(0) = ITEMEDITOR_CHAR) Then
        InItemsEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For I = 1 To MAX_ITEMS
            frmIndex.lstIndex.AddItem I & ": " & Trim$(Item(I).Name)
        Next I
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (Parse$(0) = UPDATEITEM_CHAR) Then
        n = Val(Parse$(1))
        
        ' Update the item
        Item(n).Name = Parse$(2)
        Item(n).Pic = Val(Parse$(3))
        Item(n).Type = Val(Parse$(4))
        Item(n).Data1 = Val(Parse$(5))
        Item(n).Data2 = Val(Parse$(6))
        Item(n).Data3 = Val(Parse$(7))
        Item(n).StrReq = Val(Parse$(8))
        Item(n).DefReq = Val(Parse$(9))
        Item(n).SpeedReq = Val(Parse$(10))
        Item(n).MagicReq = Val(Parse$(11))
        Item(n).ClassReq = Val(Parse$(12))
        Item(n).AccessReq = Val(Parse$(13))
        
        Item(n).AddHP = Val(Parse$(14))
        Item(n).AddMP = Val(Parse$(15))
        Item(n).AddSP = Val(Parse$(16))
        Item(n).AddStr = Val(Parse$(17))
        Item(n).AddDef = Val(Parse$(18))
        Item(n).AddMagi = Val(Parse$(19))
        Item(n).AddSpeed = Val(Parse$(20))
        Item(n).AddEXP = Val(Parse$(21))
        Item(n).desc = Parse$(22)
        Item(n).AttackSpeed = Val(Parse$(23))
        Exit Sub
    End If
       
    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (Parse$(0) = EDITITEM_CHAR) Then
        n = Val(Parse$(1))
        
        ' Update the item
        Item(n).Name = Parse$(2)
        Item(n).Pic = Val(Parse$(3))
        Item(n).Type = Val(Parse$(4))
        Item(n).Data1 = Val(Parse$(5))
        Item(n).Data2 = Val(Parse$(6))
        Item(n).Data3 = Val(Parse$(7))
        Item(n).StrReq = Val(Parse$(8))
        Item(n).DefReq = Val(Parse$(9))
        Item(n).SpeedReq = Val(Parse$(10))
        Item(n).MagicReq = Val(Parse$(11))
        Item(n).ClassReq = Val(Parse$(12))
        Item(n).AccessReq = Val(Parse$(13))
        
        Item(n).AddHP = Val(Parse$(14))
        Item(n).AddMP = Val(Parse$(15))
        Item(n).AddSP = Val(Parse$(16))
        Item(n).AddStr = Val(Parse$(17))
        Item(n).AddDef = Val(Parse$(18))
        Item(n).AddMagi = Val(Parse$(19))
        Item(n).AddSpeed = Val(Parse$(20))
        Item(n).AddEXP = Val(Parse$(21))
        Item(n).desc = Parse$(22)
        Item(n).AttackSpeed = Val(Parse$(23))
        
        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If Parse$(0) = SPAWNNPC_CHAR Then
        n = Val(Parse$(1))
        
        MapNpc(n).Num = Val(Parse$(2))
        MapNpc(n).x = Val(Parse$(3))
        MapNpc(n).y = Val(Parse$(4))
        MapNpc(n).Dir = Val(Parse$(5))
        MapNpc(n).Big = Val(Parse$(6))
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If Parse$(0) = NPCDEAD_CHAR Then
        n = Val(Parse$(1))
        
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
    
    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    If (Parse$(0) = NPCEDITOR_CHAR) Then
        InNpcEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For I = 1 To MAX_NPCS
            frmIndex.lstIndex.AddItem I & ": " & Trim$(Npc(I).Name)
        Next I
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (Parse$(0) = UPDATENPC_CHAR) Then
        n = Val(Parse$(1))
        
        ' Update the npc
        Npc(n).Name = Parse$(2)
        Npc(n).AttackSay = vbNullString
        Npc(n).Sprite = Val(Parse$(3))
        Npc(n).SpawnSecs = 0
        Npc(n).Behavior = 0
        Npc(n).Range = 0
        For I = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(I).Chance = 0
            Npc(n).ItemNPC(I).ItemNum = 0
            Npc(n).ItemNPC(I).ItemValue = 0
        Next I
        Npc(n).STR = 0
        Npc(n).DEF = 0
        Npc(n).Speed = 0
        Npc(n).MAGI = 0
        Npc(n).Big = Val(Parse$(4))
        Npc(n).MaxHP = Val(Parse$(5))
        Npc(n).EXP = 0
        Npc(n).Speech = Val(Parse$(6))
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet :: <- Used for item editor admins only
    ' :::::::::::::::::::::
    If (Parse$(0) = EDITNPC_CHAR) Then
        n = Val(Parse$(1))
        
        ' Update the npc
        Npc(n).Name = Parse$(2)
        Npc(n).AttackSay = Parse$(3)
        Npc(n).Sprite = Val(Parse$(4))
        Npc(n).SpawnSecs = Val(Parse$(5))
        Npc(n).Behavior = Val(Parse$(6))
        Npc(n).Range = Val(Parse$(7))
        Npc(n).STR = Val(Parse$(8))
        Npc(n).DEF = Val(Parse$(9))
        Npc(n).Speed = Val(Parse$(10))
        Npc(n).MAGI = Val(Parse$(11))
        Npc(n).Big = Val(Parse$(12))
        Npc(n).MaxHP = Val(Parse$(13))
        Npc(n).EXP = Val(Parse$(14))
        Npc(n).SpawnTime = Val(Parse$(15))
        Npc(n).Speech = Val(Parse$(16))
        
        z = 17
        For I = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(I).Chance = Val(Parse$(z))
            Npc(n).ItemNPC(I).ItemNum = Val(Parse$(z + 1))
            Npc(n).ItemNPC(I).ItemValue = Val(Parse$(z + 2))
            z = z + 3
        Next I
        
        ' Initialize the npc editor
        Call NpcEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (Parse$(0) = MAPKEY_CHAR) Then
        x = Val(Parse$(1))
        y = Val(Parse$(2))
        n = Val(Parse$(3))
                
        TempTile(x, y).DoorOpen = n
        
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
    If (Parse$(0) = EDITMAP_CHAR) Then
        Call EditorInit
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    If (Parse$(0) = SHOPEDITOR_CHAR) Then
        InShopEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For I = 1 To MAX_SHOPS
            frmIndex.lstIndex.AddItem I & ": " & Trim$(Shop(I).Name)
        Next I
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (Parse$(0) = UPDATESHOP_CHAR) Then
        n = Val(Parse$(1))
        
        ' Update the shop name
        Shop(n).Name = Parse$(2)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    If (Parse$(0) = EDITSHOP_CHAR) Then
        ShopNum = Val(Parse$(1))
        
        ' Update the shop
        Shop(ShopNum).Name = Parse$(2)
        Shop(ShopNum).JoinSay = Parse$(3)
        Shop(ShopNum).LeaveSay = Parse$(4)
        Shop(ShopNum).FixesItems = Val(Parse$(5))
        
        n = 6
        For z = 1 To 6
            For I = 1 To MAX_TRADES
                
                GiveItem = Val(Parse$(n))
                GiveValue = Val(Parse$(n + 1))
                GetItem = Val(Parse$(n + 2))
                GetValue = Val(Parse$(n + 3))
                
                Shop(ShopNum).TradeItem(z).Value(I).GiveItem = GiveItem
                Shop(ShopNum).TradeItem(z).Value(I).GiveValue = GiveValue
                Shop(ShopNum).TradeItem(z).Value(I).GetItem = GetItem
                Shop(ShopNum).TradeItem(z).Value(I).GetValue = GetValue
                
                n = n + 4
            Next I
        Next z
        
        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Main editor packet ::
    ' ::::::::::::::::::::::::
    If (Parse$(0) = MAINEDITOR_CHAR) Then
    
        If LCase$(Dir(App.Path & "\Scripts", vbDirectory)) <> "scripts" Then
            Call MkDir(App.Path & "\Scripts")
        End If
        
        AFileName = "Scripts\Main.txt"
             
        f = FreeFile
        Open App.Path & "\" & AFileName For Output As #f
            Print #f, Parse$(1)
        Close #f
        
        Unload frmEditor
        frmEditor.Show
    End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    If (Parse$(0) = SPELLEDITOR_CHAR) Then
        InSpellEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For I = 1 To MAX_SPELLS
            frmIndex.lstIndex.AddItem I & ": " & Trim$(Spell(I).Name)
        Next I
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (Parse$(0) = UPDATESPELL_CHAR) Then
        n = Val(Parse$(1))
        
        ' Update the spell name
        Spell(n).Name = Parse$(2)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (Parse$(0) = EDITSPELL_CHAR) Then
        n = Val(Parse$(1))
        
        ' Update the spell
        Spell(n).Name = Parse$(2)
        Spell(n).ClassReq = Val(Parse$(3))
        Spell(n).LevelReq = Val(Parse$(4))
        Spell(n).Type = Val(Parse$(5))
        Spell(n).Data1 = Val(Parse$(6))
        Spell(n).Data2 = Val(Parse$(7))
        Spell(n).Data3 = Val(Parse$(8))
        Spell(n).MPCost = Val(Parse$(9))
        Spell(n).Sound = Val(Parse$(10))
        Spell(n).Range = Val(Parse$(11))
        Spell(n).SpellAnim = Val(Parse$(12))
        Spell(n).SpellTime = Val(Parse$(13))
        Spell(n).SpellDone = Val(Parse$(14))
        Spell(n).AE = Val(Parse$(15))
                        
        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (Parse$(0) = TRADE_CHAR) Then
        ShopNum = Val(Parse$(1))
        If Val(Parse$(2)) = 1 Then
            frmTrade.picFixItems.Visible = True
        Else
            frmTrade.picFixItems.Visible = False
        End If
        
        n = 3
        For z = 1 To 6
            For I = 1 To MAX_TRADES
                GiveItem = Val(Parse$(n))
                GiveValue = Val(Parse$(n + 1))
                GetItem = Val(Parse$(n + 2))
                GetValue = Val(Parse$(n + 3))
                
                Trade(z).Items(I).ItemGetNum = GetItem
                Trade(z).Items(I).ItemGiveNum = GiveItem
                Trade(z).Items(I).ItemGetVal = GetValue
                Trade(z).Items(I).ItemGiveVal = GiveValue
                
                n = n + 4
            Next I
        Next z
        
        Dim xx As Long
        For xx = 1 To 6
            Trade(xx).Selected = NO
        Next xx
        
        Trade(1).Selected = YES
                    
        frmTrade.shopType.Top = frmTrade.Label1.Top
        frmTrade.shopType.Left = frmTrade.Label1.Left
        frmTrade.shopType.Height = frmTrade.Label1.Height
        frmTrade.shopType.Width = frmTrade.Label1.Width
        Trade(1).SelectedItem = 1
        
        Call ItemSelected(1, 1)
            
        frmTrade.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Start Speech ::
    ' ::::::::::::::::::
    If Parse$(0) = STARTSPEECH_CHAR Then
        ' should work!
        If frmTalk.Visible = True Then Unload frmTalk
        
        I = Val(Parse$(1))
        n = Val(Parse$(2))
        
        frmTalk.txtActual.Caption = Speech(I).Num(n).Text

        frmTalk.txtActual.Left = frmTalk.Picture4.Left + frmTalk.Picture4.Width + 16
        
        If Speech(I).Num(n).Exit = 0 Then
            If Speech(I).Num(n).Respond > 0 Then
                frmTalk.lblChoice(0).Caption = Speech(I).Num(n).Responces(1).Text
            Else
                frmTalk.lblChoice(0).Caption = vbNullString
            End If
        
            If Speech(I).Num(n).Respond > 1 Then
                frmTalk.lblChoice(1).Caption = Speech(I).Num(n).Responces(2).Text
            Else
                frmTalk.lblChoice(1).Caption = vbNullString
            End If
        
            If Speech(I).Num(n).Respond > 2 Then
                frmTalk.lblChoice(2).Caption = Speech(I).Num(n).Responces(3).Text
            Else
                frmTalk.lblChoice(2).Caption = vbNullString
            End If
        Else
            frmTalk.lblChoice(0).Caption = vbNullString
            frmTalk.lblChoice(1).Caption = vbNullString
            frmTalk.lblChoice(2).Caption = vbNullString
            frmTalk.lblQuit.Caption = "Done"
        End If
        
        SpeechConvo1 = I
        SpeechConvo2 = n
        SpeechConvo3 = Val(Parse$(3))
        
        frmTalk.Show vbModeless, frmMirage
        'frmTalk.Show vbModal
    End If
    
    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (Parse$(0) = SPELLS_CHAR) Then
        
        frmMirage.picPlayerSpells.Visible = True
        frmMirage.lstSpells.Clear
        
        ' Put spells known in player record
        For I = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(I) = Val(Parse$(I))
            If Player(MyIndex).Spell(I) <> 0 Then
                frmMirage.lstSpells.AddItem I & ": " & Trim$(Spell(Player(MyIndex).Spell(I)).Name)
            Else
                frmMirage.lstSpells.AddItem "<free spells slot>"
            End If
        Next I
        
        frmMirage.lstSpells.ListIndex = 0
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (Parse$(0) = WEATHER_CHAR) Then
        If Val(Parse$(1)) = WEATHER_RAINING And GameWeather <> WEATHER_RAINING Then
            Call AddText("You see drops of rain falling from the sky above!", BrightGreen)
        End If
        If Val(Parse$(1)) = WEATHER_THUNDER And GameWeather <> WEATHER_THUNDER Then
            Call AddText("You see thunder in the sky above!", BrightGreen)
        End If
        If Val(Parse$(1)) = WEATHER_SNOWING And GameWeather <> WEATHER_SNOWING Then
            Call AddText("You see snow falling from the sky above!", BrightGreen)
        End If
        
        If Val(Parse$(1)) = WEATHER_NONE Then
            If GameWeather = WEATHER_RAINING Then
                Call AddText("The rain beings to calm.", BrightGreen)
            ElseIf GameWeather = WEATHER_SNOWING Then
                Call AddText("The snow is melting away.", BrightGreen)
            ElseIf GameWeather = WEATHER_THUNDER Then
                Call AddText("The thunder begins to disapear.", BrightGreen)
            End If
        End If
        GameWeather = Val(Parse$(1))
        RainIntensity = Val(Parse$(2))
        If MAX_RAINDROPS <> RainIntensity Then
            MAX_RAINDROPS = RainIntensity
            ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
            ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
        End If
    End If

    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (Parse$(0) = TIME_CHAR) Then
        GameTime = Val(Parse$(1))
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Get Online List ::
    ' ::::::::::::::::::::::::::
    If Parse$(0) = ONLINELIST_CHAR Then
    frmMirage.lstOnline.Clear
    
        n = 2
        z = Val(Parse$(1))
        For x = n To (z + 1)
            frmMirage.lstOnline.AddItem Trim$(Parse$(n))
            n = n + 2
        Next x
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
    If Parse$(0) = BLITPLAYERDMG_CHAR Then
        DmgDamage = Val(Parse$(1))
        NPCWho = Val(Parse$(2))
        DmgTime = GetTickCount
        iii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
    If Parse$(0) = BLITNPCDMG_CHAR Then
        NPCDmgDamage = Val(Parse$(1))
        NPCDmgTime = GetTickCount
        ii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::
    ' :: Retrieve the player's inventory ::
    ' :::::::::::::::::::::::::::::::::::::
    If Parse$(0) = PPTRADING_CHAR Then
        frmPlayerTrade.Items1.Clear
        frmPlayerTrade.Items2.Clear
        For I = 1 To MAX_PLAYER_TRADES
            Trading(I).InvNum = 0
            Trading(I).InvName = vbNullString
            Trading2(I).InvNum = 0
            Trading2(I).InvName = vbNullString
            frmPlayerTrade.Items1.AddItem I & ": <Nothing>"
            frmPlayerTrade.Items2.AddItem I & ": <Nothing>"
        Next I
        
        frmPlayerTrade.Items1.ListIndex = 0
        
        Call UpdateTradeInventory
        frmPlayerTrade.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If Parse$(0) = QTRADE_CHAR Then
        For I = 1 To MAX_PLAYER_TRADES
            Trading(I).InvNum = 0
            Trading(I).InvName = vbNullString
            Trading2(I).InvNum = 0
            Trading2(I).InvName = vbNullString
        Next I
        
        frmPlayerTrade.Command1.ForeColor = &H0&
        frmPlayerTrade.Command2.ForeColor = &H0&
        
        frmPlayerTrade.Visible = False
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If Parse$(0) = UPDATETRADEITEM_CHAR Then
            n = Val(Parse$(1))
            
            Trading2(n).InvNum = Val(Parse$(2))
            Trading2(n).InvName = Parse$(3)
            
            If STR(Trading2(n).InvNum) <= 0 Then
                frmPlayerTrade.Items2.List(n - 1) = n & ": <Nothing>"
            Else
                frmPlayerTrade.Items2.List(n - 1) = n & ": " & Trim$(Trading2(n).InvName)
            End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If Parse$(0) = TRADING_CHAR Then
        n = Val(Parse$(1))
            If n = 0 Then frmPlayerTrade.Command2.ForeColor = &H0&
            If n = 1 Then frmPlayerTrade.Command2.ForeColor = &HFF00&
        Exit Sub
    End If
    
' :::::::::::::::::::::::::
' :: Chat System Packets ::
' :::::::::::::::::::::::::
    If Parse$(0) = PPCHATTING_CHAR Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Label1.Caption = "Chatting With: " & Trim$(Player(Val(Parse$(1))).Name)

        frmPlayerChat.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    If Parse$(0) = QCHAT_CHAR Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Visible = False
        frmPlayerTrade.Command2.ForeColor = &H8000000F
        frmPlayerTrade.Command1.ForeColor = &H8000000F
        Exit Sub
    End If
    
    If Parse$(0) = SENDCHAT_CHAR Then
        Dim s As String
  
        s = vbNewLine & GetPlayerName(Val(Parse$(2))) & "> " & Trim$(Parse$(1))
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text)
        frmPlayerChat.txtChat.SelColor = QBColor(Brown)
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
    If Parse$(0) = SOUND_CHAR Then
    Dim Sound As String
    Dim Sound_ID As Byte
    
        Sound_ID = Parse$(1)
        
        If IsNumeric(Sound_ID) = False Then
            Call PlaySound(Sound_ID & ".wav")
            Exit Sub
        Else
            Sound_ID = Val(Sound_ID)
        End If
        
        If Sound_ID = MAGIC_SOUND Then
            Call PlaySound("Magic" & Parse$(2))
            Exit Sub
        End If
        
        Select Case Sound_ID
        
            Case ATTACK_SOUND
                Sound = "Attack" & Int(Rnd * 7) + 1
            Case CRITICALHIT_SOUND
                Sound = "CriticalHit"
            Case SERVERSHUTDOWN_SOUND
                Sound = "TheServerIsShuttingDown"
            Case DEAD_SOUND
                Sound = "Dead"
            Case PAIN_SOUND
                Sound = "Pain"
            Case NEWLEVEL_SOUND
                Sound = "Levelup"
            Case LOGINTOSERVER_SOUND
                Sound = "LoggingIntoServer"
            Case PLAYERJOINED_SOUND
                Sound = "NewPlayerJoined"
            Case LOGOUTOFSERVER_SOUND
                Sound = "LoggingOutOfServer"
            Case PLAYERHASLEFT_SOUND
                Sound = "PlayerHasLeft"
            Case KEY_SOUND
                Sound = "Key"
            Case WARP_SOUND
                Sound = "Warp"
            Case NEWVERSIONRELEASED_SOUND
                Sound = "NewVersionReleased"
            Case MISS_SOUND
                Sound = "miss"
            Case STRENGTHRAISED_SOUND
                Sound = "StrengthRaised"
            Case DEFENSERAISED_SOUND
                Sound = "DefenseRaised"
            Case MAGICRAISED_SOUND
                Sound = "MagicRaised"
            Case SPEEDRAISED_SOUND
                Sound = "SpeedRaised"
            Case Else
                Exit Sub
        End Select
        
        If Len(Sound) <> 0 Then
            Call PlaySound(Sound & ".wav")
        End If
        
        Sound = vbNullString
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Sprite Change Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If Parse$(0) = SPRITECHANGE_CHAR Then
        If Val(Parse$(1)) = 1 Then
            I = MsgBox("Would you like to buy this sprite?", 4, "Buying Sprite")
            If I = 6 Then
                Call SendData(BUYSPRITE_CHAR & END_CHAR)
            End If
        Else
            Call SendData(BUYSPRITE_CHAR & END_CHAR)
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Change Player Direction Packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If Parse$(0) = CHANGEDIR_CHAR Then
        Player(Val(Parse$(2))).Dir = Val(Parse$(1))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::
    ' :: Change Pet Direction Packet ::
    ' :::::::::::::::::::::::::::::::::
    If Parse$(0) = CHANGEPETDIR_CHAR Then
        Player(Val(Parse$(2))).Pet.Dir = Val(Parse$(1))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Flash Movie Event Packet ::
    ' ::::::::::::::::::::::::::::::
    If Parse$(0) = FLASHEVENT_CHAR Then
        If LCase$(Mid(Trim$(Parse$(1)), 1, 7)) = "http://" Then
            Call PutVar(App.Path & "\config.ini", "CONFIG", "Music", 0)
            Call PutVar(App.Path & "\config.ini", "CONFIG", "Sound", 0)
            MusicOn = NO
            SoundOn = NO
            Call StopMidi
            Call StopSound
            frmFlash.Flash.LoadMovie 0, Trim$(Parse$(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        ElseIf FileExist("Flashs\" & Trim$(Parse$(1))) = True Then
            Call PutVar(App.Path & "\config.ini", "CONFIG", "Music", 0)
            MusicOn = NO
            SoundOn = NO
            Call PutVar(App.Path & "\config.ini", "CONFIG", "Sound", 0)
            Call StopMidi
            Call StopSound
            frmFlash.Flash.LoadMovie 0, App.Path & "\Flashs\" & Trim$(Parse$(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If Parse$(0) = PROMPT_CHAR Then
        I = MsgBox(Trim$(Parse$(1)), vbYesNo)
        Call SendData(PROMPT_CHAR & SEP_CHAR & I & SEP_CHAR & Val(Parse$(2)) & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Speecj editor packet ::
    ' ::::::::::::::::::::::::::
    If (Parse$(0) = SPEECHEDITOR_CHAR) Then
        InSpeechEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        For I = 1 To MAX_SPEECH
            frmIndex.lstIndex.AddItem I & ": " & Trim$(Speech(I).Name)
        Next I
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    If (Parse$(0) = SPEECH_CHAR) Then
        n = Val(Parse$(1))

        Speech(n).Name = Parse$(2)
        
        Dim p, O As Long
        p = 3
        
        For I = 0 To MAX_SPEECH_OPTIONS
            Speech(n).Num(I).Exit = Val(Parse$(p))
            Speech(n).Num(I).Text = Parse$(p + 1)
            Speech(n).Num(I).SaidBy = Val(Parse$(p + 2))
            Speech(n).Num(I).Respond = Val(Parse$(p + 3))
            Speech(n).Num(I).Script = Val(Parse$(p + 4))
            
            p = p + 5
            For O = 1 To 3
                Speech(n).Num(I).Responces(O).Exit = Val(Parse$(p))
                Speech(n).Num(I).Responces(O).GoTo = Val(Parse$(p + 1))
                Speech(n).Num(I).Responces(O).Text = Parse$(p + 2)
                p = p + 3
            Next O
        Next I
        Exit Sub
    End If
    
    If (Parse$(0) = EDITSPEECH_CHAR) Then
        n = Val(Parse$(1))

        Speech(n).Name = Parse$(2)
        
        p = 3
        
        For I = 0 To MAX_SPEECH_OPTIONS
            Speech(n).Num(I).Exit = Val(Parse$(p))
            Speech(n).Num(I).Text = Parse$(p + 1)
            Speech(n).Num(I).SaidBy = Val(Parse$(p + 2))
            Speech(n).Num(I).Respond = Val(Parse$(p + 3))
            Speech(n).Num(I).Script = Val(Parse$(p + 4))
            
            p = p + 5
            For O = 1 To 3
                Speech(n).Num(I).Responces(O).Exit = Val(Parse$(p))
                Speech(n).Num(I).Responces(O).GoTo = Val(Parse$(p + 1))
                Speech(n).Num(I).Responces(O).Text = Parse$(p + 2)
                p = p + 3
            Next O
        Next I
        
        Call SpeechEditorInit
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Emoticon editor packet ::
    ' ::::::::::::::::::::::::::::
    If (Parse$(0) = EMOTICONEDITOR_CHAR) Then
        InEmoticonEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        For I = 0 To MAX_EMOTICONS
            frmIndex.lstIndex.AddItem I & ": " & Trim$(Emoticons(I).Command)
        Next I
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    If (Parse$(0) = UPDATEEMOTICON_CHAR) Then
        n = Val(Parse$(1))
        
        Emoticons(n).Type = Val(Parse$(2))
        Emoticons(n).Command = Parse$(3)
        Emoticons(n).Pic = Val(Parse$(4))
        Emoticons(n).Sound = Parse$(5)
        Exit Sub
    End If

    If (Parse$(0) = EDITEMOTICON_CHAR) Then
        n = Val(Parse$(1))

        Emoticons(n).Type = Val(Parse$(2))
        Emoticons(n).Command = Parse$(3)
        Emoticons(n).Pic = Val(Parse$(4))
        Emoticons(n).Sound = Parse$(5)
        
        Call EmoticonEditorInit
        Exit Sub
    End If
    
    If (Parse$(0) = CLEARTEMPTILE_CHAR) Then
        Call ClearTempTile
        Exit Sub
    End If
    
    If Parse$(0) = FRIENDLIST_CHAR Then
        frmMirage.lstFriend.Clear
        
        n = 1
        frmMirage.lstFriend.AddItem "Online:"
        Do While Parse$(n) <> NEXT_CHAR
            If Trim$(Parse$(n)) <> vbNullString Then frmMirage.lstFriend.AddItem Parse$(n)
            n = n + 1
        Loop
        
        frmMirage.lstFriend.AddItem " "
        frmMirage.lstFriend.AddItem "All friends:"
        
        n = n + 1
        
        Do While Parse$(n) <> NEXT_CHAR
            frmMirage.lstFriend.AddItem Parse$(n)
            n = n + 1
        Loop
        Exit Sub
    End If
    
    If (Parse$(0) = UPDATEEMOTICON_CHAR) Then
        n = Val(Parse$(1))
        
        Emoticons(n).Type = Val(Parse$(2))
        Emoticons(n).Command = Parse$(3)
        Emoticons(n).Pic = Val(Parse$(4))
        Emoticons(n).Sound = Parse$(5)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Arrow editor packet ::
    ' ::::::::::::::::::::::::::::
    If (Parse$(0) = ARROWEDITOR_CHAR) Then
        InArrowEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        For I = 1 To MAX_ARROWS
            frmIndex.lstIndex.AddItem I & ": " & Trim$(Arrows(I).Name)
        Next I
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    If (Parse$(0) = UPDATEARROW_CHAR) Then
        n = Val(Parse$(1))
        
        Arrows(n).Name = Parse$(2)
        Arrows(n).Pic = Val(Parse$(3))
        Arrows(n).Range = Val(Parse$(4))
        Exit Sub
    End If

    If (Parse$(0) = EDITARROW_CHAR) Then
        n = Val(Parse$(1))

        Arrows(n).Name = Parse$(2)
        
        Call ArrowEditorInit
        Exit Sub
    End If
    
    If (Parse$(0) = UPDATEARROW_CHAR) Then
        n = Val(Parse$(1))
        
        Arrows(n).Name = Parse$(2)
        Arrows(n).Pic = Val(Parse$(3))
        Arrows(n).Range = Val(Parse$(4))
        Exit Sub
    End If

    If (Parse$(0) = CHECKARROWS_CHAR) Then
        n = GetPlayerMap(Val(Parse$(1)))
        z = Val(Parse$(2))
        I = Val(Parse$(3))
       
        For x = 1 To MAX_PLAYER_ARROWS
            If Map(n).Arrow(x).Arrow = 0 Then
                Map(n).Arrow(x).Arrow = 1
                Map(n).Arrow(x).ArrowNum = z
                Map(n).Arrow(x).ArrowAnim = Arrows(z).Pic
                Map(n).Arrow(x).ArrowTime = GetTickCount
                Map(n).Arrow(x).ArrowVarX = 0
                Map(n).Arrow(x).ArrowVarY = 0
                Map(n).Arrow(x).ArrowOwner = Val(Parse$(1))
                Map(n).Arrow(x).ArrowY = GetPlayerY(Map(n).Arrow(x).ArrowOwner)
                Map(n).Arrow(x).ArrowX = GetPlayerX(Map(n).Arrow(x).ArrowOwner)
                               
                If I = DIR_DOWN Then
                    Map(n).Arrow(x).ArrowY = GetPlayerY(Map(n).Arrow(x).ArrowOwner) + 1
                    Map(n).Arrow(x).ArrowPosition = 0
                    If Map(n).Arrow(x).ArrowY - 1 > MAX_MAPY Then
                        Map(n).Arrow(x).Arrow = 0
                        Exit Sub
                    End If
                End If
                If I = DIR_UP Then
                    Map(n).Arrow(x).ArrowY = GetPlayerY(Map(n).Arrow(x).ArrowOwner) - 1
                    Map(n).Arrow(x).ArrowPosition = 1
                    If Map(n).Arrow(x).ArrowY + 1 < 0 Then
                        Map(n).Arrow(x).Arrow = 0
                        Exit Sub
                    End If
                End If
                If I = DIR_RIGHT Then
                    Map(n).Arrow(x).ArrowX = GetPlayerX(Map(n).Arrow(x).ArrowOwner) + 1
                    Map(n).Arrow(x).ArrowPosition = 2
                    If Map(n).Arrow(x).ArrowX - 1 > MAX_MAPX Then
                        Map(n).Arrow(x).Arrow = 0
                        Exit Sub
                    End If
                End If
                If I = DIR_LEFT Then
                    Map(n).Arrow(x).ArrowX = GetPlayerX(Map(n).Arrow(x).ArrowOwner) - 1
                    Map(n).Arrow(x).ArrowPosition = 3
                    If Map(n).Arrow(x).ArrowX + 1 < 0 Then
                        Map(n).Arrow(x).Arrow = 0
                        Exit Sub
                    End If
                End If
                Exit For
            End If
        Next x
        Exit Sub
    End If

    If (Parse$(0) = CHECKSPRITE_CHAR) Then
        n = Val(Parse$(1))
        
        Player(n).Sprite = Val(Parse$(2))
        Exit Sub
    End If
    
    If (Parse$(0) = MAPREPORT_CHAR) Then
        n = 1
        
        frmMapReport.lstIndex.Clear
        For I = 1 To MAX_MAPS
            frmMapReport.lstIndex.AddItem I & ": " & Trim$(Parse$(n))
            n = n + 1
        Next I
        
        frmMapReport.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (Parse$(0) = TIME_CHAR) Then
        GameTime = Val(Parse$(1))
        If GameTime = TIME_DAY Then
            Call AddText("Day has dawned in this realm.", White)
        Else
            Call AddText("Night has fallen upon the weary eyed nightowls.", White)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Spell anim packet ::
    ' :::::::::::::::::::::::
    If (Parse$(0) = SPELLANIM_CHAR) Then
        Dim SpellNum As Long
        SpellNum = Val(Parse$(1))
        
        Spell(SpellNum).SpellAnim = Val(Parse$(2))
        Spell(SpellNum).SpellTime = Val(Parse$(3))
        Spell(SpellNum).SpellDone = Val(Parse$(4))
        
        Player(Val(Parse$(5))).SpellNum = SpellNum
        
        For I = 1 To MAX_SPELL_ANIM
            If Player(Val(Parse$(5))).SpellAnim(I).CastedSpell = NO Then
                Player(Val(Parse$(5))).SpellAnim(I).SpellDone = 0
                Player(Val(Parse$(5))).SpellAnim(I).SpellVar = 0
                Player(Val(Parse$(5))).SpellAnim(I).SpellTime = GetTickCount
                Player(Val(Parse$(5))).SpellAnim(I).TargetType = Val(Parse$(6))
                Player(Val(Parse$(5))).SpellAnim(I).Target = Val(Parse$(7))
                Player(Val(Parse$(5))).SpellAnim(I).CastedSpell = YES
                Exit For
            End If
        Next I
        Exit Sub
    End If
    
    If (Parse$(0) = CHECKEMOTICONS_CHAR) Then
        n = Val(Parse$(1))
        
        Player(n).EmoticonType = Val(Parse$(2))
        Player(n).EmoticonNum = Val(Parse$(3))
        Player(n).EmoticonSound = Parse$(4)
        Player(n).EmoticonTime = GetTickCount
        Player(n).EmoticonVar = 0
        Player(n).EmoticonPlayed = False
        Exit Sub
    End If
    
    If (Parse$(0) = TEMPTILE_CHAR) Then
        If Val(Parse$(5)) <> 0 Then
            TempTile(Val(Parse$(1)), Val(Parse$(2))).Ground = 0
            TempTile(Val(Parse$(1)), Val(Parse$(2))).Mask = 0
            TempTile(Val(Parse$(1)), Val(Parse$(2))).Anim = 0
            TempTile(Val(Parse$(1)), Val(Parse$(2))).Mask2 = 0
            TempTile(Val(Parse$(1)), Val(Parse$(2))).M2Anim = 0
            TempTile(Val(Parse$(1)), Val(Parse$(2))).Fringe = 0
            TempTile(Val(Parse$(1)), Val(Parse$(2))).FAnim = 0
            TempTile(Val(Parse$(1)), Val(Parse$(2))).Fringe2 = 0
            TempTile(Val(Parse$(1)), Val(Parse$(2))).F2Anim = 0
        End If
        
        Parse$(4) = Trim$(LCase$(Parse$(4)))
        If Parse$(4) = "ground" Then TempTile(Val(Parse$(1)), Val(Parse$(2))).Ground = Val(Parse$(3))
        If Parse$(4) = "mask" Then TempTile(Val(Parse$(1)), Val(Parse$(2))).Mask = Val(Parse$(3))
        If Parse$(4) = "anim" Then TempTile(Val(Parse$(1)), Val(Parse$(2))).Anim = Val(Parse$(3))
        If Parse$(4) = "mask2" Then TempTile(Val(Parse$(1)), Val(Parse$(2))).Mask2 = Val(Parse$(3))
        If Parse$(4) = "m2anim" Then TempTile(Val(Parse$(1)), Val(Parse$(2))).M2Anim = Val(Parse$(3))
        If Parse$(4) = "fringe" Then TempTile(Val(Parse$(1)), Val(Parse$(2))).Fringe = Val(Parse$(3))
        If Parse$(4) = "fanim" Then TempTile(Val(Parse$(1)), Val(Parse$(2))).FAnim = Val(Parse$(3))
        If Parse$(4) = "fringe2" Then TempTile(Val(Parse$(1)), Val(Parse$(2))).Fringe2 = Val(Parse$(3))
        If Parse$(4) = "f2anim" Then TempTile(Val(Parse$(1)), Val(Parse$(2))).F2Anim = Val(Parse$(3))
        Exit Sub
    End If
    
    If (Parse$(0) = TEMPATTRIBUTE_CHAR) Then
        TempTile(Val(Parse$(1)), Val(Parse$(2))).Type = Val(Parse$(3))
        TempTile(Val(Parse$(1)), Val(Parse$(2))).Data1 = Val(Parse$(4))
        TempTile(Val(Parse$(1)), Val(Parse$(2))).Data2 = Val(Parse$(5))
        TempTile(Val(Parse$(1)), Val(Parse$(2))).Data3 = Val(Parse$(6))
        TempTile(Val(Parse$(1)), Val(Parse$(2))).String1 = Val(Parse$(7))
        TempTile(Val(Parse$(1)), Val(Parse$(2))).String2 = Val(Parse$(8))
        TempTile(Val(Parse$(1)), Val(Parse$(2))).String3 = Val(Parse$(9))
        Exit Sub
    End If
    
    If Parse$(0) = LEVELUP_CHAR Then
        Player(Val(Parse$(1))).LevelUpT = GetTickCount
        Player(Val(Parse$(1))).LevelUp = 1
        Exit Sub
    End If
    
    If Parse$(0) = DAMAGEDISPLAY_CHAR Then
        For I = 1 To MAX_BLT_LINE
            If Val(Parse$(1)) = 0 Then
                If BattlePMsg(I).Index <= 0 Then
                    BattlePMsg(I).Index = 1
                    BattlePMsg(I).Msg = Parse$(2)
                    BattlePMsg(I).Color = Val(Parse$(3))
                    BattlePMsg(I).Time = GetTickCount
                    BattlePMsg(I).done = 1
                    BattlePMsg(I).y = 0
                    Exit Sub
                Else
                    BattlePMsg(I).y = BattlePMsg(I).y - 15
                End If
            Else
                If BattleMMsg(I).Index <= 0 Then
                    BattleMMsg(I).Index = 1
                    BattleMMsg(I).Msg = Parse$(2)
                    BattleMMsg(I).Color = Val(Parse$(3))
                    BattleMMsg(I).Time = GetTickCount
                    BattleMMsg(I).done = 1
                    BattleMMsg(I).y = 0
                    Exit Sub
                Else
                    BattleMMsg(I).y = BattleMMsg(I).y - 15
                End If
            End If
        Next I
        
        z = 1
        If Val(Parse$(1)) = 0 Then
            For I = 1 To MAX_BLT_LINE
                If I < MAX_BLT_LINE Then
                    If BattlePMsg(I).y < BattlePMsg(I + 1).y Then z = I
                Else
                    If BattlePMsg(I).y < BattlePMsg(1).y Then z = I
                End If
            Next I
                        
            BattlePMsg(z).Index = 1
            BattlePMsg(z).Msg = Parse$(2)
            BattlePMsg(z).Color = Val(Parse$(3))
            BattlePMsg(z).Time = GetTickCount
            BattlePMsg(z).done = 1
            BattlePMsg(z).y = 0
        Else
            For I = 1 To MAX_BLT_LINE
                If I < MAX_BLT_LINE Then
                    If BattleMMsg(I).y < BattleMMsg(I + 1).y Then z = I
                Else
                    If BattleMMsg(I).y < BattleMMsg(1).y Then z = I
                End If
            Next I
                        
            BattleMMsg(z).Index = 1
            BattleMMsg(z).Msg = Parse$(2)
            BattleMMsg(z).Color = Val(Parse$(3))
            BattleMMsg(z).Time = GetTickCount
            BattleMMsg(z).done = 1
            BattleMMsg(z).y = 0
        End If
        Exit Sub
    End If
    
    If Parse$(0) = ITEMBREAK_CHAR Then
        ItemDur(Val(Parse$(1))).Item = Val(Parse$(2))
        ItemDur(Val(Parse$(1))).Dur = Val(Parse$(3))
        ItemDur(Val(Parse$(1))).done = 1
        Exit Sub
    End If
End Sub

Function ConnectToServer() As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMirage.Socket.Close
    frmMirage.Socket.Connect
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 2000)
        DoEvents
    Loop
    
    If IsConnected Then
        ConnectToServer = True
    Else
        ConnectToServer = False
    End If
End Function

Function IsConnected() As Boolean
    If frmMirage.Socket.State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If GetPlayerName(Index) <> vbNullString Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Sub SendData(ByVal Data As String)
    If IsConnected Then
        frmMirage.Socket.SendData Data
        DoEvents
    End If
End Sub

Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim Packet As String

    Packet = NEWFACCOUNTIED_CHAR & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
Dim Packet As String
    
    Packet = DELIMACCOUNTED_CHAR & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim Packet As String

    Packet = LOGINATION_CHAR & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & CLIENT_MAJOR & SEP_CHAR & CLIENT_MINOR & SEP_CHAR & CLIENT_REVISION & SEP_CHAR & SEC_CODE & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
Dim Packet As String

    Packet = ADDACHARA_CHAR & SEP_CHAR & Trim$(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & Slot & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelChar(ByVal Slot As Long)
Dim Packet As String
    
    Packet = DELIMBOCHARU_CHAR & SEP_CHAR & Slot & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGetClasses()
Dim Packet As String

    Packet = GATGLASSES_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseChar(ByVal CharSlot As Long)
Dim Packet As String

    Packet = USAGAKARIM_CHAR & SEP_CHAR & CharSlot & END_CHAR
    Call SendData(Packet)
End Sub

Sub SayMsg(ByVal Text As String)
Dim Packet As String

    Packet = SAYMSG_CHAR & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub GlobalMsg(ByVal Text As String)
Dim Packet As String

    Packet = GLOBALMSG_CHAR & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub BroadcastMsg(ByVal Text As String)
Dim Packet As String

    Packet = BROADCASTMSG_CHAR & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub EmoteMsg(ByVal Text As String)
Dim Packet As String

    Packet = EMOTEMSG_CHAR & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub PartyMsg(ByVal Text As String)
Dim Packet As String

    Packet = PARTYCHAT_CHAR & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub GuildMsg(ByVal Text As String)
Dim Packet As String

    Packet = GUILDCHAT_CHAR & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub MapMsg(ByVal Text As String)
Dim Packet As String

    Packet = MAPMSG_CHAR & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Packet As String

    Packet = PLAYERMSG_CHAR & SEP_CHAR & MsgTo & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub AdminMsg(ByVal Text As String)
Dim Packet As String

    Packet = ADMINMSG_CHAR & SEP_CHAR & Text & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerMove()
Dim Packet As String

    Packet = PLAYERMOVE_CHAR & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerDir()
Dim Packet As String

    Packet = PLAYERDIR_CHAR & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerRequestNewMap()
Dim Packet As String
    
    Packet = REQUESTNEWMAP_CHAR & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendMap()
Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long
Dim I As Long
Dim O As Long
Dim MapNum As Long

    MapNum = GetPlayerMap(MyIndex)

    Packet = MAPDATA_CHAR & SEP_CHAR & MapNum & SEP_CHAR & Trim$(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
        With Map(MapNum).Tile(x, y)
            I = 0
            O = 0
            
            If .Ground <> 0 Then I = 0
            If .GroundSet <> -1 Then I = 1
            If .Mask <> 0 Then I = 2
            If .MaskSet <> -1 Then I = 3
            If .Anim <> 0 Then I = 4
            If .AnimSet <> -1 Then I = 5
            If .Fringe <> 0 Then I = 6
            If .FringeSet <> -1 Then I = 7
            If .Type <> 0 Then I = 8
            If .Data1 <> 0 Then I = 9
            If .Data2 <> 0 Then I = 10
            If .Data3 <> 0 Then I = 11
            If .String1 <> vbNullString Then I = 12
            If .String2 <> vbNullString Then I = 13
            If .String3 <> vbNullString Then I = 14
            If .Mask2 <> 0 Then I = 15
            If .Mask2Set <> -1 Then I = 16
            If .M2Anim <> 0 Then I = 17
            If .M2AnimSet <> -1 Then I = 18
            If .FAnim <> 0 Then I = 19
            If .FAnimSet <> -1 Then I = 20
            If .Fringe2 <> 0 Then I = 21
            If .Fringe2Set <> -1 Then I = 22
            If .Light <> 0 Then I = 23
            If .F2Anim <> 0 Then I = 24
            If .F2AnimSet <> -1 Then I = 25
            
            Packet = Packet & .Ground & SEP_CHAR
            If O < I Then
                O = O + 1
                Packet = Packet & .GroundSet & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Mask & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .MaskSet & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Anim & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .AnimSet & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Fringe & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .FringeSet & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Type & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Data1 & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Data2 & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Data3 & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .String1 & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .String2 & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .String3 & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Mask2 & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Mask2Set & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .M2Anim & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .M2AnimSet & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .FAnim & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .FAnimSet & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Fringe2 & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Fringe2Set & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .Light & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .F2Anim & SEP_CHAR
            End If
            If O < I Then
                O = O + 1
                Packet = Packet & .F2AnimSet & SEP_CHAR
            End If
            Packet = Packet & NEXT_CHAR & SEP_CHAR
        End With
        Next x
    Next y
    
    For x = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(MapNum).Npc(x) & SEP_CHAR
        Packet = Packet & Map(MapNum).NpcSpawn(x).Used & SEP_CHAR & Map(MapNum).NpcSpawn(x).x & SEP_CHAR & Map(MapNum).NpcSpawn(x).y & SEP_CHAR
    Next x
        
    Packet = Packet & END_CHAR

    Call SendData(Packet)
End Sub

Sub WarpMeTo(ByVal Name As String)
Dim Packet As String

    Packet = WARPPLAYER_CHAR & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpToMe(ByVal Name As String)
Dim Packet As String

    Packet = WARPTOME_CHAR & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub WarpTo(ByVal MapNum As Long)
Dim Packet As String
    
    Packet = WARPTO_CHAR & SEP_CHAR & MapNum & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim Packet As String

    Packet = SETACCESS_CHAR & SEP_CHAR & Name & SEP_CHAR & Access & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSetSprite(ByVal SpriteNum As Long)
Dim Packet As String

    Packet = SETSPRITE_CHAR & SEP_CHAR & SpriteNum & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendKick(ByVal Name As String)
Dim Packet As String

    Packet = KICKPLAYER_CHAR & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBan(ByVal Name As String)
Dim Packet As String

    Packet = BANPLAYER_CHAR & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBanList()
Dim Packet As String

    Packet = BANLIST_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditItem()
Dim Packet As String

    Packet = REQUESTEDITITEM_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveItem(ByVal ItemNum As Long)
Dim Packet As String

    Packet = SAVEITEM_CHAR & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).desc & SEP_CHAR & Item(ItemNum).AttackSpeed
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditSpeech()
Dim Packet As String

    Packet = REQUESTEDITSPEECH_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveSpeech(ByVal SpcNum As Long)
Dim Packet As String
Dim I, O As Long

    Packet = SAVESPEECH_CHAR & SEP_CHAR & SpcNum & SEP_CHAR & Speech(SpcNum).Name & SEP_CHAR
    For I = 0 To MAX_SPEECH_OPTIONS
        Packet = Packet & Speech(SpcNum).Num(I).Exit & SEP_CHAR & Speech(SpcNum).Num(I).Text & SEP_CHAR & Speech(SpcNum).Num(I).SaidBy & SEP_CHAR & Speech(SpcNum).Num(I).Respond & SEP_CHAR & Speech(SpcNum).Num(I).Script & SEP_CHAR
        For O = 1 To 3
            Packet = Packet & Speech(SpcNum).Num(I).Responces(O).Exit & SEP_CHAR & Speech(SpcNum).Num(I).Responces(O).GoTo & SEP_CHAR & Speech(SpcNum).Num(I).Responces(O).Text & SEP_CHAR
        Next O
    Next I
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub
                
Sub SendRequestEditEmoticon()
Dim Packet As String

    Packet = REQUESTEDITEMOTICON_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveEmoticon(ByVal EmoNum As Long)
Dim Packet As String

    Packet = SAVEEMOTICON_CHAR & SEP_CHAR & EmoNum & SEP_CHAR & Emoticons(EmoNum).Type & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & Emoticons(EmoNum).Sound & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditArrow()
Dim Packet As String

    Packet = REQUESTEDITARROW_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveArrow(ByVal ArrowNum As Long)
Dim Packet As String

    Packet = SAVEARROW_CHAR & SEP_CHAR & ArrowNum & SEP_CHAR & Trim$(Arrows(ArrowNum).Name) & SEP_CHAR & Arrows(ArrowNum).Pic & SEP_CHAR & Arrows(ArrowNum).Range & END_CHAR
    Call SendData(Packet)
End Sub
                
Sub SendRequestEditNpc()
Dim Packet As String

    Packet = REQUESTEDITNPC_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveNpc(ByVal NpcNum As Long)
Dim Packet As String
Dim I As Long
    
    Packet = SAVENPC_CHAR & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).Speed & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHP & SEP_CHAR & Npc(NpcNum).EXP & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR & Npc(NpcNum).Speech & SEP_CHAR
    For I = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(NpcNum).ItemNPC(I).Chance
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(I).ItemNum
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(I).ItemValue & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendMapRespawn()
Dim Packet As String

    Packet = MAPRESPAWN_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendUseItem(ByVal InvNum As Long)
Dim Packet As String

    Packet = USEITEM_CHAR & SEP_CHAR & InvNum & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDropItem(ByVal InvNum, ByVal Ammount As Long)
Dim Packet As String

    Packet = MAPDROPITEM_CHAR & SEP_CHAR & InvNum & SEP_CHAR & Ammount & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendWhosOnline()
Dim Packet As String

    Packet = WHOSONLINE_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendOnlineList()
Dim Packet As String

Packet = ONLINELIST_CHAR & END_CHAR
Call SendData(Packet)
End Sub
            
Sub SendMOTDChange(ByVal MOTD As String)
Dim Packet As String

    Packet = SETMOTD_CHAR & SEP_CHAR & MOTD & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditShop()
Dim Packet As String

    Packet = REQUESTEDITSHOP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveShop(ByVal ShopNum As Long)
Dim Packet As String
Dim I As Long, z As Long

    Packet = SAVESHOP_CHAR & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For I = 1 To 6
        For z = 1 To MAX_TRADES
            Packet = Packet & Shop(ShopNum).TradeItem(I).Value(z).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(z).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(z).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).Value(z).GetValue & SEP_CHAR
        Next z
    Next I
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditMain()
Dim Packet As String

    Packet = REQUESTEDITMAIN_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestBackupMain()
Dim Packet As String

    Packet = REQUESTBACKUPMAIN_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditSpell()
Dim Packet As String

    Packet = REQUESTEDITSPELL_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveSpell(ByVal SpellNum As Long)
Dim Packet As String

    Packet = SAVESPELL_CHAR & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Trim$(Spell(SpellNum).Sound) & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendWarp(ByVal Where As String)
Dim Packet As String

    Packet = WARPPLAYER_CHAR & SEP_CHAR & Where & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditMap()
Dim Packet As String

    Packet = REQUESTEDITMAP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendTradeRequest(ByVal Name As String)
Dim Packet As String

    Packet = PPTRADE_CHAR & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAcceptTrade()
Dim Packet As String

    Packet = ATRADE_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDeclineTrade()
Dim Packet As String

    Packet = DTRADE_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPartyRequest(ByVal Name As String)
Dim Packet As String

    Packet = PARTY_CHAR & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendNewParty(ByVal Name As String)
Dim Packet As String

    Packet = NEWPARTY_CHAR & SEP_CHAR & Name & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendJoinParty()
Dim Packet As String

    Packet = JOINPARTY_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLeaveParty()
Dim Packet As String

    Packet = LEAVEPARTY_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendBanDestroy()
Dim Packet As String
    
    Packet = BANDESTROY_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestLocation()
Dim Packet As String

    Packet = REQUESTLOCATION_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
Sub SendSetPlayerSprite(ByVal Name As String, ByVal SpriteNum As Byte)
Dim Packet As String

    Packet = SETPLAYERSPRITE_CHAR & SEP_CHAR & Name & SEP_CHAR & SpriteNum & END_CHAR
    Call SendData(Packet)
End Sub
