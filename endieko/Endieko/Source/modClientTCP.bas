Attribute VB_Name = "modClientTCP"
Option Explicit

Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean
Public ItemGiveS(1 To MAX_TRADES) As Long
Public ItemGetS(1 To MAX_TRADES) As Long
Public ItemGiveSS(1 To MAX_TRADES) As Long
Public ItemGetSS(1 To MAX_TRADES) As Long
Public TradePlayer As Long

Sub TcpInit()
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    PlayerBuffer = ""
    
    Dim IP As String
    Dim Port As String
    Dim FileName As String
    FileName = App.Path & "\config.ini"
    If FileExist("config.ini") Then
        IP = ReadINI("IPCONFIG", "IP", FileName)
        Port = ReadINI("IPCONFIG", "PORT", FileName)
    Else
        IP = "127.0.0.1"
        Port = 4000
        WriteINI "IPCONFIG", "IP", IP, (App.Path & "\config.ini")
        WriteINI "IPCONFIG", "PORT", Port, (App.Path & "\config.ini")
        WriteINI "CONFIG", "Account", "", (App.Path & "\config.ini")
        WriteINI "CONFIG", "Password", "", (App.Path & "\config.ini")
        WriteINI "CONFIG", "WebSite", "", (App.Path & "\config.ini")
        WriteINI "CONFIG", "NpcBar", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "NPCName", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "NPCDamage", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "PlayerBar", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "PlayerName", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "PlayerDamage", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "MapGrid", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "Music", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "Sound", 1, App.Path & "\config.ini"
    End If
    frmEndieko.Socket.RemoteHost = IP
    frmEndieko.Socket.RemotePort = Val(Port)
End Sub

Sub TcpDestroy()
    frmEndieko.Socket.Close
    
    If frmChars.Visible Then frmChars.Visible = False
    If frmCredits.Visible Then frmCredits.Visible = False
    If frmDeleteAccount.Visible Then frmDeleteAccount.Visible = False
    If frmLogin.Visible Then frmLogin.Visible = False
    If frmNewAccount.Visible Then frmNewAccount.Visible = False
    If frmNewChar.Visible Then frmNewChar.Visible = False
End Sub

'Sub IncomingData(ByVal DataLength As Long)
'Dim Buffer As String
'Dim Packet As String
'Dim Top As String * 3
'Dim Start As Long
'
'    frmEndieko.Socket.GetData Buffer, vbString, DataLength
'    PlayerBuffer = PlayerBuffer & Buffer
'
'    Start = InStr(PlayerBuffer, END_CHAR)
'    Do While Start > 0
'        Packet = Mid$(PlayerBuffer, 1, Start - 1)
'        PlayerBuffer = Mid$(PlayerBuffer, Start + 1, Len(PlayerBuffer))
'        Start = InStr(PlayerBuffer, END_CHAR)
'        If Len(Packet) > 0 Then
'            Call HandleData(Packet)
'        End If
'    Loop
'End Sub

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim Top As String * 3
Dim Start As Integer
Dim Sploc As Integer
Dim lR As Long

    frmEndieko.Socket.GetData Buffer, vbString, DataLength
    Sploc = InStr(1, Buffer, SEP_CHAR)
        lR = Mid(Buffer, 1, Sploc - 1)
        Buffer = Mid(Buffer, Sploc + 1, Len(Buffer) - Sploc)
        Buffer = Uncompress(Buffer, lR)
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
Dim Dir As Long
Dim InvNum As Long
Dim Ammount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Level As Long
Dim i As Long, n As Long, x As Long, y As Long
Dim ShopNum As Long, GiveItem As Long, GiveValue As Long, GetItem As Long, GetValue As Long
Dim z As Long
Dim Sprite As Integer

    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
    
    ' Add the data to the debug window if we are in debug mode
    If Trim$(Command) = "-debug" Then
        If frmDebug.Visible = False Then frmDebug.Visible = True
        Call TextAdd(frmDebug.txtDebug, "((( Processed Packet " & Parse$(0) & " )))", True)
    End If

' Optimized Function, Select Case
' :::::::::::::::::::::::::
' ::::: Added 4/16/06 :::::
' :::::::::::::::::::::::::
Select Case LCase$(Parse$(0))
    ' :::::::::::::::::::::::
    ' :: Get players stats ::
    ' :::::::::::::::::::::::
    Case "maxinfo"
    'If LCase$(Parse$(0)) = "maxinfo" Then
        For i = 0 To MAX_EMOTICONS
            Emoticons(i).Pic = 0
            Emoticons(i).Command = ""
        Next i
        
        Call ClearTempTile
        
        ' Clear out players
        For i = 1 To MAX_PLAYERS
            Call ClearPlayer(i)
        Next i
    
        App.Title = "Endieko"
 
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::
    ' :: Get players stats ::
    ' :::::::::::::::::::::::
    Case "playershpmp"
    'If LCase$(Parse$(0)) = "playershpmp" Then
        n = 1
 
        For i = 1 To MAX_PLAYERS
            Player(i).HP = Val(Parse$(n))
            Player(i).MaxHp = Val(Parse$(n + 1))
 
            n = n + 2
        Next i
 
        Exit Sub
    'End If
    
    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    Case "npchp"
    'If LCase$(Parse$(0)) = "npchp" Then
        n = 1
       
        For i = 1 To MAX_MAP_NPCS
            MapNpc(i).HP = Val(Parse$(n))
            MapNpc(i).MaxHp = Val(Parse$(n + 1))
            
            n = n + 2
        Next i
        
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    Case "alertmsg"
    'If LCase$(Parse$(0)) = "alertmsg" Then
        frmSendGetData.Visible = False
        frmMainMenu.Visible = True
        
        Msg = Parse$(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
    Case "allchars"
    'If LCase$(Parse$(0)) = "allchars" Then
        n = 1
        
        frmChars.Visible = True
        frmSendGetData.Visible = False
        
        For i = 1 To MAX_CHARS
            Name = Parse$(n)
            Msg = Parse$(n + 1)
            Level = Val(Parse$(n + 2))
            Sprite = Val(Parse$(n + 3))
            
            If Trim$(Name) = "" Then
                frmChars.lblName(i - 1).Caption = "Free Character Slot"
            Else
                frmChars.lblName(i - 1).Caption = Name
                frmChars.lblClass(i - 1).Caption = Msg
                frmChars.lblLevel(i - 1).Caption = Level
            End If
            
            Call BitBlt(frmChars.picSprite.hDC, 0, 0, PIC_X, PIC_Y, frmChars.picSprites.hDC, 3 * PIC_X, Sprite * PIC_Y, SRCCOPY)

            n = n + 4
        Next i
        Exit Sub
    'End If

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    Case "loginok"
    'If LCase$(Parse$(0)) = "loginok" Then
        ' Now we can receive game data
        MyIndex = Val(Parse$(1))
        
        frmSendGetData.Visible = True
        frmChars.Visible = False
        
        Call SetStatus("Receiving game data...")
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    Case "newcharclasses"
    'If LCase$(Parse$(0)) = "newcharclasses" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse$(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1

        For i = 0 To Max_Classes
            Class(i).Name = Parse$(n)
            
            Class(i).HP = Val(Parse$(n + 1))
            Class(i).MP = Val(Parse$(n + 2))
            Class(i).SP = Val(Parse$(n + 3))
            
            Class(i).STR = Val(Parse$(n + 4))
            Class(i).DEF = Val(Parse$(n + 5))
            Class(i).Speed = Val(Parse$(n + 6))
            Class(i).MAGI = Val(Parse$(n + 7))
            'Class(i).INTEL = Val(Parse$(n + 8))
            Class(i).MaleSprite = Val(Parse$(n + 8))
            Class(i).FemaleSprite = Val(Parse$(n + 9))
            Class(i).Locked = Val(Parse$(n + 10))
        
        n = n + 11
        Next i
        
        ' Used for if the player is creating a new character
        frmNewChar.Visible = True
        frmSendGetData.Visible = False

        'frmNewChar.cmbClass.Clear
        'For i = 0 To Max_Classes
            'If Class(i).Locked = 0 Then
                frmNewChar.lblClass.Caption = "Class: " & Trim$(Class(0).Name)
            'End If
        'Next i
        
        'frmNewChar.cmbClass.ListIndex = 0
        frmNewChar.lblHP.Caption = STR$(Class(0).HP)
        frmNewChar.lblMP.Caption = STR$(Class(0).MP)
        frmNewChar.lblSP.Caption = STR$(Class(0).SP)
    
        frmNewChar.lblSTR.Caption = STR$(Class(0).STR)
        frmNewChar.lblDEF.Caption = STR$(Class(0).DEF)
        frmNewChar.lblSPEED.Caption = STR$(Class(0).Speed)
        frmNewChar.lblMAGI.Caption = STR$(Class(0).MAGI)
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    Case "classesdata"
    'If LCase$(Parse$(0)) = "classesdata" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse$(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For i = 0 To Max_Classes
            Class(i).Name = Parse$(n)
            
            Class(i).HP = Val(Parse$(n + 1))
            Class(i).MP = Val(Parse$(n + 2))
            Class(i).SP = Val(Parse$(n + 3))
            
            Class(i).STR = Val(Parse$(n + 4))
            Class(i).DEF = Val(Parse$(n + 5))
            Class(i).Speed = Val(Parse$(n + 6))
            Class(i).MAGI = Val(Parse$(n + 7))
            
            Class(i).Locked = Val(Parse$(n + 8))
            
            n = n + 8
        Next i
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    Case "ingame"
    'If LCase$(Parse$(0)) = "ingame" Then
        InGame = True
        Call GameInit
        Call GameLoop
        If Parse$(1) = END_CHAR Then
            MsgBox ("here")
            End
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    Case "playerinv"
    'If LCase$(Parse$(0)) = "playerinv" Then
        n = 1
        For i = 1 To MAX_INV
            Call SetPlayerInvItemNum(MyIndex, i, Val(Parse$(n)))
            Call SetPlayerInvItemValue(MyIndex, i, Val(Parse$(n + 1)))
            Call SetPlayerInvItemDur(MyIndex, i, Val(Parse$(n + 2)))
            
            n = n + 3
        Next i
        Call UpdateInventory
        Call UpdateVisInv
        Exit Sub
    'End If

    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
    Case "playerinvupdate"
    'If LCase$(Parse$(0)) = "playerinvupdate" Then
        n = Val(Parse$(1))
        
        Call SetPlayerInvItemNum(MyIndex, n, Val(Parse$(2)))
        Call SetPlayerInvItemValue(MyIndex, n, Val(Parse$(3)))
        Call SetPlayerInvItemDur(MyIndex, n, Val(Parse$(4)))
        Call UpdateInventory
        Call UpdateVisInv
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    Case "playerworneq"
    'If LCase$(Parse$(0)) = "playerworneq" Then
        Call SetPlayerArmorSlot(MyIndex, Val(Parse$(1)))
        Call SetPlayerWeaponSlot(MyIndex, Val(Parse$(2)))
        Call SetPlayerHelmetSlot(MyIndex, Val(Parse$(3)))
        Call SetPlayerShieldSlot(MyIndex, Val(Parse$(4)))
        Call SetPlayerLegSlot(MyIndex, Val(Parse$(5)))
        Call SetPlayerBootSlot(MyIndex, Val(Parse$(6)))
        Call UpdateInventory
        Call UpdateVisInv
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player bank update packet ::
    ' :::::::::::::::::::::::::::::::
    Case "playerbankupdate"
    'If LCase$(Parse$(0)) = "playerbankupdate" Then
        n = Val(Parse$(1))
        z = Val(Parse$(2))
        
        Call SetPlayerBankItemNum(z, n, Val(Parse$(3)))
        Call SetPlayerBankItemValue(z, n, Val(Parse$(4)))
        Call SetPlayerBankItemDur(z, n, Val(Parse$(5)))
        If z = MyIndex Then Call UpdateBank
        Exit Sub
    'End If
    
    ' :::::::::::::::
    ' :: Load Bank ::
    ' :::::::::::::::
    Case "loadbank"
    'If LCase$(Parse$(0)) = "loadbank" Then
        frmBank.Show 'vbModal
        Exit Sub
    'End If

        
    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    Case "playerhp"
    'If LCase$(Parse$(0)) = "playerhp" Then
        Player(MyIndex).MaxHp = Val(Parse$(1))
        Call SetPlayerHP(MyIndex, Val(Parse$(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then
            frmEndieko.shpHP.Height = (((GetPlayerHP(MyIndex) / 272) / (GetPlayerMaxHP(MyIndex) / 272)) * 272)
            frmEndieko.shpHP.Top = 291 - frmEndieko.shpHP.Height
            Call GradObj(frmEndieko.shpHP, RGB(231, 93, 93), RGB(83, 2, 0))
        End If
        Exit Sub
    'End If

    ' ::::::::::::::::::::::
    ' :: Player mp packet ::
    ' ::::::::::::::::::::::
    Case "playermp"
    'If LCase$(Parse$(0)) = "playermp" Then
        Player(MyIndex).MaxMP = Val(Parse$(1))
        Call SetPlayerMP(MyIndex, Val(Parse$(2)))
        If GetPlayerMaxMP(MyIndex) > 0 Then
            frmEndieko.shpMP.Height = (((GetPlayerMP(MyIndex) / 272) / (GetPlayerMaxMP(MyIndex) / 272)) * 272)
            frmEndieko.shpMP.Top = 291 - frmEndieko.shpMP.Height
            Call GradObj(frmEndieko.shpMP, RGB(93, 171, 231), RGB(0, 4, 31))
        End If
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    Case "playersp"
    'If LCase$(Parse$(0)) = "playersp" Then
        Player(MyIndex).MaxSP = Val(Parse$(1))
        Call SetPlayerSP(MyIndex, Val(Parse$(2)))
        If GetPlayerMaxSP(MyIndex) > 0 Then
            frmEndieko.shpSp.Height = (((GetPlayerSP(MyIndex) / 272) / (GetPlayerMaxSP(MyIndex) / 272)) * 272)
            frmEndieko.shpSp.Top = 291 - frmEndieko.shpSp.Height
            Call GradObj(frmEndieko.shpSp, RGB(114, 209, 117), RGB(0, 38, 0))
            'frmEndieko.shpSp.Height = frmEndieko.shpSp.Top - (GetPlayerMaxSP(MyIndex) - GetPlayerMaxSP(MyIndex))
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::
    ' :: Player Stats Packet ::
    ' :::::::::::::::::::::::::
    Case "playerstatspacket"
    'If (LCase$(Parse$(0)) = "playerstatspacket") Then
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
        
        'frmEndieko.shpTNL.Width = (((Val(Parse$(6)) / 100) / (Val(Parse$(5)) / 100)) * 100)
        'frmEndieko.shpTNL.Width = (((GetPlayerExp(MyIndex) / 100) / (GetPlayerMaxExp(MyIndex) / 100)) * 100)
        'frmEndieko.lblTNL.Caption = (((Val(Parse$(6)) / 100) & " / " & (Val(Parse$(5)) / 100)) * 100)
        Exit Sub
    'End If
                

    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    Case "playerdata"
    'If LCase$(Parse$(0)) = "playerdata" Then
        i = Val(Parse$(1))
        On Error Resume Next
        Call SetPlayerName(i, Parse$(2))
        Call SetPlayerSprite(i, Val(Parse$(3)))
        Call SetPlayerMap(i, Val(Parse$(4)))
        Call SetPlayerX(i, Val(Parse$(5)))
        Call SetPlayerY(i, Val(Parse$(6)))
        Call SetPlayerDir(i, Val(Parse$(7)))
        Call SetPlayerAccess(i, Val(Parse$(8)))
        Call SetPlayerPK(i, Val(Parse$(9)))
        Call SetPlayerGuild(i, Parse$(10))
        Call SetPlayerGuildAccess(i, Val(Parse$(11)))
        Call SetPlayerClass(i, Val(Parse$(12)))

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
    'End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    Case "playermove"
    'If (LCase$(Parse$(0)) = "playermove") Then
        i = Val(Parse$(1))
        x = Val(Parse$(2))
        y = Val(Parse$(3))
        Dir = Val(Parse$(4))
        n = Val(Parse$(5))

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
    'End If
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    Case "npcmove"
    'If (LCase$(Parse$(0)) = "npcmove") Then
        i = Val(Parse$(1))
        x = Val(Parse$(2))
        y = Val(Parse$(3))
        Dir = Val(Parse$(4))
        n = Val(Parse$(5))

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
    'End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    Case "playerdir"
    'If (LCase$(Parse$(0)) = "playerdir") Then
        i = Val(Parse$(1))
        Dir = Val(Parse$(2))
        Call SetPlayerDir(i, Dir)
        
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).Moving = 0
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    Case "npcdir"
    'If (LCase$(Parse$(0)) = "npcdir") Then
        i = Val(Parse$(1))
        Dir = Val(Parse$(2))
        MapNpc(i).Dir = Dir
        
        MapNpc(i).XOffset = 0
        MapNpc(i).YOffset = 0
        MapNpc(i).Moving = 0
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    Case "playerxy"
    'If (LCase$(Parse$(0)) = "playerxy") Then
        x = Val(Parse$(1))
        y = Val(Parse$(2))
        
        Call SetPlayerX(MyIndex, x)
        Call SetPlayerY(MyIndex, y)
        
        ' Make sure they aren't walking
        Player(MyIndex).Moving = 0
        Player(MyIndex).XOffset = 0
        Player(MyIndex).YOffset = 0
        
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    Case "attack"
    'If (LCase$(Parse$(0)) = "attack") Then
        i = Val(Parse$(1))
        
        ' Set player to attacking
        Player(i).Attacking = 1
        Player(i).AttackTimer = GetTickCount
        
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    Case "npcattack"
    'If (LCase$(Parse$(0)) = "npcattack") Then
        i = Val(Parse$(1))
        
        ' Set player to attacking
        MapNpc(i).Attacking = 1
        MapNpc(i).AttackTimer = GetTickCount
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::
    ' :: Arrow editor packet ::
    ' :::::::::::::::::::::::::
    Case "arroweditor"
    'If (LCase$(Parse$(0)) = "arroweditor") Then
        InArrowEditor = True
       
        frmIndex.Show
        frmIndex.lstIndex.Clear
       
        For i = 1 To MAX_ARROWS
             frmIndex.lstIndex.AddItem i & ": " & Trim$(Arrows(i).Name)
        Next i
       
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    'End If
    
    Case "updatearrow"
    'If (LCase$(Parse$(0)) = "updatearrow") Then
        n = Val(Parse$(1))
       
        Arrows(n).Name = Parse$(2)
        Arrows(n).Pic = Val(Parse$(3))
        Arrows(n).Range = Val(Parse$(4))
        Arrows(n).HasAmmo = Val(Parse$(5))
        Arrows(n).Ammunition = Val(Parse$(6))
        Exit Sub
    'End If
    
    Case "editarrow"
    'If (LCase$(Parse$(0)) = "editarrow") Then
        n = Val(Parse$(1))

        Arrows(n).Name = Parse$(2)
       
        Call ArrowEditorInit
        Exit Sub
    'End If

    Case "checkarrows"
    'If (LCase$(Parse$(0)) = "checkarrows") Then
        n = Val(Parse$(1))
        z = Val(Parse$(2))
        i = Val(Parse$(3))
       
        For x = 1 To MAX_PLAYER_ARROWS
             If Player(n).Arrow(x).Arrow = 0 Then
                 Player(n).Arrow(x).Arrow = 1
                 Player(n).Arrow(x).ArrowNum = z
                 Player(n).Arrow(x).ArrowAnim = Arrows(z).Pic
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
    'End If
    
    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    Case "checkformap"
    'If (LCase$(Parse$(0)) = "checkformap") Then
        ' Erase all players except self
        For i = 1 To MAX_PLAYERS
            If i <> MyIndex Then
                Call SetPlayerMap(i, 0)
            End If
        Next i
        
        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        x = Val(Parse$(1))
        
        ' Get revision
        y = Val(Parse$(2))
    
        If FileExist("maps\map" & x & ".dat") Then
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
    'End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    Case "mapdata"
    'If LCase$(Parse$(0)) = "mapdata" Then
        n = 1
        
        SaveMap.Name = Parse$(n + 1)
        SaveMap.Revision = Val(Parse$(n + 2))
        SaveMap.Moral = Val(Parse$(n + 3))
        SaveMap.Up = Val(Parse$(n + 4))
        SaveMap.Down = Val(Parse$(n + 5))
        SaveMap.Left = Val(Parse$(n + 6))
        SaveMap.Right = Val(Parse$(n + 7))
        SaveMap.Music = Parse$(n + 8)
        SaveMap.BootMap = Val(Parse$(n + 9))
        SaveMap.BootX = Val(Parse$(n + 10))
        SaveMap.BootY = Val(Parse$(n + 11))
        
        n = n + 12
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                SaveMap.Tile(x, y).Ground = Val(Parse$(n))
                SaveMap.Tile(x, y).Mask = Val(Parse$(n + 1))
                SaveMap.Tile(x, y).Anim = Val(Parse$(n + 2))
                SaveMap.Tile(x, y).Mask2 = Val(Parse$(n + 3))
                SaveMap.Tile(x, y).M2Anim = Val(Parse$(n + 4))
                SaveMap.Tile(x, y).Fringe = Val(Parse$(n + 5))
                SaveMap.Tile(x, y).FAnim = Val(Parse$(n + 6))
                SaveMap.Tile(x, y).Fringe2 = Val(Parse$(n + 7))
                SaveMap.Tile(x, y).F2Anim = Val(Parse$(n + 8))
                SaveMap.Tile(x, y).Type = Val(Parse$(n + 9))
                SaveMap.Tile(x, y).Data1 = Val(Parse$(n + 10))
                SaveMap.Tile(x, y).Data2 = Val(Parse$(n + 11))
                SaveMap.Tile(x, y).Data3 = Val(Parse$(n + 12))
                SaveMap.Tile(x, y).String1 = Parse$(n + 13)
                SaveMap.Tile(x, y).String2 = Parse$(n + 14)
                SaveMap.Tile(x, y).String3 = Parse$(n + 15)
                
                n = n + 16
            Next x
        Next y
        
        For x = 1 To MAX_MAP_NPCS
            SaveMap.Npc(x) = Val(Parse$(n))
            n = n + 1
        Next x
                
        ' Save the map
        Call SaveLocalMap(Val(Parse$(1)))
        
        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmEditMap.Visible = False
            frmEndieko.Show
            'frmEndieko.picMapEditor.Visible = False
            
            If frmMapWarp.Visible Then
                Unload frmMapWarp
            End If
            
            If frmMapProperties.Visible Then
                Unload frmMapProperties
            End If
        End If
        
        Exit Sub
    'End If
        
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    Case "mapitemdata"
    'If LCase$(Parse$(0)) = "mapitemdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).Num = Val(Parse$(n))
            SaveMapItem(i).Value = Val(Parse$(n + 1))
            SaveMapItem(i).Dur = Val(Parse$(n + 2))
            SaveMapItem(i).x = Val(Parse$(n + 3))
            SaveMapItem(i).y = Val(Parse$(n + 4))
            
            n = n + 5
        Next i
        
        Exit Sub
    'End If
        
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    Case "mapnpcdata"
    'If LCase$(Parse$(0)) = "mapnpcdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_NPCS
            SaveMapNpc(i).Num = Val(Parse$(n))
            SaveMapNpc(i).x = Val(Parse$(n + 1))
            SaveMapNpc(i).y = Val(Parse$(n + 2))
            SaveMapNpc(i).Dir = Val(Parse$(n + 3))
            
            n = n + 4
        Next i
        
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    Case "mapdone"
    'If LCase$(Parse$(0)) = "mapdone" Then
        Map = SaveMap
        
        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i
        
        GettingMap = False
        
        ' Play music
        If Trim$(Map.Music) <> "None" Then
            Call PlayMidi(Trim$(Map.Music))
        Else
            Call StopMidi
        End If
        
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    Case "saymsg"
    'If (LCase$(Parse$(0)) = "saymsg") Or (LCase$(Parse$(0)) = "broadcastmsg") Or (LCase$(Parse$(0)) = "globalmsg") Or (LCase$(Parse$(0)) = "playermsg") Or (LCase$(Parse$(0)) = "mapmsg") Or (LCase$(Parse$(0)) = "adminmsg") Then
        Call AddText(Parse$(1), Val(Parse$(2)))
        Exit Sub
    'End If
    
    Case "broadcastmsg"
        Call AddText(Parse$(1), Val(Parse$(2)))
        Exit Sub
    Case "globalmsg"
        Call AddText(Parse$(1), Val(Parse$(2)))
        Exit Sub
    Case "playermsg"
        Call AddText(Parse$(1), Val(Parse$(2)))
        Exit Sub
    Case "mapmsg"
        Call AddText(Parse$(1), Val(Parse$(2)))
        Exit Sub
    Case "adminmsg"
        Call AddText(Parse$(1), Val(Parse$(2)))
        Exit Sub
        
    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    Case "spawnitem"
    'If LCase$(Parse$(0)) = "spawnitem" Then
        n = Val(Parse$(1))
        
        MapItem(n).Num = Val(Parse$(2))
        MapItem(n).Value = Val(Parse$(3))
        MapItem(n).Dur = Val(Parse$(4))
        MapItem(n).x = Val(Parse$(5))
        MapItem(n).y = Val(Parse$(6))
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    Case "itemeditor"
    'If (LCase$(Parse$(0)) = "itemeditor") Then
        InItemsEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_ITEMS
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    Case "updateitem"
    'If (LCase$(Parse$(0)) = "updateitem") Then
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
        Item(n).ClassReq = Val(Parse$(11))
        Item(n).AccessReq = Val(Parse$(12))
        
        Item(n).AddHP = Val(Parse$(13))
        Item(n).AddMP = Val(Parse$(14))
        Item(n).AddSP = Val(Parse$(15))
        Item(n).AddStr = Val(Parse$(16))
        Item(n).AddDef = Val(Parse$(17))
        Item(n).AddMagi = Val(Parse$(18))
        Item(n).AddSpeed = Val(Parse$(19))
        Item(n).AddEXP = Val(Parse$(20))
        
        Item(n).desc = Parse$(21)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::
    ' :: Update effect packet ::
    ' ::::::::::::::::::::::::::
    Case "updateeffect"
        n = Val(Parse$(1))
        
        ' Update the Effect
        Effect(n).Name = Parse$(2)
        Effect(n).Effect = Val(Parse$(3))
        Effect(n).Time = Val(Parse$(4))
        Effect(n).Data1 = Val(Parse$(5))
        Effect(n).Data2 = Val(Parse$(6))
        Effect(n).Data3 = Val(Parse$(7))
        
        Exit Sub
       
    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    Case "edititem"
    'If (LCase$(Parse$(0)) = "edititem") Then
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
        Item(n).MagiReq = Val(Parse$(11))
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
        
        Item(n).CannotBeRepaired = Val(Parse$(23))
        Item(n).DropOnDeath = Val(Parse$(24))
        
        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    Case "spawnnpc"
    'If LCase$(Parse$(0)) = "spawnnpc" Then
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
    'End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    Case "npcdead"
    'If LCase$(Parse$(0)) = "npcdead" Then
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
    'End If
    
    ' ::::::::::::::::::::::::::
    ' :: Script editor packet ::
    ' ::::::::::::::::::::::::::
    Case "scripteditor"
    'If (LCase$(Parse$(0)) = "scripteditor") Then
        InScriptEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        frmIndex.lstIndex.AddItem "1: OnJoinGame"
        frmIndex.lstIndex.AddItem "2: OnLeftGame"
        frmIndex.lstIndex.AddItem "3: OnPlayerLevelUp"
        frmIndex.lstIndex.AddItem "4: OnUsingStatPoints"
        frmIndex.lstIndex.AddItem "5: OnScriptedTile"
        frmIndex.lstIndex.AddItem "6: OnPlayerPrompt"
        frmIndex.lstIndex.AddItem "7: OnUseCommand"
        
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::
    ' :: Effect editor packet ::
    ' ::::::::::::::::::::::::::
    Case "effecteditor"
    'If (LCase$(Parse$(0)) = "Effecteditor") Then
        InEffectEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_EFFECTS
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Effect(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    Case "npceditor"
    'If (LCase$(Parse$(0)) = "npceditor") Then
        InNpcEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_NPCS
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    Case "updatenpc"
    'If (LCase$(Parse$(0)) = "updatenpc") Then
        n = Val(Parse$(1))
        
        ' Update the item
        Npc(n).Name = Parse$(2)
        Npc(n).AttackSay = ""
        Npc(n).Sprite = Val(Parse$(3))
        Npc(n).SpawnSecs = 0
        Npc(n).Behavior = 0
        Npc(n).Range = 0
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = 0
            Npc(n).ItemNPC(i).ItemNum = 0
            Npc(n).ItemNPC(i).ItemValue = 0
        Next i
        Npc(n).STR = 0
        Npc(n).DEF = 0
        Npc(n).Speed = 0
        Npc(n).MAGI = 0
        Npc(n).Big = Val(Parse$(4))
        Npc(n).MaxHp = Val(Parse$(5))
        Npc(n).EXP = 0
        Exit Sub
    'End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet :: <- Used for item editor admins only
    ' :::::::::::::::::::::
    Case "editnpc"
    'If (LCase$(Parse$(0)) = "editnpc") Then
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
        Npc(n).MaxHp = Val(Parse$(13))
        Npc(n).EXP = Val(Parse$(14))
        Npc(n).Alignment = Val(Parse$(15))
        
        z = 16
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = Val(Parse$(z))
            Npc(n).ItemNPC(i).ItemNum = Val(Parse$(z + 1))
            Npc(n).ItemNPC(i).ItemValue = Val(Parse$(z + 2))
            z = z + 3
        Next i
        
        ' Initialize the npc editor
        Call NpcEditorInit

        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Edit script packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::::
    Case "editscript"
    'If (LCase$(Parse$(0)) = "editscript") Then
        n = Val(Parse$(2))
        'Script(n).Text = frmEditScript.txtMain.Text
        'Script(n).ScriptNum = n
        
        ' Initialize the script editor
        'frmEditScript.Show vbModal
        Exit Sub
    'End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    Case "mapkey"
    'If (LCase$(Parse$(0)) = "mapkey") Then
        x = Val(Parse$(1))
        y = Val(Parse$(2))
        n = Val(Parse$(3))
                
        TempTile(x, y).DoorOpen = n
        
        Exit Sub
    'End If

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
    Case "editmap"
    'If (LCase$(Parse$(0)) = "editmap") Then
        Call EditorInit
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    Case "shopeditor"
    'If (LCase$(Parse$(0)) = "shopeditor") Then
        InShopEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SHOPS
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    Case "updateshop"
    'If (LCase$(Parse$(0)) = "updateshop") Then
        n = Val(Parse$(1))
        
        ' Update the shop name
        Shop(n).Name = Parse$(2)
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    Case "editshop"
    'If (LCase$(Parse$(0)) = "editshop") Then
        ShopNum = Val(Parse$(1))
        
        ' Update the shop
        Shop(ShopNum).Name = Parse$(2)
        Shop(ShopNum).JoinSay = Parse$(3)
        Shop(ShopNum).LeaveSay = Parse$(4)
        Shop(ShopNum).FixesItems = Val(Parse$(5))
        
        n = 6
        For i = 1 To MAX_TRADES
            
            GiveItem = Val(Parse$(n))
            GiveValue = Val(Parse$(n + 1))
            GetItem = Val(Parse$(n + 2))
            GetValue = Val(Parse$(n + 3))
            
            Shop(ShopNum).TradeItem(i).GiveItem = GiveItem
            Shop(ShopNum).TradeItem(i).GiveValue = GiveValue
            Shop(ShopNum).TradeItem(i).GetItem = GetItem
            Shop(ShopNum).TradeItem(i).GetValue = GetValue
            
            n = n + 4
        Next i
        
        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    'End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    Case "spelleditor"
    'If (LCase$(Parse$(0)) = "spelleditor") Then
        InSpellEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SPELLS
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' :::::::::::::::::::::::::
    Case "updatespell"
    'If (LCase$(Parse$(0)) = "updatespell") Then
        n = Val(Parse$(1))
        
        ' Update the spell name
        Spell(n).Name = Parse$(2)
        Spell(n).Pic = Val(Parse$(3))
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    Case "editspell"
    'If (LCase$(Parse$(0)) = "editspell") Then
        n = Val(Parse$(1))
        
        ' Update the spell
        Spell(n).Name = Parse$(2)
        Spell(n).Pic = Val(Parse$(3))
        Spell(n).ClassReq = Val(Parse$(4))
        Spell(n).LevelReq = Val(Parse$(5))
        Spell(n).Type = Val(Parse$(6))
        Spell(n).Data1 = Val(Parse$(7))
        Spell(n).Data2 = Val(Parse$(8))
        Spell(n).Data3 = Val(Parse$(9))
        Spell(n).MPCost = Val(Parse$(10))
        Spell(n).Sound = Val(Parse$(11))
                        
        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Edit Effect packet :: <- Used for Effect editor admins only
    ' ::::::::::::::::::::::::
    Case "editeffect"
    'If (LCase$(Parse$(0)) = "editEffect") Then
        n = Val(Parse$(1))
        
        ' Update the spell
        Effect(n).Name = Parse$(2)
        Effect(n).Effect = Val(Parse$(3))
        Effect(n).Time = Val(Parse$(4))
        Effect(n).Data1 = Val(Parse$(5))
        Effect(n).Data2 = Val(Parse$(6))
        Effect(n).Data3 = Val(Parse$(7))
                        
        ' Initialize the spell editor
        Call EffectEditorInit

        Exit Sub
    'End If
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    Case "trade"
    'If (LCase$(Parse$(0)) = "trade") Then
        ShopNum = Val(Parse$(1))
        If Val(Parse$(2)) = 1 Then
            frmTrade.picFixItems.Visible = True
        Else
            frmTrade.picFixItems.Visible = False
        End If
        
        n = 3
        For i = 1 To MAX_TRADES
            GiveItem = Val(Parse$(n))
            GiveValue = Val(Parse$(n + 1))
            GetItem = Val(Parse$(n + 2))
            GetValue = Val(Parse$(n + 3))
            
            Trade(1).ItemS(i).ItemGetNum = GetItem
            Trade(1).ItemS(i).ItemGiveNum = GiveItem
            Trade(1).ItemS(i).ItemGetVal = GetValue
            Trade(1).ItemS(i).ItemGiveVal = GiveValue
            
            n = n + 4
        Next i
        
        Dim xx As Long
        For xx = 1 To 6
            Trade(xx).Selected = NO
        Next xx
        
        Trade(1).Selected = YES
        Trade(1).SelectedItem = 1
        
        Call ItemSelected(1, 1)
            
        frmTrade.Show vbModeless, frmEndieko
        Exit Sub
    'End If


    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    Case "spells"
    'If (LCase$(Parse$(0)) = "spells") Then
        
        frmEndieko.picPlayerSpells.Visible = True
        frmEndieko.lstSpells.Clear
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(i) = Val(Parse$(i))
            If Player(MyIndex).Spell(i) <> 0 Then
                frmEndieko.lstSpells.AddItem i & ": " & Trim$(Spell(Player(MyIndex).Spell(i)).Name)
            Else
                frmEndieko.lstSpells.AddItem "<free spells slot>"
            End If
        Next i
        
        frmEndieko.lstSpells.ListIndex = 0
        Call UpdateVisSpell
    'End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    Case "weather"
    'If (LCase$(Parse$(0)) = "weather") Then
        GameWeather = Val(Parse$(1))
    'End If

    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    Case "time"
    'If (LCase$(Parse$(0)) = "time") Then
        GameTime = Val(Parse$(1))
    'End If
    
    ' :::::::::::::::::::::
    ' :: Get Online List ::
    ' :::::::::::::::::::::
    Case "onlinelist"
    'If LCase$(Parse$(0)) = "onlinelist" Then
        frmEndieko.lstOnline.Clear
    
        n = 2
        z = Val(Parse$(1))
        For x = n To (z + 1)
            frmEndieko.lstOnline.AddItem Trim$(Parse$(n))
            n = n + 2
        Next x
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
    Case "blitplayerdmg"
    'If LCase$(Parse$(0)) = "blitplayerdmg" Then
        DmgDamage = Val(Parse$(1))
        NPCWho = Val(Parse$(2))
        DmgTime = GetTickCount
        iii = 0
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
    If LCase$(Parse$(0)) = "blitnpcdmg" Then
        NPCDmgDamage = Val(Parse$(1))
        NPCDmgTime = GetTickCount
        ii = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Blit Overhead Damage ::
    ' ::::::::::::::::::::::::::
    Case "bltoverhead"
    'If LCase$(Parse$(0)) = "bltoverhead" Then
        Overhead.Color = Val(Parse$(1))
        Overhead.Msg = Parse$(2)
        Overhead.Time = GetTickCount
        Overhead.ii = 0
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::::::::::
    ' :: Retrieve the player's inventory ::
    ' :::::::::::::::::::::::::::::::::::::
    Case "pptrading"
    'If LCase$(Parse$(0)) = "pptrading" Then
        frmPlayerTrade.Items1.Clear
        frmPlayerTrade.Items2.Clear
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = ""
            Trading2(i).InvNum = 0
            Trading2(i).InvName = ""
            frmPlayerTrade.Items1.AddItem i & ": <Nothing>"
            frmPlayerTrade.Items2.AddItem i & ": <Nothing>"
        Next i
        
        frmPlayerTrade.Items1.ListIndex = 0
        
        Call UpdateTradeInventory
        frmPlayerTrade.Show vbModeless, frmEndieko
        Exit Sub
    'End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    Case "qtrade"
    'If LCase$(Parse$(0)) = "qtrade" Then
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = ""
            Trading2(i).InvNum = 0
            Trading2(i).InvName = ""
        Next i
        
        frmPlayerTrade.Command1.ForeColor = &HFF00&
        frmPlayerTrade.Command2.ForeColor = &HFF00&
        
        frmPlayerTrade.Visible = False
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    Case "updatetradeitem"
    'If LCase$(Parse$(0)) = "updatetradeitem" Then
            n = Val(Parse$(1))
            
            Trading2(n).InvNum = Val(Parse$(2))
            Trading2(n).InvName = Parse$(3)
            
            If STR$(Trading2(n).InvNum) <= 0 Then
                frmPlayerTrade.Items2.List(n - 1) = n & ": <Nothing>"
            Else
                frmPlayerTrade.Items2.List(n - 1) = n & ": " & Trim$(Trading2(n).InvName)
            End If
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    Case "trading"
    'If LCase$(Parse$(0)) = "trading" Then
        n = Val(Parse$(1))
            If STR$(n) = 0 Then frmPlayerTrade.Command2.ForeColor = &H80000012
            If STR$(n) = 1 Then frmPlayerTrade.Command2.ForeColor = &HFF00&
        Exit Sub
    'End If

    ' :::::::::::::::::::::::::
    ' :: Chat System Packets ::
    ' :::::::::::::::::::::::::
    Case "ppchatting"
    'If LCase$(Parse$(0)) = "ppchatting" Then
        frmPlayerChat.txtChat.Text = ""
        frmPlayerChat.txtSay.Text = ""
        frmPlayerChat.Label1.Caption = "Chatting With: " & Trim$(Player(Val(Parse$(1))).Name)

        frmPlayerChat.Show vbModeless, frmEndieko
        Exit Sub
    'End If
    
    Case "qchat"
    'If LCase$(Parse$(0)) = "qchat" Then
        frmPlayerChat.txtChat.Text = ""
        frmPlayerChat.txtSay.Text = ""
        frmPlayerChat.Visible = False
        frmPlayerTrade.Command2.ForeColor = &H8000000F
        frmPlayerTrade.Command1.ForeColor = &H8000000F
        Exit Sub
    'End If
    
    Case "sendchat"
    'If LCase$(Parse$(0)) = "sendchat" Then
        Dim s As String
  
        s = vbNewLine & GetPlayerName(Val(Parse$(2))) & "> " & Trim$(Parse$(1))
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text)
        frmPlayerChat.txtChat.SelColor = QBColor(Brown)
        frmPlayerChat.txtChat.SelText = s
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text) - 1
        Exit Sub
    'End If
' :::::::::::::::::::::::::::::
' :: END Chat System Packets ::
' :::::::::::::::::::::::::::::

    ' :::::::::::::::::::::::
    ' :: Play Sound Packet ::
    ' :::::::::::::::::::::::
    Case "sound"
    'If LCase$(Parse$(0)) = "sound" Then
        s = LCase$(Parse$(1))
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
                Call PlaySound("magic" & Val(Parse$(2)) & ".wav")
            Case "warp"
                Call PlaySound("warp.wav")
            Case "pain"
                Call PlaySound("pain.wav")
            Case "soundattribute"
                Call PlaySound(Trim$(Parse$(2)))
        End Select
        Exit Sub
    'End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Sprite Change Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    Case "spritechange"
    'If LCase$(Parse$(0)) = "spritechange" Then
        i = MsgBox("Would you like to buy this sprite?", 4, "Buying Sprite")
        If i = 6 Then
            Call SendData("buysprite" & SEP_CHAR & END_CHAR)
        End If
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Change Player Direction Packet ::
    ' ::::::::::::::::::::::::::::::::::::
    Case "changedir"
    'If LCase$(Parse$(0)) = "changedir" Then
        Player(Val(Parse$(2))).Dir = Val(Parse$(1))
        Exit Sub
    'End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Flash Movie Event Packet ::
    ' ::::::::::::::::::::::::::::::
    Case "flashevent"
    'If LCase$(Parse$(0)) = "flashevent" Then
        If LCase$(Mid$(Trim$(Parse$(1)), 1, 7)) = "http://" Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopMidi
            Call StopSound
            frmFlash.Flash.LoadMovie 0, Trim$(Parse$(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmEndieko
        ElseIf FileExist("Animations\" & Trim$(Parse$(1))) = True Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopMidi
            Call StopSound
            frmFlash.Flash.LoadMovie 0, App.Path & "\Animations\" & Trim$(Parse$(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmEndieko
        End If
        Exit Sub
    'End If
    
    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    Case "prompt"
    'If LCase$(Parse$(0)) = "prompt" Then
        i = MsgBox(Trim$(Parse$(1)), vbYesNo)
        Call SendData("prompt" & SEP_CHAR & i & SEP_CHAR & Val(Parse$(2)) & SEP_CHAR & END_CHAR)
        Exit Sub
    'End If

    ' ::::::::::::::::::::::::::::
    ' :: Emoticon editor packet ::
    ' ::::::::::::::::::::::::::::
    Case "emoticoneditor"
    'If (LCase$(Parse$(0)) = "emoticoneditor") Then
        InEmoticonEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        For i = 0 To MAX_EMOTICONS
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Emoticons(i).Command)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    'End If
    
    Case "updateemoticon"
    'If (LCase$(Parse$(0)) = "updateemoticon") Then
        n = Val(Parse$(1))
        
        Emoticons(n).Command = Parse$(2)
        Emoticons(n).Pic = Val(Parse$(3))
        Exit Sub
    'End If
    
    Case "editemoticon"
    'If (LCase$(Parse$(0)) = "editemoticon") Then
        n = Val(Parse$(1))

        Emoticons(n).Command = Parse$(2)
        Emoticons(n).Pic = Val(Parse$(3))
        
        Call EmoticonEditorInit
        Exit Sub
    'End If
    
    Case "updateemoticon"
    'If (LCase$(Parse$(0)) = "updateemoticon") Then
        n = Val(Parse$(1))
        
        Emoticons(n).Command = Parse$(2)
        Emoticons(n).Pic = Val(Parse$(3))
        Exit Sub
    'End If
    
    Case "checkemoticons"
    'If (LCase$(Parse$(0)) = "checkemoticons") Then
        n = Val(Parse$(1))
        
        Player(n).Emoticon = Val(Parse$(2))
        Player(n).EmoticonT = GetTickCount
        Exit Sub
    'End If
    
    Case "checksprite"
    'If (LCase$(Parse$(0)) = "checksprite") Then
        n = Val(Parse$(1))
        
        Player(n).Sprite = Val(Parse$(2))
        Exit Sub
    'End If
    
    Case "mapreport"
    'If (LCase$(Parse$(0)) = "mapreport") Then
        n = 1
        
        frmMapReport.lstIndex.Clear
        For i = 1 To MAX_MAPS
            frmMapReport.lstIndex.AddItem i & ": " & Trim$(Parse$(n))
            n = n + 1
        Next i
        
        frmMapReport.Show vbModeless, frmEndieko
        Exit Sub
    'End If
    
    Case Else
        MsgBox (LCase$(Parse$(0)) & " was not correctly handled.")
End Select
End Sub

Function ConnectToServer() As Boolean
Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmEndieko.Socket.Close
    frmEndieko.Socket.Connect
    
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
    If frmEndieko.Socket.State = sckConnected Then
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

'Sub SendData(ByVal Data As String)
'    If IsConnected Then
'        frmEndieko.Socket.SendData Data
'        DoEvents
'    End If
'End Sub

Sub SendData(ByVal Data As String)
Dim lR As Long
    If IsConnected Then
        Data = Compress(Data, lR)
             Data = lR & SEP_CHAR & Data
             frmEndieko.Socket.SendData Data
        DoEvents
    End If
End Sub

Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim Packet As String
    
    Call SendData("HDSerial" & SEP_CHAR & GetHDSerial("C") & SEP_CHAR & END_CHAR)

    Packet = "newaccount" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
Dim Packet As String
    
    Call SendData("HDSerial" & SEP_CHAR & GetHDSerial("C") & SEP_CHAR & END_CHAR)
    
    Packet = "delaccount" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim Packet As String

    Call SendData("HDSerial" & SEP_CHAR & GetHDSerial("C") & SEP_CHAR & END_CHAR)

    Packet = "login" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & SEC_CODE1 & SEP_CHAR & SEC_CODE2 & SEP_CHAR & SEC_CODE3 & SEP_CHAR & SEC_CODE4 & SEP_CHAR & END_CHAR
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

Sub SendUnBan(ByVal Name As String)
Dim Packet As String

    Packet = "UNBANPLAYER" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
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

Sub SendPlayerMove()
Dim Packet As String

    Packet = "playermove" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & SEP_CHAR & Player(MyIndex).SP & SEP_CHAR & END_CHAR
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

Sub SendMap()
Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long

    Packet = "MAPDATA" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim$(Map.Name) & SEP_CHAR & Map.Revision & SEP_CHAR & Map.Moral & SEP_CHAR & Map.Up & SEP_CHAR & Map.Down & SEP_CHAR & Map.Left & SEP_CHAR & Map.Right & SEP_CHAR & Map.Music & SEP_CHAR & Map.BootMap & SEP_CHAR & Map.BootX & SEP_CHAR & Map.BootY & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map.Tile(x, y)
                Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR
            End With
        Next x
    Next y
    
    For x = 1 To MAX_MAP_NPCS
        Packet = Packet & Map.Npc(x) & SEP_CHAR
    Next x
    
    Packet = Packet & END_CHAR
    
    x = Int(Len(Packet) / 2)
    P1 = Mid$(Packet, 1, x)
    P2 = Mid$(Packet, x + 1, Len(Packet) - x)
    Call SendData(Packet)
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

Sub SendSaveItem(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "SAVEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagiReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).desc & SEP_CHAR & Item(ItemNum).CannotBeRepaired & SEP_CHAR & Item(ItemNum).DropOnDeath
    Packet = Packet & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
                
Sub SendRequestEditEmoticon()
Dim Packet As String

    Packet = "REQUESTEDITEMOTICON" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveEmoticon(ByVal EmoNum As Long)
Dim Packet As String

    Packet = "SAVEEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditScript()
Dim Packet As String

    Packet = "REQUESTEDITSCRIPT" & SEP_CHAR & END_CHAR
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
    
    Packet = "SAVENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).Speed & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).EXP & SEP_CHAR & Npc(NpcNum).Alignment & SEP_CHAR
    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(NpcNum).ItemNPC(i).Chance
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
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
Sub SendOnlineList()
Dim Packet As String

Packet = "ONLINELIST" & SEP_CHAR & END_CHAR
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

Sub SendSaveShop(ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "SAVESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditEffect()
Dim Packet As String

    Packet = "REQUESTEDITEFFECT" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveEffect(ByVal EffectNum As Long)
Dim Packet As String

    Packet = "SAVEEFFECT" & SEP_CHAR & EffectNum & SEP_CHAR & Trim$(Effect(EffectNum).Name) & SEP_CHAR & Effect(EffectNum).Effect & SEP_CHAR & Effect(EffectNum).Time & SEP_CHAR & Effect(EffectNum).Data1 & SEP_CHAR & Effect(EffectNum).Data2 & SEP_CHAR & Effect(EffectNum).Data3 & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditSpell()
Dim Packet As String

    Packet = "REQUESTEDITSPELL" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveSpell(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "SAVESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).Pic & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditMap()
Dim Packet As String

    Packet = "REQUESTEDITMAP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendTradeRequest(ByVal Name As String)
Dim Packet As String

    Packet = "PPTRADE" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendAcceptTrade()
Dim Packet As String

    Packet = "ATRADE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDeclineTrade()
Dim Packet As String

    Packet = "DTRADE" & SEP_CHAR & END_CHAR
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

Sub SendSetPlayerSprite(ByVal Name As String, ByVal SpriteNum As Byte)
Dim Packet As String

    Packet = "SETPLAYERSPRITE" & SEP_CHAR & Name & SEP_CHAR & SpriteNum & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditArrow()
Dim Packet As String

    Packet = "REQUESTEDITARROW" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveArrow(ByVal ArrowNum As Long)
Dim Packet As String

    Packet = "SAVEARROW" & SEP_CHAR & ArrowNum & SEP_CHAR & Trim$(Arrows(ArrowNum).Name) & SEP_CHAR & Arrows(ArrowNum).Pic & SEP_CHAR & Arrows(ArrowNum).Range & SEP_CHAR & Arrows(ArrowNum).HasAmmo & SEP_CHAR & Arrows(ArrowNum).Ammunition & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendGameTime()
Dim Packet As String

Packet = "GmTime" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub JailPlayer(ByVal Name As String)
Dim Packet As String

Packet = "JAILPLAYER" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub MutePlayer(ByVal Name As String)
Dim Packet As String

Packet = "MUTEPLAYER" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub UnMutePlayer(ByVal Name As String)
Dim Packet As String

Packet = "UNMUTEPLAYER" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub SetInvisiblity()
Dim Packet As String

Packet = "INVISIBLE" & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub PlayerMap(ByVal Name As String)
Dim Packet As String

Packet = "DISABLEPLAYERMAP" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub PlayerBroadcast(ByVal Name As String)
Dim Packet As String

Packet = "DISABLEPLAYERBROADCAST" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub PlayerEmote(ByVal Name As String)
Dim Packet As String

Packet = "DISABLEPLAYEREMOTE" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub PlayerPrivate(ByVal Name As String)
Dim Packet As String

Packet = "DISABLEPLAYERPRIVATE" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub PlayerGlobal(ByVal Name As String)
Dim Packet As String

Packet = "DISABLEPLAYERGLOBAL" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub PlayerGuild(ByVal Name As String)
Dim Packet As String

Packet = "DISABLEPLAYERGUILD" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub PlayerParty(ByVal Name As String)
Dim Packet As String

Packet = "DISABLEPLAYERPARTY" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub PlayerAdmin(ByVal Name As String)
Dim Packet As String

Packet = "DISABLEPLAYERADMIN" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Public Sub CastSpell()
If Player(MyIndex).Spell(frmEndieko.lstSpells.ListIndex + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & frmEndieko.lstSpells.ListIndex + 1 & SEP_CHAR & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If
End Sub
