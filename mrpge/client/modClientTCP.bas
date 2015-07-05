Attribute VB_Name = "modClientTCP"
Option Explicit

Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean
Public RecieveMapTest As Boolean

Sub TcpInit()
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    PlayerBuffer = ""
End Sub

Sub TcpInit2(ByVal ip As String, ByVal port As Integer)
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    PlayerBuffer = ""

    'frmMirage.Socket.RemoteHost = "80.192.161.89"
    frmMirage.Socket.RemoteHost = ip
    
    
    
    
    'REMEMBER DON'T KILL THIS LINE IT WILL HAVE YOU SEARCHING FOR HOURS!!!
    frmMirage.Socket.RemotePort = port
End Sub

Sub TcpDestroy()
    frmMirage.Socket.Close
    
    If frmChars.Visible Then frmChars.Visible = False
    If frmDeleteAccount.Visible Then frmDeleteAccount.Visible = False
    If frmLogin.Visible Then frmLogin.Visible = False
    If frmNewAccount.Visible Then frmNewAccount.Visible = False
    If frmNewChar.Visible Then frmNewChar.Visible = False
End Sub

Sub IncomingData(ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim top As String * 3
Dim start As Integer
Dim i As Long

    frmMirage.Socket.GetData Buffer, vbString, DataLength
    'Dim abCipher() As Byte
        'Dim aKet() As Byte
        'aKet() = StrConv("6$db sYS5&(£'S HseT£ w4uaz5 \gw43 y\4wu", vbFromUnicode)
        'Call blf_KeyInit(aKet)
       ' abCipher = StrConv(Buffer, vbFromUnicode)
        'Dim abPlain() As Byte
        '    abPlain = blf_BytesDec(abCipher)
        '    Buffer = StrConv(abPlain, vbUnicode)
        '    Buffer = Replace(Buffer, Chr(1), "")
'            Buffer = Replace(Buffer, Chr(2), "")
'            Buffer = Replace(Buffer, Chr(3), "")
'            Buffer = Replace(Buffer, Chr(4), "")
'            Buffer = Replace(Buffer, Chr(5), "")
'            Buffer = Replace(Buffer, Chr(6), "")
'            Buffer = Replace(Buffer, Chr(7), "")
'            Buffer = Replace(Buffer, Chr(8), "")

        'Debug.Print "buffer: " & Buffer
        'For i = 1 To Len(Buffer)
        '    Debug.Print Mid(Buffer, i, 1) & " - " & Asc(Mid(Buffer, i, 1))
        'Next i
        
       ' Buffer = Replace(Buffer, Chr(0), "")
    PlayerBuffer = PlayerBuffer & Buffer
   '     For i = 1 To Len(Buffer)
    '     MsgBox Asc(Mid(Buffer, i, 1))
'        Next i
        'MsgBox Asc(Mid(Buffer, 1, 1))
    start = InStr(PlayerBuffer, END_CHAR)
    Do While start > 0
        Packet = Mid(PlayerBuffer, 1, start - 1)
        PlayerBuffer = Mid(PlayerBuffer, start + 1, Len(PlayerBuffer))
        start = InStr(PlayerBuffer, END_CHAR)
        If Len(Packet) > 0 Then
            Call HandleData(Packet)
        End If
    Loop
End Sub

Sub HandleData(ByVal data As String)
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

    ' Handle Data
    Parse = Split(data, SEP_CHAR)
    
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
    
    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playsound" Then
        Name = Parse(1)
        Call PlaySound(Name)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Updater Launch packet ::
    ' :::::::::::::::::::::::::::
    
    If LCase(Parse(0)) = "version" Then
        'MsgBox App.Path & "\update.exe"
        Call Shell(App.Path & "\update.exe", vbNormalFocus)
        GameDestroy
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "version1" Then
        MsgBox "Your version of the updater is incorrect." & vbCrLf & "Please download the latest version."
        x = ShellExecute(frmMirage.hwnd, "Open", "http://afterdarkness.squiggleuk.com", 0&, 0&, 0&)
        GameDestroy
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Quest message packet  ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "questmsg" Then
    Msg = Parse(1)
        'show the quest message screen.
        frmMirage.picQuestMsg.top = 0
        frmMirage.picQuestMsg.Left = 517
        frmMirage.lblQuestMsg.Caption = Msg
        frmMirage.picQuestMsg.Visible = True
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: map report packet     ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapreport" Then
    On Error Resume Next
        InWarp = True
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_MAPS Step 1
            frmIndex.lstIndex.AddItem i & ": " & Parse(i)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        frmIndex.SetFocus
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
            
            If Trim(Name) = "" Then
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
        StopMidi
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
            
            Class(i).str = Val(Parse(n + 4))
            Class(i).intel = Val(Parse(n + 5))
            Class(i).dex = Val(Parse(n + 6))
            Class(i).con = Val(Parse(n + 7))
            Class(i).wiz = Val(Parse(n + 8))
            Class(i).cha = Val(Parse(n + 9))
            'Class(i).DEF = Val(Parse(n + 5))
            'Class(i).SPEED = Val(Parse(n + 6))
            'Class(i).MAGI = Val(Parse(n + 7))
            
            n = n + 10
        Next i
        
        ' Used for if the player is creating a new character
        frmNewChar.Visible = True
        frmSendGetData.Visible = False

        frmNewChar.cmbClass.Clear

        For i = 0 To Max_Classes
            frmNewChar.cmbClass.AddItem Trim(Class(i).Name)
        Next i
            
        frmNewChar.cmbClass.ListIndex = 0
        frmNewChar.lblHP.Caption = str(Class(0).HP)
        frmNewChar.lblMP.Caption = str(Class(0).MP)
        frmNewChar.lblSP.Caption = str(Class(0).SP)
    
        frmNewChar.lblSTR.Caption = str(Class(0).str)
        frmNewChar.lblCha.Caption = str(Class(0).cha)
        frmNewChar.lblINT.Caption = str(Class(0).intel)
        frmNewChar.lblDex.Caption = str(Class(0).dex)
        frmNewChar.lblCon.Caption = str(Class(0).con)
        frmNewChar.lblWiz.Caption = str(Class(0).wiz)
        
        'frmNewChar.lblDEF.Caption = STR(Class(0).DEF)
        'frmNewChar.lblSPEED.Caption = STR(Class(0).SPEED)
        'frmNewChar.lblMAGI.Caption = STR(Class(0).MAGI)
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
            
            Class(i).str = Val(Parse(n + 4))
            Class(i).intel = Val(Parse(n + 5))
            Class(i).dex = Val(Parse(n + 6))
            Class(i).con = Val(Parse(n + 7))
            Class(i).wiz = Val(Parse(n + 8))
            Class(i).cha = Val(Parse(n + 9))
            'Class(i).DEF = Val(Parse(n + 5))
            'Class(i).SPEED = Val(Parse(n + 6))
            'Class(i).MAGI = Val(Parse(n + 7))
            
            n = n + 10
        Next i
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "ingame" Then
        frmMirage.Caption = Parse(1)
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
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player bank packet      ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerbank" Then
        n = 1
        For i = 1 To MAX_BANK
            Call SetPlayerBankItemNum(MyIndex, i, Val(Parse(n)))
            Call SetPlayerBankItemValue(MyIndex, i, Val(Parse(n + 1)))
            Call SetPlayerBankItemDur(MyIndex, i, Val(Parse(n + 2)))
            
            n = n + 3
        Next i
        Call UpdateBank
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
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player bank update packet      ::
    ' ::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerbankupdate" Then
        n = Val(Parse(1))
        
        Call SetPlayerBankItemNum(MyIndex, n, Val(Parse(2)))
        Call SetPlayerBankItemValue(MyIndex, n, Val(Parse(3)))
        Call SetPlayerBankItemDur(MyIndex, n, Val(Parse(4)))
        'DoEvents
        'Call UpdateBank
        'DoEvents
        Call SendData("playerbank" & SEP_CHAR & END_CHAR)
        DoEvents
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
        Player(MyIndex).maxHP = Val(Parse(1))
        Call SetPlayerHP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then
            frmMirage.lblHP.Caption = GetPlayerHP(MyIndex) & "/" & GetPlayerMaxHP(MyIndex) 'Int(GetPlayerHP(MyIndex) / GetPlayerMaxHP(MyIndex) * 100) & "%"
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
            frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & "/" & GetPlayerMaxMP(MyIndex) 'Int(GetPlayerMP(MyIndex) / GetPlayerMaxMP(MyIndex) * 100) & "%"
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playersp" Then
    On Error Resume Next
        Player(MyIndex).MaxSP = Val(Parse(1))
        Call SetPlayerSP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxSP(MyIndex) > 0 Then
            frmMirage.lblSP.Caption = GetPlayerSP(MyIndex) & "/" & GetPlayerMaxSP(MyIndex) 'Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Player pp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playerpp" Then
        Player(MyIndex).MaxPP = Val(Parse(1))
        Call SetPlayerPP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxPP(MyIndex) > 0 Then
            frmMirage.lblPP.Caption = GetPlayerPP(MyIndex) & "/" & GetPlayerMaxPP(MyIndex) 'Int(GetPlayerpP(MyIndex) / GetPlayerMaxpP(MyIndex) * 100) & "%"
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Player stats packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerstats" Then
        Call SetPlayerSTR(MyIndex, Val(Parse(1)))
        Call SetPlayerInt(MyIndex, Val(Parse(2)))
        Call SetPlayerDex(MyIndex, Val(Parse(3)))
        Call SetPlayerCon(MyIndex, Val(Parse(4)))
        Call SetPlayerWiz(MyIndex, Val(Parse(5)))
        Call SetPlayerCha(MyIndex, Val(Parse(6)))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player stats info  packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerstatsinfo" Then
    
        frmMirage.lblStats1 = "str: " & Val(Parse(1))
        frmMirage.lblStats1 = frmMirage.lblStats1 & vbNewLine & "int: " & Val(Parse(2))
        frmMirage.lblStats1 = frmMirage.lblStats1 & vbNewLine & "dex: " & Val(Parse(3))
        
        frmMirage.lblStats2 = "con: " & Val(Parse(4))
        frmMirage.lblStats2 = frmMirage.lblStats2 & vbNewLine & "wiz: " & Val(Parse(5))
        frmMirage.lblStats2 = frmMirage.lblStats2 & vbNewLine & "cha: " & Val(Parse(6))
        
        frmMirage.lblLevel = "level: " & Val(Parse(7))
        
        frmMirage.lblexp = Val(Parse(8)) & "/" & Val(Parse(9))
        
        frmMirage.lblcrit = Val(Parse(10)) & "%"
        frmMirage.lblblock = Val(Parse(11)) & "%"
        
        If Val(Parse(8)) >= Val(Parse(9)) Then
            frmMirage.shpexp.Width = 130
            frmMirage.shpexp.BackColor = &H80FF&
        Else
            frmMirage.shpexp.Width = (Val(Parse(8)) / Val(Parse(9))) * 130
            frmMirage.shpexp.BackColor = &HC0&
        End If
        
        
        frmMirage.shpblock.Width = (Val(Parse(11)) / 100) * 130
        frmMirage.shpcrit.Width = (Val(Parse(10)) / 100) * 130
        
        frmMirage.clearPanes
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
        Call SetPlayerColour(i, Val(Parse(10)))
        Call SetPlayerHP(i, Val(Parse(11)))
        Call SetPlayerMP(i, Val(Parse(12)))
        Call SetPlayerMaxHP(i, Val(Parse(13)))
        Call SetPlayerMaxMP(i, Val(Parse(14)))
        
        
        ' Make sure they aren't walking
        Player(i).moving = 0
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
        x = Val(Parse(2))
        y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        Call SetPlayerX(i, x)
        Call SetPlayerY(i, y)
        Call SetPlayerDir(i, Dir)
                
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).moving = n
        
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
        MapNpc(i).moving = n
        
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
    
    ' :::::::::::::::::::::::::
    ' :: Pet movement packet ::
    ' :::::::::::::::::::::::::
'    If (LCase(Parse(0)) = "petmove") Then
'
'        i = Val(Parse(1))
'        x = Val(Parse(2))
'        y = Val(Parse(3))
'        Dir = Val(Parse(4))
'        n = Val(Parse(5))
'
'        Pets(i).x = x
'        Pets(i).y = y
'        Pets(i).Dir = Dir
'        Pets(i).XOffset = 0
'        Pets(i).YOffset = 0
'        Pets(i).moving = 1
'        Debug.Print "Move n = " & n
'        Select Case Pets(i).Dir
'            Case DIR_UP
'            Debug.Print "Move: Up"
'                Pets(i).YOffset = PIC_Y
'            Case DIR_DOWN
'            Debug.Print "Move: Down"
'                Pets(i).YOffset = PIC_Y * -1
'            Case DIR_LEFT
'            Debug.Print "Move: Left"
'                Pets(i).XOffset = PIC_X
'            Case DIR_RIGHT
'            Debug.Print "Move: Right"
'                Pets(i).XOffset = PIC_X * -1
'        End Select
'        Exit Sub
'    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        Call SetPlayerDir(i, Dir)
        
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).moving = 0
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
        MapNpc(i).moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerxy") Then
        x = Val(Parse(1))
        y = Val(Parse(2))
        
        Call SetPlayerX(MyIndex, x)
        Call SetPlayerY(MyIndex, y)
        
        ' Make sure they aren't walking
        Player(MyIndex).moving = 0
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
        Call SendData("getstats" & SEP_CHAR & END_CHAR)
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
        x = Val(Parse(1))
        
        ' Get revision
        y = Val(Parse(2))
        
        If FileExist("maps\map" & x & ".dat") Then
            ' Check to see if the revisions match
            If GetMapRevision(x) = y Then
                ' We do so we dont need the map
                
                ' Load the map
                Call LoadMap(x)
                
                Call SendData("needmap" & SEP_CHAR & "no" & SEP_CHAR & END_CHAR)
                canMoveNow = True
                DoEvents
                Exit Sub
            End If
        End If
        
        ' Either the revisions didn't match or we dont have the map, so we need it
        'Debug.Print "NEED MAP"
        Call SendData("needmap" & SEP_CHAR & "yes" & SEP_CHAR & END_CHAR)
        DoEvents
        'Call SendData("needmap" & SEP_CHAR & "yes" & SEP_CHAR & END_CHAR)
        'DoEvents
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "mapdata" Then
        n = 1
        'On Error Resume Next
        SaveMap.Name = Parse(n + 1)
        SaveMap.Revision = Val(Parse(n + 2))
        SaveMap.Moral = Val(Parse(n + 3))
        SaveMap.Up = Val(Parse(n + 4))
        SaveMap.Down = Val(Parse(n + 5))
        SaveMap.Left = Val(Parse(n + 6))
        SaveMap.Right = Val(Parse(n + 7))
        SaveMap.music = Val(Parse(n + 8))
        SaveMap.BootMap = Val(Parse(n + 9))
        SaveMap.BootX = Val(Parse(n + 10))
        SaveMap.BootY = Val(Parse(n + 11))
        SaveMap.Shop = Val(Parse(n + 12))
        
        n = n + 13
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                SaveMap.Tile(x, y).Ground = Val(Parse(n))
                SaveMap.Tile(x, y).mask = Val(Parse(n + 1))
                SaveMap.Tile(x, y).Anim = Val(Parse(n + 2))
                SaveMap.Tile(x, y).Fringe = Val(Parse(n + 3))
                SaveMap.Tile(x, y).type = Val(Parse(n + 4))
                SaveMap.Tile(x, y).Data1 = Val(Parse(n + 5))
                SaveMap.Tile(x, y).Data2 = Val(Parse(n + 6))
                SaveMap.Tile(x, y).Data3 = Val(Parse(n + 7))
                SaveMap.Tile(x, y).Data4 = Val(Parse(n + 8))
                SaveMap.Tile(x, y).Data5 = Val(Parse(n + 9))
                SaveMap.Tile(x, y).TileSheet_Ground = Val(Parse(n + 10))
                SaveMap.Tile(x, y).TileSheet_Fringe = Val(Parse(n + 11))
                SaveMap.Tile(x, y).TileSheet_Anim = Val(Parse(n + 12))
                SaveMap.Tile(x, y).TileSheet_Mask = Val(Parse(n + 13))
                DoEvents
                n = n + 14
            Next x
        Next y
        'Exit Sub
    'End If
    
    'If LCase(Parse(0)) = "mapdata2" Then
        'n = 1
        'On Error Resume Next
      ' For y = 0 To MAX_MAPY
       '     For x = 0 To MAX_MAPX
       '         SaveMap.Tile(x, y).TileSheet_Ground = Val(Parse(n))
       '         SaveMap.Tile(x, y).TileSheet_Fringe = Val(Parse(n + 1))
       '         SaveMap.Tile(x, y).TileSheet_Anim = Val(Parse(n + 2))
       '         SaveMap.Tile(x, y).TileSheet_Mask = Val(Parse(n + 3))
       '         n = n + 4
       '     Next x
       ' Next y
       
       
       
        For x = 1 To MAX_MAP_NPCS
            SaveMap.Npc(x) = Val(Parse(n))
            n = n + 1
        Next x
           SaveMap.Respawn = Parse(n)
           SaveMap.Night = Val(Parse(n + 1))

           SaveMap.Bank = CBool(Parse(n + 2))
           SaveMap.street = Trim(Parse(n + 3))
        ' Save the map
        Call SaveLocalMap(Val(Parse(n + 4)))
        
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
        'If RecieveMapTest = False Then
        '    Call SendData("needmap" & SEP_CHAR & "yes" & SEP_CHAR & END_CHAR)
       '     RecieveMapTest = True
       ' Else
       '     RecieveMapTest = False
       ' End If
        canMoveNow = True
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapitemdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).num = Val(Parse(n))
            SaveMapItem(i).value = Val(Parse(n + 1))
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
            SaveMapNpc(i).num = Val(Parse(n))
            SaveMapNpc(i).x = Val(Parse(n + 1))
            SaveMapNpc(i).y = Val(Parse(n + 2))
            SaveMapNpc(i).Dir = Val(Parse(n + 3))
            SaveMapNpc(i).maxHP = Val(Parse(n + 4))
            SaveMapNpc(i).HP = Val(Parse(n + 5))
            MapNpc(i).maxHP = Val(Parse(n + 4))
            MapNpc(i).HP = Val(Parse(n + 5))
            n = n + 6
        Next i
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapdone" Then
        map = SaveMap
        
        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i
        
        GettingMap = False
        
        ' Play music
        'Call StopMidi
        If map.music > 0 Then
            Call PlayMidi("music" & Trim(str(map.music)))
            DoEvents
        End If
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Day or Night  packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "night" Then
        If Parse(1) = 1 Then
            blnNight = True
        Else
            blnNight = False
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "playermsg") Then
        Call AddTextNew(frmMirage.txtChannelPrivate, Parse(1), Val(Parse(2)))
        Call AddTextNew(frmMirage.txtChannelAll, Parse(1), Val(Parse(2)))
    End If
    
    If (LCase(Parse(0)) = "guildmsg") Then
        Call AddTextNew(frmMirage.txtChannelGuild, Parse(1), Val(Parse(2)))
        Call AddTextNew(frmMirage.txtChannelAll, Parse(1), Val(Parse(2)))
    End If
    
    If (LCase(Parse(0)) = "saymsg") Or (LCase(Parse(0)) = "mapmsg") Then
        Call AddTextNew(frmMirage.txtChannelAll, Parse(1), Val(Parse(2)))
        'Call AddText(Parse(1), Val(Parse(2)))
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "massmsg") Then
        Call AddTextNew(frmMirage.txtChannelAll, Parse(1), RGB_MassMsg)
        'Call AddText(Parse(1), Val(Parse(2)))
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "servermsg") Then
        Call addServerText(Parse(1))
    End If
    
    If (LCase(Parse(0)) = "broadcastmsg") Or (LCase(Parse(0)) = "globalmsg") Or (LCase(Parse(0)) = "adminmsg") Then
        Call AddGlobalText(Parse(1), Val(Parse(2)))
        Exit Sub
    End If
    

    
    
    ' ::::::::::::::::::
    ' :: Sign packets ::
    ' ::::::::::::::::::
    If (LCase(Parse(0)) = "playersign") Then
        Call ShowSign(Parse(1), Parse(2))
        Exit Sub
    End If
    
    
    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnitem" Then
        n = Val(Parse(1))
        
        MapItem(n).num = Val(Parse(2))
        MapItem(n).value = Val(Parse(3))
        MapItem(n).Dur = Val(Parse(4))
        MapItem(n).x = Val(Parse(5))
        MapItem(n).y = Val(Parse(6))
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
        For i = 1 To MAX_ITEMS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Item(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        frmIndex.SetFocus
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: quest editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "questeditor") Then
        InquestEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_QUESTS
            frmIndex.lstIndex.AddItem i & ": Quest " & i
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        frmIndex.SetFocus
        Exit Sub
    End If
    
    
    ' ::::::::::::::::::::::::
    ' :: sign editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "signeditor") Then
        InSignEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SIGNS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Signs(i).header)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        frmIndex.SetFocus
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: sign editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "prayereditor") Then
        InPrayerEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SPELLS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Prayer(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        frmIndex.SetFocus
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateitem") Then
        n = Val(Parse(1))
        If n = 0 Then
            Exit Sub
        End If
        ' Update the item
        Item(n).Name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).type = Val(Parse(4))
        Item(n).Data1 = 0
        Item(n).Data2 = 0
        Item(n).Data3 = 0
        Item(n).Description = Parse(5)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update sign packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatesign") Then
        n = Val(Parse(1))
        
        ' Update the item
        Signs(n).header = Parse(2)
        Signs(n).Msg = Parse(3)
        Exit Sub
    End If
       
    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "edititem") Then
        n = Val(Parse(1))
        
        ' Update the item
        Item(n).Name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Item(n).BaseDamage = Val(Parse(8))
        Item(n).str = Val(Parse(9))
        Item(n).intel = Val(Parse(10))
        Item(n).dex = Val(Parse(11))
        Item(n).con = Val(Parse(12))
        Item(n).wiz = Val(Parse(13))
        Item(n).cha = Val(Parse(14))
        Item(n).Description = Parse(15)
        Item(n).weaponType = Parse(16)
        
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
        MapNpc(n).x = Val(Parse(3))
        MapNpc(n).y = Val(Parse(4))
        MapNpc(n).Dir = Val(Parse(5))
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "npcdead" Then
        n = Val(Parse(1))
        
        MapNpc(n).num = 0
        MapNpc(n).x = 0
        MapNpc(n).y = 0
        MapNpc(n).Dir = 0
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).moving = 0
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
        For i = 1 To MAX_NPCS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Npc(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        frmIndex.SetFocus
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
        Npc(n).sprite = Val(Parse(3))
        Npc(n).SpawnSecs = 0
        Npc(n).Behavior = 0
        Npc(n).Range = 0
        Npc(n).DropChance = 0
        Npc(n).DropItem = 0
        Npc(n).DropItemValue = 0
        Npc(n).str = 0
        Npc(n).DEF = 0
        Npc(n).speed = 0
        Npc(n).MAGI = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Update pet packet ::
    ' :::::::::::::::::::::::
'    If (LCase(Parse(0)) = "updatepet") Then
'        n = Val(Parse(1))
'
'        ' Update the item
'        With Pets(n)
'        .Name = Parse(2)
'        .sprite = Val(Parse(3))
'        .x = Val(Parse(4))
'        .y = Val(Parse(5))
'        .map = Val(Parse(6))
'        .Dir = Val(Parse(7))
'        End With
'
'        Exit Sub
'    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet :: <- Used for item editor admins only
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "editnpc") Then
        n = Val(Parse(1))
        
        ' Update the npc
        Npc(n).Name = Parse(2)
        Npc(n).AttackSay = Parse(3)
        Npc(n).sprite = Val(Parse(4))
        Npc(n).SpawnSecs = Val(Parse(5))
        Npc(n).Behavior = Val(Parse(6))
        Npc(n).Range = Val(Parse(7))
        Npc(n).DropChance = Val(Parse(8))
        Npc(n).DropItem = Val(Parse(9))
        Npc(n).DropItemValue = Val(Parse(10))
        Npc(n).str = Val(Parse(11))
        Npc(n).DEF = Val(Parse(12))
        Npc(n).speed = Val(Parse(13))
        Npc(n).MAGI = Val(Parse(14))
        Npc(n).HP = Val(Parse(15))
        Npc(n).ExpGiven = Val(Parse(16))
        Npc(n).Respawn = Parse(17)
        Npc(n).QuestID = Parse(18)
        Npc(n).opensBank = Parse(19)
        Npc(n).opensShop = CBool(Parse(20))
        Npc(n).type = Parse(21)
        
        'Npc(N).Attack_with_Poison = Parse(18)
        'Npc(N).Poison_length = Parse(19)
        'Npc(N).Poison_vital = Parse(20)
        
        
        ' Initialize the npc editor
        Call NpcEditorInit

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
        For i = 1 To MAX_SHOPS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Shop(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        frmIndex.SetFocus
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
        
        n = 6
        For i = 1 To MAX_TRADES
            
            GiveItem = Val(Parse(n))
            GiveValue = Val(Parse(n + 1))
            GetItem = Val(Parse(n + 2))
            GetValue = Val(Parse(n + 3))
            
            Shop(ShopNum).TradeItem(i).GiveItem = GiveItem
            Shop(ShopNum).TradeItem(i).GiveValue = GiveValue
            Shop(ShopNum).TradeItem(i).GetItem = GetItem
            Shop(ShopNum).TradeItem(i).GetValue = GetValue
            
            n = n + 4
        Next i
        
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
        For i = 1 To MAX_SPELLS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Spell(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        frmIndex.SetFocus
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatespell") Then
        n = Val(Parse(1))
        
        ' Update the spell name
        Spell(n).Name = Parse(2)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Update prayer packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateprayer") Then
        n = Val(Parse(1))
        
        ' Update the prayer name
        Prayer(n).Name = Parse(2)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Update quest packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatequest") Then
        n = Val(Parse(1))
        
        ' Update the prayer name
        Quests(n).ID = Parse(2)
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
        Spell(n).type = Val(Parse(5))
        Spell(n).Data1 = Val(Parse(6))
        Spell(n).Data2 = Val(Parse(7))
        Spell(n).Data3 = Val(Parse(8))
        Spell(n).Sound = Val(Parse(9))
        Spell(n).ManaUse = Val(Parse(10))
                        
        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "editprayer") Then
        n = Val(Parse(1))
        
        ' Update the spell
        Prayer(n).Name = Parse(2)
        Prayer(n).ClassReq = Val(Parse(3))
        Prayer(n).LevelReq = Val(Parse(4))
        Prayer(n).type = Val(Parse(5))
        Prayer(n).Data1 = Val(Parse(6))
        Prayer(n).Data2 = Val(Parse(7))
        Prayer(n).Data3 = Val(Parse(8))
        Prayer(n).Sound = Val(Parse(9))
        Prayer(n).ManaUse = Val(Parse(10))
                        
        ' Initialize the spell editor
        Call PrayerEditorInit

        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit QUEST packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "editquest") Then
        n = Val(Parse(1))
        
        ' Update the quest
        Quests(n).ExpGiven = Parse(2)
        Quests(n).FinishQuestMessage = Parse(3)
        Quests(n).GetItemQuestMsg = Parse(4)
        Quests(n).ItemGiven = Parse(5)
        Quests(n).ItemToObtain = Parse(6)
        Quests(n).ItemValGiven = Parse(7)
        Quests(n).requiredLevel = Parse(8)
        Quests(n).StartQuestMsg = Parse(9)
        Quests(n).GetItemQuestMsg = Parse(10)
        Quests(n).goldGiven = Parse(11)
                        
        ' Initialize the quest editor
        Call QuestEditorInit

        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit BIO packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "editbio") Then
        
        frmBio.txtBio = Parse(1)
        frmBio.txtName = Parse(2)
        frmBio.txtemail = Parse(3)
                        
        Load frmBio
        frmBio.Show
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit sign packet :: <- Used for sign editor admins only
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "editsign") Then
        n = Val(Parse(1))
        
        ' Update the spell
        Signs(n).header = Parse(2)
        Signs(n).Msg = Parse(3)
                        
        ' Initialize the spell editor
        Call SignEditorInit

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
        
        n = 3
        For i = 1 To MAX_TRADES
            GiveItem = Val(Parse(n))
            GiveValue = Val(Parse(n + 1))
            GetItem = Val(Parse(n + 2))
            GetValue = Val(Parse(n + 3))
            
            If GiveItem > 0 And GetItem > 0 Then
                frmTrade.lstTrade.AddItem "Give " & Trim(Shop(ShopNum).Name) & " " & GiveValue & " " & Trim(Item(GiveItem).Name) & " for " & GetValue & " " & Trim(Item(GetItem).Name)
            End If
            n = n + 4
        Next i
        
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
        frmMirage.lblCastType.Caption = "Spells"
        
        frmMirage.picSpellPane.Visible = True
'        For i = 0 To 364 Step 4
'            frmMirage.picSpellPane.Height = i
'            frmMirage.Refresh
'            DoEvents
'        Next i
'        frmMirage.picSpellPane.Height = 364
        frmMirage.lstSpells.Clear
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(i) = Val(Parse(i))
            If Player(MyIndex).Spell(i) <> 0 Then
                frmMirage.lstSpells.AddItem i & ": " & Trim(Spell(Player(MyIndex).Spell(i)).Name)
            Else
                frmMirage.lstSpells.AddItem ""
            End If
        Next i
        
        frmMirage.lstSpells.ListIndex = 0
    End If
    
    ' ::::::::::::::::::::
    ' :: Prayers packet ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "prayers") Then
        frmMirage.lblCastType.Caption = "Prayers"
        frmMirage.picSpellPane.Visible = True
        frmMirage.picPrayerPane.Visible = True
'        For i = 0 To 364 Step 4
'            frmMirage.picSpellPane.Height = i
'            frmMirage.Refresh
'            DoEvents
'        Next i
'        frmMirage.picSpellPane.Height = 364
        frmMirage.lstSpells.Clear
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Prayer(i) = Val(Parse(i))
            If Player(MyIndex).Prayer(i) <> 0 Then
                frmMirage.lstSpells.AddItem i & ": " & Trim(Prayer(Player(MyIndex).Prayer(i)).Name)
            Else
                frmMirage.lstSpells.AddItem ""
            End If
        Next i
        
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
    End If
    
    ' :::::::::::::::::::::
    ' :: item lib packet ::
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "itemlib") Then
        frmItemLib.lblName = Parse(1)
        frmItemLib.lblDesc = Parse(2)
        frmItemLib.lblitemno = Parse(3)
        frmItemLib.Show
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
    If frmMirage.Socket.State = sckConnected Then
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

Sub SendData(ByVal data As String)
'Dim aKet() As Byte
'Dim abCipher() As Byte
'aKet() = StrConv("6$db sYS5&(£'S HseT£ w4uaz5 \gw43 y\4wu", vbFromUnicode)
'Call blf_KeyInit(aKet)
    If IsConnected Then
        'abCipher = blf_BytesEnc(StrConv(data, vbFromUnicode))
        frmMirage.Socket.SendData data ' StrConv(abCipher, vbUnicode)
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

Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal slot As Long, ByVal sprite As Long)
Dim Packet As String

    Packet = "addchar" & SEP_CHAR & Trim(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & slot & SEP_CHAR & sprite & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendDelChar(ByVal slot As Long)
Dim Packet As String
    
    Packet = "delchar" & SEP_CHAR & slot & SEP_CHAR & END_CHAR
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

Sub SayMsg(ByVal text As String)
Dim Packet As String

    Packet = "saymsg" & SEP_CHAR & text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub GlobalMsg(ByVal text As String)
Dim Packet As String

    Packet = "globalmsg" & SEP_CHAR & text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub MassMsg(ByVal text As String)
Dim Packet As String

    Packet = "massmsg" & SEP_CHAR & text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub BroadcastMsg(ByVal text As String)
Dim Packet As String

    Packet = "broadcastmsg" & SEP_CHAR & text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub EmoteMsg(ByVal text As String)
Dim Packet As String

    Packet = "emotemsg" & SEP_CHAR & text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub MapMsg(ByVal text As String)
Dim Packet As String

    Packet = "mapmsg" & SEP_CHAR & text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub PlayerMsg(ByVal text As String, ByVal MsgTo As String)
Dim Packet As String

    Packet = "playermsg" & SEP_CHAR & MsgTo & SEP_CHAR & text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub AdminMsg(ByVal text As String)
Dim Packet As String

    Packet = "adminmsg" & SEP_CHAR & text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerMove()
Dim Packet As String

    Packet = "playermove" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).moving & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerDir()
Dim Packet As String

    Packet = "playerdir" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPlayerRequestNewMap()
Dim Packet As String
    canMoveNow = False
    Packet = "requestnewmap" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR

    Call SendData(Packet)
End Sub

Sub SendMap()
Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long

    Packet = "MAPDATA" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim(map.Name) & SEP_CHAR & map.Revision & SEP_CHAR & map.Moral & SEP_CHAR & map.Up & SEP_CHAR & map.Down & SEP_CHAR & map.Left & SEP_CHAR & map.Right & SEP_CHAR & map.music & SEP_CHAR & map.BootMap & SEP_CHAR & map.BootX & SEP_CHAR & map.BootY & SEP_CHAR & map.Shop & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With map.Tile(x, y)
                Packet = Packet & .Ground & SEP_CHAR & .mask & SEP_CHAR & .Anim & SEP_CHAR & .Fringe & SEP_CHAR & .type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .Data4 & SEP_CHAR & .Data5 & SEP_CHAR & .TileSheet_Ground & SEP_CHAR & .TileSheet_Fringe & SEP_CHAR & .TileSheet_Anim & SEP_CHAR & .TileSheet_Mask & SEP_CHAR
                'Debug.Print .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Fringe & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .Data4 & SEP_CHAR & .Data5 & SEP_CHAR & .TileSheet_Ground & SEP_CHAR & .TileSheet_Fringe & SEP_CHAR & .TileSheet_Anim & SEP_CHAR & .TileSheet_Mask & SEP_CHAR & vbCrLf & vbCrLf
            End With
        Next x
    Next y
    
    For x = 1 To MAX_MAP_NPCS
        Packet = Packet & map.Npc(x) & SEP_CHAR
    Next x
    
    Packet = Packet & map.Respawn & SEP_CHAR & map.Night & SEP_CHAR & map.Bank & SEP_CHAR & map.street & SEP_CHAR & END_CHAR
    
    'x = Int(Len(Packet) / 2)
    'P1 = Mid(Packet, 1, x)
    'P2 = Mid(Packet, x + 1, Len(Packet) - x)
    'Debug.Print P1
    'Debug.Print P2
    'P2 = "MAP0ATA" & P2
    'Call SendData("RANDOM CRAP RANDOM CRAP RANDOM CRAP" & END_CHAR)
    'MsgBox Packet
    Call SendData(Packet)
    'Call SendData(P1)
    'DoEvents
    'Call SendData(P2)
    'DoEvents
    
    
    
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

Sub WarpTo(ByVal MapNum As Long, Optional ByVal M As Long = -1, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1)
Dim Packet As String
    If M >= 0 And x >= 0 And y >= 0 Then
        Packet = "WARPTO" & SEP_CHAR & M & SEP_CHAR & x & SEP_CHAR & y & END_CHAR
    Else
        Packet = "WARPTO" & SEP_CHAR & MapNum & SEP_CHAR & END_CHAR
    End If
    Call SendData(Packet)
End Sub

Sub WarpTo_U(ByVal MapNum As Long, Optional ByVal M As Long = -1, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1)
'Dim Packet As String
'    If M >= 0 And x >= 0 And y >= 0 Then
'        Packet = "WARPTO_U" & SEP_CHAR & M & SEP_CHAR & x & SEP_CHAR & y & END_CHAR
'    Else
'        Packet = "WARPTO_U" & SEP_CHAR & MapNum & SEP_CHAR & END_CHAR
'    End If
'    Call SendData(Packet)
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

Sub SendRequestEditQuest()
Dim Packet As String

    Packet = "REQUESTEDITQUEST" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub



Sub SendSaveItem(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "SAVEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).BaseDamage & SEP_CHAR & Item(ItemNum).str & SEP_CHAR & Item(ItemNum).intel & SEP_CHAR & Item(ItemNum).dex & SEP_CHAR & Item(ItemNum).con & SEP_CHAR & Item(ItemNum).wiz & SEP_CHAR & Item(ItemNum).cha & SEP_CHAR & Item(ItemNum).Description & SEP_CHAR & Item(ItemNum).weaponType & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub
                
Sub SendRequestEditNpc()
Dim Packet As String

    Packet = "REQUESTEDITNPC" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveNpc(ByVal NpcNum As Long)
Dim Packet As String
    
    Packet = "SAVENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).str & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).speed & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).HP & SEP_CHAR & Npc(NpcNum).ExpGiven & SEP_CHAR & Npc(NpcNum).Respawn & SEP_CHAR & Npc(NpcNum).Attack_with_Poison & SEP_CHAR & Npc(NpcNum).Poison_length & SEP_CHAR & Npc(NpcNum).Poison_vital & SEP_CHAR & Npc(NpcNum).QuestID & SEP_CHAR & Npc(NpcNum).opensBank & SEP_CHAR & Npc(NpcNum).opensShop & SEP_CHAR & Npc(NpcNum).type & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLevel()
Dim Packet As String
    
    Packet = "CHECKLEVEL" & SEP_CHAR & END_CHAR
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

Sub SendSaveShop(ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "SAVESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditSpell()
Dim Packet As String

    Packet = "REQUESTEDITSPELL" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditPrayer()
Dim Packet As String

    Packet = "REQUESTEDITPRAYER" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditSign()
Dim Packet As String

    Packet = "REQUESTEDITSIGN" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveSpell(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "SAVESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).Sound & SEP_CHAR & Spell(SpellNum).ManaUse & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSavePrayer(ByVal PrayerNum As Long)
Dim Packet As String

    Packet = "SAVEPRAYER" & SEP_CHAR & PrayerNum & SEP_CHAR & Trim(Prayer(PrayerNum).Name) & SEP_CHAR & Prayer(PrayerNum).ClassReq & SEP_CHAR & Prayer(PrayerNum).LevelReq & SEP_CHAR & Prayer(PrayerNum).type & SEP_CHAR & Prayer(PrayerNum).Data1 & SEP_CHAR & Prayer(PrayerNum).Data2 & SEP_CHAR & Prayer(PrayerNum).Data3 & SEP_CHAR & Prayer(PrayerNum).Sound & SEP_CHAR & Prayer(PrayerNum).ManaUse & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveQuest(ByVal QuestNum As Long)
Dim Packet As String
    Packet = "SAVEQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim(Quests(QuestNum).ExpGiven) & SEP_CHAR & Trim(Quests(QuestNum).FinishQuestMessage) & SEP_CHAR & Trim(Quests(QuestNum).GetItemQuestMsg) & SEP_CHAR & Trim(Quests(QuestNum).ItemGiven) & SEP_CHAR & Trim(Quests(QuestNum).ItemToObtain) & SEP_CHAR & Trim(Quests(QuestNum).ItemValGiven) & SEP_CHAR & Trim(Quests(QuestNum).requiredLevel) & SEP_CHAR & Trim(Quests(QuestNum).StartQuestMsg) & SEP_CHAR & Trim(Quests(QuestNum).GetItemQuestMsg) & SEP_CHAR & Trim(Quests(QuestNum).goldGiven) & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendSaveSign(ByVal SignNum As Long)
Dim Packet As String

    Packet = "SAVESIGN" & SEP_CHAR & SignNum & SEP_CHAR & Trim(Signs(SignNum).header) & SEP_CHAR & Signs(SignNum).Msg & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestEditMap()
Dim Packet As String

    Packet = "REQUESTEDITMAP" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestQuest()
Dim Packet As String

    Packet = "REQUESTQUEST" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendPartyRequest(ByVal Name As String)
Dim Packet As String

    Packet = "PARTY" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendItemLibRequest(ByVal Name As String)
Dim Packet As String

    Packet = "ITEMLIB" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendJoinParty()
Dim Packet As String

    Packet = "JOINPARTY" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendStamina()
Dim Packet As String

    Packet = "STAMINA" & SEP_CHAR & GetPlayerSP(MyIndex) & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendLeaveParty()
Dim Packet As String

    Packet = "LEAVEPARTY" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Sub SendRequestAdminHelp()
Dim Packet As String
    
    Packet = "ADMINMENU" & SEP_CHAR & END_CHAR
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

Sub SendJail(ByVal Name As String, Optional ByVal cell As Long = 7)
Dim Packet As String

Packet = "JAILPLAYER" & SEP_CHAR & Name & SEP_CHAR & cell & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub
