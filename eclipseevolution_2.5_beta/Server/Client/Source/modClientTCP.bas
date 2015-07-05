Attribute VB_Name = "modClientTCP"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean
Public TradePlayer As Long

Sub TcpInit()
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    PlayerBuffer = vbNullString
    
    Dim FileName As String
    FileName = App.Path & "\config.ini"

    frmMirage.Socket.RemoteHost = ReadINI("IPCONFIG", "IP", FileName)
    frmMirage.Socket.RemotePort = val#(ReadINI("IPCONFIG", "PORT", FileName))
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
Dim packet As String
Dim Start As Long

    frmMirage.Socket.GetData Buffer, vbString, DataLength
    
    PlayerBuffer = PlayerBuffer & Buffer
        
    Start = InStr(PlayerBuffer, END_CHAR)
    Do While Start > 0
        packet = Mid$(PlayerBuffer, 1, Start - 1)
        PlayerBuffer = Mid$(PlayerBuffer, Start + 1, Len(PlayerBuffer))
        Start = InStr(PlayerBuffer, END_CHAR)
        If Len(packet) > 0 Then
            Call HandleData(packet)
        End If
    Loop
End Sub

Sub HandleData(ByVal Data As String)
Dim parse() As String
Dim Name As String
Dim Msg As String
Dim Dir As Long
Dim Level As Long
Dim i As Long, n As Long, x As Long, y As Long, p As Long
Dim shopNum As Long
Dim z As Long
Dim strfilename As String
Dim CustomX As Long
Dim CustomY As Long
Dim CustomIndex As Long
Dim customcolour As Long
Dim customsize As Long
Dim customtext As String
Dim casestring As String
Dim packet As String
Dim m As Long
Dim j As Long

    ' Handle Data
    parse = Split(Data, SEP_CHAR)
    
    ' Add the data to the debug window if we are in debug mode
    If Trim$(Command) = "-debug" Then
        If frmDebug.Visible = False Then frmDebug.Visible = True
        Call TextAdd(frmDebug.txtDebug, "((( Processed Packet " & parse$(0) & " )))", True)
    End If
    
    ' :::::::::::::::::::::::
    ' :: Get players stats ::
    ' :::::::::::::::::::::::
    
    casestring = LCase$(parse$(0))
    
    If casestring = "leaveparty211" Then
        For i = 1 To MAX_PARTY_MEMBERS
            Player(MyIndex).Party.Member(i) = 0
        Next i
        Exit Sub
    End If
    
    If casestring = "playerhpreturn" Then
        Player(val(parse(1))).HP = val(parse(2))
        Player(val(parse(1))).MaxHp = val(parse(3))
'        Call MsgBox("player(" & val(parse(1)) & ").hp = " & val(parse(2)))
'        Call BltPlayerBars(val(parse(1)))
        Exit Sub
    End If
    
    If casestring = "maxinfo" Then
        GAME_NAME = Trim$(parse$(1))
        MAX_PLAYERS = val#(parse$(2))
        MAX_ITEMS = val#(parse$(3))
        MAX_NPCS = val#(parse$(4))
        MAX_SHOPS = val#(parse$(5))
        MAX_SPELLS = val#(parse$(6))
        MAX_MAPS = val#(parse$(7))
        MAX_MAP_ITEMS = val#(parse$(8))
        MAX_MAPX = val#(parse$(9))
        MAX_MAPY = val#(parse$(10))
        MAX_EMOTICONS = val#(parse$(11))
        MAX_ELEMENTS = val#(parse$(12))
        paperdoll = val#(parse$(13))
        Spritesize = val#(parse$(14))
        MAX_SCRIPTSPELLS = val#(parse$(15))
        ENCRYPT_PASS = Trim$(parse$(16))
        ENCRYPT_TYPE = Trim$(parse$(17))
        MAX_SKILLS = Trim$(parse$(18))
        MAX_QUESTS = Trim$(parse$(19))
        customplayers = val(parse(20))
        lvl = val(parse(21))
        MAX_PARTY_MEMBERS = val(parse(22))
        STAT1 = parse(23)
        STAT2 = parse(24)
        STAT3 = parse(25)
        STAT4 = parse(26)
        
        If 0 + customplayers > 0 Then
            frmNewChar.Picture4.Visible = False
            frmNewChar.HScroll1.Visible = True
            frmNewChar.HScroll2.Visible = True
            frmNewChar.HScroll3.Visible = True
            frmNewChar.Label14.Visible = True
            frmNewChar.Label11.Visible = True
            frmNewChar.Label12.Visible = True
            frmNewChar.Picture1.Visible = True
        
            If FileExist("GFX\Heads.bmp") Then frmNewChar.iconn(0).Picture = LoadPicture(App.Path & "\GFX\Heads.bmp")
            If FileExist("GFX\Bodys.bmp") Then frmNewChar.iconn(1).Picture = LoadPicture(App.Path & "\GFX\Bodys.bmp")
            If FileExist("GFX\Legs.bmp") Then frmNewChar.iconn(2).Picture = LoadPicture(App.Path & "\GFX\Legs.bmp")
                
                
                If Spritesize = 1 Then
                    frmNewChar.iconn(0).Left = -val(5 * PIC_X)
                    frmNewChar.iconn(0).Top = -val(PIC_Y - 15)
                    
                    frmNewChar.iconn(1).Left = -val(5 * PIC_X)
                    frmNewChar.iconn(1).Top = -val(PIC_Y - 7)
                    
                    frmNewChar.iconn(2).Left = -val(5 * PIC_X)
                    frmNewChar.iconn(2).Top = -val(PIC_Y + 3)
                Else
                    frmNewChar.iconn(0).Left = -val(5 * PIC_X)
                    frmNewChar.iconn(0).Top = -val(PIC_Y)
                    
                    frmNewChar.iconn(1).Left = -val(5 * PIC_X)
                    frmNewChar.iconn(1).Top = -val(PIC_Y)
                    
                    frmNewChar.iconn(2).Left = -val(5 * PIC_X)
                    frmNewChar.iconn(2).Top = -val(PIC_Y)
                End If
        End If
        
        ReDim Map(1 To MAX_MAPS) As MapRec
        ReDim MapAttributeNpc(1 To MAX_ATTRIBUTE_NPCS, 0 To MAX_MAPX, 0 To MAX_MAPY) As MapNpcRec
        ReDim SaveMapAttributeNpc(1 To MAX_ATTRIBUTE_NPCS, 0 To MAX_MAPX, 0 To MAX_MAPY) As MapNpcRec
        ReDim Player(1 To MAX_PLAYERS) As PlayerRec
        ReDim item(1 To MAX_ITEMS) As ItemRec
        ReDim Npc(1 To MAX_NPCS) As NpcRec
        ReDim MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim Shop(1 To MAX_SHOPS) As ShopRec
        ReDim Spell(1 To MAX_SPELLS) As SpellRec
        ReDim Element(0 To MAX_ELEMENTS) As ElementRec
        ReDim Bubble(1 To MAX_PLAYERS) As ChatBubble
        ReDim ScriptBubble(1 To MAX_BUBBLES) As ScriptBubble
        ReDim SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim ScriptSpell(1 To MAX_SCRIPTSPELLS) As ScriptSpellAnimRec
        ReDim skill(1 To MAX_SKILLS) As SkillRec
        ReDim Quest(1 To MAX_QUESTS) As QuestRec
        
        
        For i = 1 To MAX_MAPS
            ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
            ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        Next i
        ReDim TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
        ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
        ReDim MapReport(1 To MAX_MAPS) As MapRec
        MAX_SPELL_ANIM = MAX_MAPX * MAX_MAPY
        
        MAX_BLT_LINE = 6
        ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
        ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
        
        For i = 1 To MAX_PLAYERS
            ReDim Player(i).SpellAnim(1 To MAX_SPELL_ANIM) As SpellAnimRec
            ReDim Player(i).SkilLvl(1 To MAX_SKILLS) As Long
            ReDim Player(i).SkilExp(1 To MAX_SKILLS) As Long
        Next i
        
        For i = 0 To MAX_EMOTICONS
            Emoticons(i).Pic = 0
            Emoticons(i).Command = vbNullString
        Next i
        
        Call ClearTempTile
        
        ' Clear out players
        For i = 1 To MAX_PLAYERS
            Call ClearPlayer(i)
        Next i
        
        For i = 1 To MAX_MAPS
            Call LoadMap(i)
        Next i
    
        frmMirage.Caption = Trim$(GAME_NAME)
        App.Title = GAME_NAME
        
        AllDataReceived = True
        
        On Error Resume Next
        Exit Sub
    End If
        
    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    If casestring = "npchp" Then
        n = val#(parse$(1))
 
        MapNpc(n).HP = val#(parse$(2))
        MapNpc(n).MaxHp = val#(parse$(3))
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    If casestring = "attributenpchp" Then
        n = val#(parse$(1))
 
        MapAttributeNpc(n, val#(parse$(4)), val#(parse$(5))).HP = val#(parse$(2))
        MapAttributeNpc(n, val#(parse$(4)), val#(parse$(5))).MaxHp = val#(parse$(3))
        Exit Sub
    End If
    
    If casestring = "mail" Then
        If val(parse(1)) = 1 Then frmNewAccount.txtEmail.Visible = True
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "alertmsg" Then
        frmMirage.Visible = False
        frmSendGetData.Visible = False
        frmMainMenu.Visible = True
        DoEvents

        Msg = parse$(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Plain message packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "plainmsg" Then
        frmSendGetData.Visible = False
        n = val#(parse$(2))
        
        If n = 1 Then frmNewAccount.Show
        If n = 2 Then frmDeleteAccount.Show
        If n = 3 Then frmLogin.Show
        If n = 4 Then frmNewChar.Show
        If n = 5 Then frmChars.Show
        
        Msg = parse$(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
    If casestring = "allchars" Then
    
        n = 1
        
        frmChars.Visible = True
        frmSendGetData.Visible = False
        
        frmChars.lstChars.Clear
        
        For i = 1 To MAX_CHARS
            Name = parse$(n)
            Msg = parse$(n + 1)
            Level = val#(parse$(n + 2))
            
            If Trim$(Name) = vbNullString Then
                frmChars.lstChars.addItem "Free Character Slot"
            Else
                frmChars.lstChars.addItem Name & " a level " & Level & " " & Msg
            End If
            
            n = n + 3
        Next i
        
        frmChars.lstChars.ListIndex = 0

        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    If casestring = "loginok" Then
        ' Now we can receive game data
        MyIndex = val#(parse$(1))
        
        frmSendGetData.Visible = True
        frmChars.Visible = False

        ReDim Player(MyIndex).SkilLvl(1 To MAX_SKILLS) As Long
        ReDim Player(MyIndex).SkilExp(1 To MAX_SKILLS) As Long
        ReDim Player(MyIndex).Party.Member(1 To MAX_PARTY_MEMBERS)

        Call SetStatus("Receiving game data...")
        Exit Sub
    End If


    ' :::::::::::::::::::::::::::::::::
    ' ::     News Recieved packet    ::
    ' :::::::::::::::::::::::::::::::::
    If casestring = "news" Then
        Call WriteINI("DATA", "News", parse$(1), (App.Path & "\News.ini"))
        Call WriteINI("DATA", "Desc", parse$(5), (App.Path & "\News.ini"))
        Call WriteINI("COLOR", "Red", CInt(parse(2)), (App.Path & "\News.ini"))
        Call WriteINI("COLOR", "Green", CInt(parse(3)), (App.Path & "\News.ini"))
        Call WriteINI("COLOR", "Blue", CInt(parse(4)), (App.Path & "\News.ini"))

        'We just gots teh news, so change the news label
        Call ParseNews
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If casestring = "newcharclasses" Then
        n = 1
        
        ' ClassesOn
        ClassesOn = Int(parse(2))
        
        ' Max classes
        Max_Classes = val#(parse$(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 2

        For i = 0 To Max_Classes
            Class(i).Name = parse$(n)
            
            Class(i).HP = val#(parse$(n + 1))
            Class(i).MP = val#(parse$(n + 2))
            Class(i).SP = val#(parse$(n + 3))
            
            Class(i).STR = val#(parse$(n + 4))
            Class(i).DEF = val#(parse$(n + 5))
            Class(i).speed = val#(parse$(n + 6))
            Class(i).MAGI = val#(parse$(n + 7))
            'Class(i).INTEL = val#(parse$(n + 8))
            Class(i).MaleSprite = val#(parse$(n + 8))
            Class(i).FemaleSprite = val#(parse$(n + 9))
            Class(i).Locked = val#(parse$(n + 10))
            Class(i).desc = parse$(n + 11)
        
        n = n + 12
        Next i
        
        
        ' Used for if the player is creating a new character
        frmNewChar.Visible = True
        frmSendGetData.Visible = False
        
        frmNewChar.cmbClass.Clear
        For i = 0 To Max_Classes
            If Class(i).Locked = 0 Then
                frmNewChar.cmbClass.addItem Trim$(Class(i).Name)
            End If
        Next i
        frmNewChar.cmbClass.ListIndex = 0
        frmNewChar.lblClassDesc = Class(0).desc
        If ClassesOn = 1 Then
            frmNewChar.cmbClass.Visible = True
            frmNewChar.lblClassDesc.Visible = True
        ElseIf ClassesOn = 0 Then
            frmNewChar.cmbClass.Visible = False
            frmNewChar.lblClassDesc.Visible = False
        End If
        
        
        frmNewChar.lblHP.Caption = STR(Class(0).HP)
        frmNewChar.lblMP.Caption = STR(Class(0).MP)
        frmNewChar.lblSP.Caption = STR(Class(0).SP)
    
        frmNewChar.lblSTR.Caption = STR(Class(0).STR)
        frmNewChar.lblDEF.Caption = STR(Class(0).DEF)
        frmNewChar.lblSPEED.Caption = STR(Class(0).speed)
        frmNewChar.lblMAGI.Caption = STR(Class(0).MAGI)
        
        frmNewChar.lblClassDesc.Caption = Class(0).desc
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If casestring = "classesdata" Then
        n = 1
        
        ' Max classes
        Max_Classes = val#(parse$(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For i = 0 To Max_Classes
            Class(i).Name = parse$(n)
            
            Class(i).HP = val#(parse$(n + 1))
            Class(i).MP = val#(parse$(n + 2))
            Class(i).SP = val#(parse$(n + 3))
            
            Class(i).STR = val#(parse$(n + 4))
            Class(i).DEF = val#(parse$(n + 5))
            Class(i).speed = val#(parse$(n + 6))
            Class(i).MAGI = val#(parse$(n + 7))
            
            Class(i).Locked = val#(parse$(n + 8))
            Class(i).desc = parse$(n + 9)
            
            n = n + 10
        Next i
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' ::  Game Clock (Time)  ::
    ' :::::::::::::::::::::::::
    If casestring = "gameclock" Then
        Seconds = val(parse$(1))
        Minutes = val(parse$(2))
        Hours = val(parse$(3))
        Gamespeed = val(parse$(4))
        frmMirage.Label4.Caption = "It is now:"
        frmMirage.Label4.Visible = True
    End If
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If casestring = "ingame" Then
        frmSendGetData.Visible = True
        SetStatus "Connected, entering " & GAME_NAME
        InGame = True
        Call GameInit
        Call GameLoop
        year = val(parse(1))
        month = val(parse(2))
        day = val(parse(3))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    If casestring = "playerinv" Then
        n = 2
        z = val#(parse$(1))
        
        For i = 1 To MAX_INV
            Call SetPlayerInvItemNum(z, i, val#(parse$(n)))
            Call SetPlayerInvItemValue(z, i, val#(parse$(n + 1)))
            Call SetPlayerInvItemDur(z, i, val#(parse$(n + 2)))
            
            n = n + 3
        Next i
        
        If z = MyIndex Then Call UpdateVisInv
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If casestring = "playerinvupdate" Then
        n = val#(parse$(1))
        z = val#(parse$(2))
        
        Call SetPlayerInvItemNum(z, n, val#(parse$(3)))
        Call SetPlayerInvItemValue(z, n, val#(parse$(4)))
        Call SetPlayerInvItemDur(z, n, val#(parse$(5)))
        If z = MyIndex Then Call UpdateVisInv
        Exit Sub
    End If
    ' ::::::::::::::::::::::::
   ' :: Player bank packet ::
   ' ::::::::::::::::::::::::
   If casestring = "playerbank" Then
       n = 1
       For i = 1 To MAX_BANK
           Call SetPlayerBankItemNum(MyIndex, i, val#(parse$(n)))
           Call SetPlayerBankItemValue(MyIndex, i, val#(parse$(n + 1)))
           Call SetPlayerBankItemDur(MyIndex, i, val#(parse$(n + 2)))
           
           n = n + 3
       Next i
       
       If frmBank.Visible = True Then Call UpdateBank
       Exit Sub
   End If
   
   ' :::::::::::::::::::::::::::::::
   ' :: Player bank update packet ::
   ' :::::::::::::::::::::::::::::::
   If casestring = "playerbankupdate" Then
       n = val#(parse$(1))
       
       Call SetPlayerBankItemNum(MyIndex, n, val#(parse$(2)))
       Call SetPlayerBankItemValue(MyIndex, n, val#(parse$(3)))
       Call SetPlayerBankItemDur(MyIndex, n, val#(parse$(4)))
       If frmBank.Visible = True Then Call UpdateBank
       Exit Sub
   End If
   
   ' :::::::::::::::::::::::::::::::
   ' :: Player bank open packet ::
   ' :::::::::::::::::::::::::::::::
   
   If casestring = "openbank" Then
       'frmBank.lblBank.Caption = Trim$(Map(GetPlayerMap(MyIndex)).Name)
       frmBank.lstInventory.Clear
       frmBank.lstBank.Clear
       For i = 1 To MAX_INV
           If GetPlayerInvItemNum(MyIndex, i) > 0 Then
               If item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
                   frmBank.lstInventory.addItem i & "> " & Trim$(item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
               Else
                   If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerNecklaceSlot(MyIndex) = i Then
                       frmBank.lstInventory.addItem i & "> " & Trim$(item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                   Else
                       frmBank.lstInventory.addItem i & "> " & Trim$(item(GetPlayerInvItemNum(MyIndex, i)).Name)
                   End If
               End If
           Else
               frmBank.lstInventory.addItem i & "> Empty"
           End If
           DoEvents
       Next i
       
       For i = 1 To MAX_BANK
           If GetPlayerBankItemNum(MyIndex, i) > 0 Then
               If item(GetPlayerBankItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerBankItemNum(MyIndex, i)).Stackable = 1 Then
                   frmBank.lstBank.addItem i & "> " & Trim$(item(GetPlayerBankItemNum(MyIndex, i)).Name) & " (" & GetPlayerBankItemValue(MyIndex, i) & ")"
               Else
                   If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerNecklaceSlot(MyIndex) = i Then
                       frmBank.lstBank.addItem i & "> " & Trim$(item(GetPlayerBankItemNum(MyIndex, i)).Name) & " (worn)"
                   Else
                       frmBank.lstBank.addItem i & "> " & Trim$(item(GetPlayerBankItemNum(MyIndex, i)).Name)
                   End If
               End If
           Else
               frmBank.lstBank.addItem i & "> Empty"
           End If
           DoEvents
       Next i
       frmBank.lstBank.ListIndex = 0
       frmBank.lstInventory.ListIndex = 0
       
       frmBank.Show vbModal
       Exit Sub
   End If
   
   If LCase$(parse$(0)) = "bankmsg" Then
       frmBank.lblMsg.Caption = Trim$(parse$(1))
       Exit Sub
   End If
   
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    If casestring = "playerworneq" Then
 
        z = val#(parse$(1))
        If z <= 0 Then Exit Sub
        Call SetPlayerArmorSlot(z, val#(parse$(2)))
        Call SetPlayerWeaponSlot(z, val#(parse$(3)))
        Call SetPlayerHelmetSlot(z, val#(parse$(4)))
        Call SetPlayerShieldSlot(z, val#(parse$(5)))
        Call SetPlayerLegsSlot(z, val#(parse$(6)))
        Call SetPlayerRingSlot(z, val#(parse$(7)))
        Call SetPlayerNecklaceSlot(z, val#(parse$(8)))
        
        If z = MyIndex Then
            Call UpdateVisInv
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player points packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "playerpoints" Then
        Player(MyIndex).POINTS = val#(parse$(1))
        frmMirage.lblPoints.Caption = val#(parse$(1))
        Exit Sub
    End If

    If casestring = "cussprite" Then
        Player(val(parse(1))).head = val(parse(2))
        Player(val(parse(1))).body = val(parse(3))
        Player(val(parse(1))).leg = val(parse(4))
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    If LCase$(casestring) = "playerhp" Then
        Player(MyIndex).MaxHp = val#(parse$(1))
        Call SetPlayerHP(MyIndex, val#(parse$(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then
            'frmMirage.shpHP.FillColor = RGB(208, 11, 0)
            frmMirage.shpHP.Width = ((GetPlayerHP(MyIndex)) / (GetPlayerMaxHP(MyIndex))) * 150
            frmMirage.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player mp packet ::
    ' ::::::::::::::::::::::
    If casestring = "playermp" Then
        Player(MyIndex).MaxMP = val#(parse$(1))
        Call SetPlayerMP(MyIndex, val#(parse$(2)))
        If GetPlayerMaxMP(MyIndex) > 0 Then
            'frmMirage.shpMP.FillColor = RGB(208, 11, 0)
            frmMirage.shpMP.Width = ((GetPlayerMP(MyIndex)) / (GetPlayerMaxMP(MyIndex))) * 150
            frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
        End If
        Exit Sub
    End If
    
    ' speech bubble parse
    If (casestring = "mapmsg2") Then
        Bubble(val(parse$(2))).Text = parse$(1)
        Bubble(val(parse$(2))).Created = GetTickCount()
        Exit Sub
    End If
    
    ' scriptbubble parse
    If (casestring = "scriptbubble") Then
        ScriptBubble(val(parse$(1))).Text = Trim$(parse$(2))
        ScriptBubble(val(parse$(1))).Map = val(parse$(3))
        ScriptBubble(val(parse$(1))).x = val(parse$(4))
        ScriptBubble(val(parse$(1))).y = val(parse$(5))
        ScriptBubble(val(parse$(1))).Colour = val(parse$(6))
        ScriptBubble(val(parse$(1))).Created = GetTickCount()
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    If casestring = "playersp" Then
       Player(MyIndex).MaxSP = val#(parse$(1))
        Call SetPlayerSP(MyIndex, val#(parse$(2)))
        If GetPlayerMaxSP(MyIndex) > 0 Then
            frmMirage.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Player Stats Packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "playerstatspacket") Then
        Dim SubDef As Long, SubMagi As Long, SubSpeed As Long, SubStr As Long
        SubStr = 0
        SubDef = 0
        SubMagi = 0
        SubSpeed = 0
        
        If GetPlayerWeaponSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddStr
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddDef
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerArmorSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddStr
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddDef
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerShieldSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddStr
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddDef
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerHelmetSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddStr
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddDef
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerLegsSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddStr
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddDef
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerRingSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddStr
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddDef
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerNecklaceSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddStr
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddDef
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddSpeed
        End If
        
        If SubStr > 0 Then
            frmMirage.lblSTR.Caption = val#(parse$(1)) - SubStr & " (+" & SubStr & ")"
        Else
            frmMirage.lblSTR.Caption = val#(parse$(1))
        End If
        If SubDef > 0 Then
            frmMirage.lblDEF.Caption = val#(parse$(2)) - SubDef & " (+" & SubDef & ")"
        Else
            frmMirage.lblDEF.Caption = val#(parse$(2))
        End If
        If SubMagi > 0 Then
            frmMirage.lblMAGI.Caption = val#(parse$(4)) - SubMagi & " (+" & SubMagi & ")"
        Else
            frmMirage.lblMAGI.Caption = val#(parse$(4))
        End If
        If SubSpeed > 0 Then
            frmMirage.lblSPEED.Caption = val#(parse$(3)) - SubSpeed & " (+" & SubSpeed & ")"
        Else
            frmMirage.lblSPEED.Caption = val#(parse$(3))
        End If
        frmMirage.lblEXP.Caption = val#(parse$(6)) & " / " & val#(parse$(5))
        
        frmMirage.shpTNL.Width = (((val(parse$(6))) / (val(parse$(5)))) * 150)
        frmMirage.lblLevel.Caption = val#(parse$(7))
        Player(MyIndex).Level = val#(parse$(7))
        
        Exit Sub
    End If
                
    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If casestring = "playerdata" Then
    Dim a As Long
        i = val#(parse$(1))
        Call SetPlayerName(i, parse$(2))
        Call SetPlayerSprite(i, val#(parse$(3)))
        Call SetPlayerMap(i, val#(parse$(4)))
        Call SetPlayerX(i, val#(parse$(5)))
        Call SetPlayerY(i, val#(parse$(6)))
        Call SetPlayerDir(i, val#(parse$(7)))
        Call SetPlayerAccess(i, val#(parse$(8)))
        Call SetPlayerPK(i, val#(parse$(9)))
        Call SetPlayerGuild(i, parse$(10))
        Call SetPlayerGuildAccess(i, val#(parse$(11)))
        Call SetPlayerClass(i, val#(parse$(12)))
        Call SetPlayerHead(i, val#(parse$(13)))
        Call SetPlayerBody(i, val#(parse$(14)))
        Call SetPlayerLeg(i, val#(parse$(15)))
        Call SetPlayerPaperdoll(i, val#(parse$(16)))
        Call SetPlayerLevel(i, val#(parse$(17)))
        a = 18
        For j = 1 To MAX_SKILLS
            Call SetPlayerSkillLvl(i, j, val#(parse$(a)))
            'Call SetPlayerSkillExp(i, j, val#(parse$(a + 1)))
            a = a + 2
        Next j


        ' Make sure they aren't walking
        Player(i).Moving = 0
        Player(i).xOffset = 0
        Player(i).yOffset = 0
        
        ' Check if the player is the client player, and if so reset directions
        If i = MyIndex Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player Level Packet  ::
    ' ::::::::::::::::::::::::::
    If casestring = "playerlevel" Then
        n = val#(parse$(1))
        Player(n).Level = val#(parse$(2))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Update Sprite Packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "updatesprite" Then
        i = val#(parse$(1))
        Call SetPlayerSprite(i, val#(parse$(1)))
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "playermove") Then
            i = val#(parse$(1))
            x = val#(parse$(2))
            y = val#(parse$(3))
            Dir = val#(parse$(4))
            n = val#(parse$(5))
    
            If Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Sub
    
            Call SetPlayerX(i, x)
            Call SetPlayerY(i, y)
            Call SetPlayerDir(i, Dir)
                    
            Player(i).xOffset = 0
            Player(i).yOffset = 0
            Player(i).Moving = n
            
            ' Er... what? -Pickle
            'Select Case GetPlayerDir(i)
        '    Case DIR_UP
           '        Player(i).YOffset = PIC_Y
           '        Player(i).XOffset = PIC_X * -1
           '    Case DIR_DOWN
            '        Player(i).YOffset = PIC_Y * -1
            '        Player(i).XOffset = PIC_X
            '    Case DIR_LEFT
            '        Player(i).XOffset = PIC_X
            '        Player(i).YOffset = PIC_Y * -1
            '    Case DIR_RIGHT
            '        Player(i).XOffset = PIC_X * -1
            '        Player(i).YOffset = PIC_Y
            'End Select
            
            'Replaced with the one from TE.
            Select Case GetPlayerDir(i)
            Case DIR_UP
                Player(i).yOffset = PIC_Y
            Case DIR_DOWN
                Player(i).yOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(i).xOffset = PIC_X
            Case DIR_RIGHT
                Player(i).xOffset = PIC_X * -1
            End Select
        
            Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "npcmove") Then
        i = val#(parse$(1))
        x = val#(parse$(2))
        y = val#(parse$(3))
        Dir = val#(parse$(4))
        n = val#(parse$(5))

        MapNpc(i).x = x
        MapNpc(i).y = y
        MapNpc(i).Dir = Dir
        MapNpc(i).xOffset = 0
        MapNpc(i).yOffset = 0
        MapNpc(i).Moving = 1
        
        If n <> 1 Then
            Select Case MapNpc(i).Dir
                Case DIR_UP
                    MapNpc(i).yOffset = PIC_Y * val(n - 1)
                Case DIR_DOWN
                    MapNpc(i).yOffset = PIC_Y * -n
                Case DIR_LEFT
                    MapNpc(i).xOffset = PIC_X * val(n - 1)
                Case DIR_RIGHT
                    MapNpc(i).xOffset = PIC_X * -n
            End Select
        Else
            Select Case MapNpc(i).Dir
                Case DIR_UP
                    MapNpc(i).yOffset = PIC_Y
                Case DIR_DOWN
                    MapNpc(i).yOffset = PIC_Y * -1
                Case DIR_LEFT
                    MapNpc(i).xOffset = PIC_X
                Case DIR_RIGHT
                    MapNpc(i).xOffset = PIC_X * -1
            End Select
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "attributenpcmove") Then
        i = val#(parse$(1))
        x = val#(parse$(2))
        y = val#(parse$(3))
        Dir = val#(parse$(4))
        n = val#(parse$(5))

        MapAttributeNpc(i, val#(parse$(6)), val#(parse$(7))).x = x
        MapAttributeNpc(i, val#(parse$(6)), val#(parse$(7))).y = y
        MapAttributeNpc(i, val#(parse$(6)), val#(parse$(7))).Dir = Dir
        MapAttributeNpc(i, val#(parse$(6)), val#(parse$(7))).xOffset = 0
        MapAttributeNpc(i, val#(parse$(6)), val#(parse$(7))).yOffset = 0
        MapAttributeNpc(i, val#(parse$(6)), val#(parse$(7))).Moving = n
        
        Select Case MapAttributeNpc(i, val#(parse$(6)), val#(parse$(7))).Dir
            Case DIR_UP
                MapAttributeNpc(i, val#(parse$(6)), val#(parse$(7))).yOffset = PIC_Y
            Case DIR_DOWN
                MapAttributeNpc(i, val#(parse$(6)), val#(parse$(7))).yOffset = PIC_Y * -1
            Case DIR_LEFT
                MapAttributeNpc(i, val#(parse$(6)), val#(parse$(7))).xOffset = PIC_X
            Case DIR_RIGHT
                MapAttributeNpc(i, val#(parse$(6)), val#(parse$(7))).xOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (casestring = "playerdir") Then
        i = val#(parse$(1))
        Dir = val#(parse$(2))
        
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Sub

        Call SetPlayerDir(i, Dir)
        
        Player(i).xOffset = 0
        Player(i).yOffset = 0
        Player(i).MovingH = 0
        Player(i).MovingV = 0
        Player(i).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "npcdir") Then
        i = val#(parse$(1))
        Dir = val#(parse$(2))
        MapNpc(i).Dir = Dir
        
        MapNpc(i).xOffset = 0
        MapNpc(i).yOffset = 0
        MapNpc(i).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "attributenpcdir") Then
        i = val#(parse$(1))
        Dir = val#(parse$(2))
        MapAttributeNpc(i, val#(parse$(3)), val#(parse$(4))).Dir = Dir
        
        MapAttributeNpc(i, val#(parse$(3)), val#(parse$(4))).xOffset = 0
        MapAttributeNpc(i, val#(parse$(3)), val#(parse$(4))).yOffset = 0
        MapAttributeNpc(i, val#(parse$(3)), val#(parse$(4))).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (casestring = "playerxy") Then
        x = val#(parse$(1))
        y = val#(parse$(2))
        
        Call SetPlayerX(MyIndex, x)
        Call SetPlayerY(MyIndex, y)
        
        ' Make sure they aren't walking
        Player(MyIndex).Moving = 0
        Player(MyIndex).xOffset = 0
        Player(MyIndex).yOffset = 0
        
        Exit Sub
    End If
    
    If LCase$(parse(0)) = "removemembers" Then
            For n = 1 To MAX_PARTY_MEMBERS
                Player(MyIndex).Party.Member(n) = 0
            Next n
        Exit Sub
    End If

    If LCase$(parse(0)) = "updatemembers" Then
            Player(MyIndex).Party.Member(val(parse(1))) = val(parse(2))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "attack") Then
        i = val#(parse$(1))
        
        ' Set player to attacking
        Player(i).Attacking = 1
        Player(i).AttackTimer = GetTickCount
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (casestring = "npcattack") Then
        i = val#(parse$(1))
        
        ' Set player to attacking
        MapNpc(i).Attacking = 1
        MapNpc(i).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (casestring = "attributenpcattack") Then
        i = val#(parse$(1))
        
        ' Set player to attacking
        MapAttributeNpc(i, val#(parse$(2)), val#(parse$(3))).Attacking = 1
        MapAttributeNpc(i, val#(parse$(2)), val#(parse$(3))).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "checkformap") Then
        ' Erase all players except self
   '     For i = 1 To MAX_PLAYERS
   '         If i <> MyIndex Then
   '             Call SetPlayerMap(i, 0)
   '         End If
   '     Next i
        
        ' Erase all temporary tile values
        Call ClearTempTile
        '!!!

        ' Get map num
        x = val#(parse$(1))
        
        ' Get revision
        y = val#(parse$(2))
        
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
    End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    
    If casestring = "mapdata" Then
        n = 1
        
        Map(val(parse$(1))).Name = parse$(n + 1)
        Map(val(parse$(1))).Revision = val#(parse$(n + 2))
        Map(val(parse$(1))).Moral = val#(parse$(n + 3))
        Map(val(parse$(1))).Up = val#(parse$(n + 4))
        Map(val(parse$(1))).Down = val#(parse$(n + 5))
        Map(val(parse$(1))).Left = val#(parse$(n + 6))
        Map(val(parse$(1))).Right = val#(parse$(n + 7))
        Map(val(parse$(1))).Music = parse$(n + 8)
        Map(val(parse$(1))).BootMap = val#(parse$(n + 9))
        Map(val(parse$(1))).BootX = val#(parse$(n + 10))
        Map(val(parse$(1))).BootY = val#(parse$(n + 11))
        Map(val(parse$(1))).Indoors = val#(parse$(n + 12))
        Map(val(parse$(1))).Weather = val#(parse$(n + 13))
        
        n = n + 14
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(val(parse$(1))).Tile(x, y).Ground = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).Mask = val#(parse$(n + 1))
                Map(val(parse$(1))).Tile(x, y).Anim = val#(parse$(n + 2))
                Map(val(parse$(1))).Tile(x, y).Mask2 = val#(parse$(n + 3))
                Map(val(parse$(1))).Tile(x, y).M2Anim = val#(parse$(n + 4))
                Map(val(parse$(1))).Tile(x, y).Fringe = val#(parse$(n + 5))
                Map(val(parse$(1))).Tile(x, y).FAnim = val#(parse$(n + 6))
                Map(val(parse$(1))).Tile(x, y).Fringe2 = val#(parse$(n + 7))
                Map(val(parse$(1))).Tile(x, y).F2Anim = val#(parse$(n + 8))
                Map(val(parse$(1))).Tile(x, y).Type = val#(parse$(n + 9))
                Map(val(parse$(1))).Tile(x, y).Data1 = val#(parse$(n + 10))
                Map(val(parse$(1))).Tile(x, y).Data2 = val#(parse$(n + 11))
                Map(val(parse$(1))).Tile(x, y).Data3 = val#(parse$(n + 12))
                Map(val(parse$(1))).Tile(x, y).String1 = parse$(n + 13)
                Map(val(parse$(1))).Tile(x, y).String2 = parse$(n + 14)
                Map(val(parse$(1))).Tile(x, y).String3 = parse$(n + 15)
                Map(val(parse$(1))).Tile(x, y).light = val#(parse$(n + 16))
                Map(val(parse$(1))).Tile(x, y).GroundSet = val#(parse$(n + 17))
                Map(val(parse$(1))).Tile(x, y).MaskSet = val#(parse$(n + 18))
                Map(val(parse$(1))).Tile(x, y).AnimSet = val#(parse$(n + 19))
                Map(val(parse$(1))).Tile(x, y).Mask2Set = val#(parse$(n + 20))
                Map(val(parse$(1))).Tile(x, y).M2AnimSet = val#(parse$(n + 21))
                Map(val(parse$(1))).Tile(x, y).FringeSet = val#(parse$(n + 22))
                Map(val(parse$(1))).Tile(x, y).FAnimSet = val#(parse$(n + 23))
                Map(val(parse$(1))).Tile(x, y).Fringe2Set = val#(parse$(n + 24))
                Map(val(parse$(1))).Tile(x, y).F2AnimSet = val#(parse$(n + 25))
                n = n + 26
            Next x
        Next y
        
        For x = 1 To 15
            Map(val(parse$(1))).Npc(x) = val#(parse$(n))
            n = n + 1
        Next x

        ' Save the map
        Call SaveLocalMap(val(parse$(1)))
        
        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmMapEditor.Visible = False
            frmMapEditor.Visible = False
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
        
    If casestring = "tilecheck" Then
     n = 5
     x = val#(parse$(2))
     y = val#(parse$(3))
     
     Select Case val#(parse$(4))
     Case 0
                Map(val(parse$(1))).Tile(x, y).Ground = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).GroundSet = val#(parse$(n + 1))
     Case 1
                Map(val(parse$(1))).Tile(x, y).Mask = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).MaskSet = val#(parse$(n + 1))
     Case 2
                Map(val(parse$(1))).Tile(x, y).Anim = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).AnimSet = val#(parse$(n + 1))
     Case 3
                Map(val(parse$(1))).Tile(x, y).Mask2 = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).Mask2Set = val#(parse$(n + 1))
     Case 4
                Map(val(parse$(1))).Tile(x, y).M2Anim = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).M2AnimSet = val#(parse$(n + 1))
     Case 5
                Map(val(parse$(1))).Tile(x, y).Fringe = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).FringeSet = val#(parse$(n + 1))
     Case 6
                Map(val(parse$(1))).Tile(x, y).FAnim = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).FAnimSet = val#(parse$(n + 1))
     Case 7
                Map(val(parse$(1))).Tile(x, y).Fringe2 = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).Fringe2Set = val#(parse$(n + 1))
     Case 8
                Map(val(parse$(1))).Tile(x, y).F2Anim = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).F2AnimSet = val#(parse$(n + 1))
     Case 9
                Map(val(parse$(1))).Tile(x, y).Floor = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).FloorSet = val#(parse$(n + 1))
     End Select
        Call SaveLocalMap(val(parse$(1)))
    End If
    
    If casestring = "tilecheckattribute" Then
     n = 5
     x = val#(parse$(2))
     y = val#(parse$(3))
     
                Map(val(parse$(1))).Tile(x, y).Type = val#(parse$(n - 1))
                Map(val(parse$(1))).Tile(x, y).Data1 = val#(parse$(n))
                Map(val(parse$(1))).Tile(x, y).Data2 = val#(parse$(n + 1))
                Map(val(parse$(1))).Tile(x, y).Data3 = val#(parse$(n + 2))
                Map(val(parse$(1))).Tile(x, y).String1 = parse$(n + 3)
                Map(val(parse$(1))).Tile(x, y).String2 = parse$(n + 4)
                Map(val(parse$(1))).Tile(x, y).String3 = parse$(n + 5)
        Call SaveLocalMap(val(parse$(1)))
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If casestring = "mapitemdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).num = val#(parse$(n))
            SaveMapItem(i).Value = val#(parse$(n + 1))
            SaveMapItem(i).Dur = val#(parse$(n + 2))
            SaveMapItem(i).x = val#(parse$(n + 3))
            SaveMapItem(i).y = val#(parse$(n + 4))
            
            n = n + 5
        Next i
        
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If casestring = "mapnpcdata" Then
        n = 1
        
        For i = 1 To 15
            SaveMapNpc(i).num = val#(parse$(n))
            SaveMapNpc(i).x = val#(parse$(n + 1))
            SaveMapNpc(i).y = val#(parse$(n + 2))
            SaveMapNpc(i).Dir = val#(parse$(n + 3))
            
            n = n + 4
        Next i
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If casestring = "mapattributenpcdata" Then
        n = 3
        
        x = val#(parse$(1))
        y = val#(parse$(2))
        
        For i = 1 To MAX_ATTRIBUTE_NPCS
            'If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                'If i <= Map(GetPlayerMap(MyIndex)).Tile(x, y).Data2 Then
                    SaveMapAttributeNpc(i, x, y).num = val#(parse$(n))
                    SaveMapAttributeNpc(i, x, y).x = val#(parse$(n + 1))
                    SaveMapAttributeNpc(i, x, y).y = val#(parse$(n + 2))
                    SaveMapAttributeNpc(i, x, y).Dir = val#(parse$(n + 3))
    
                    n = n + 4
                'End If
            'End If
        Next i
        
        Exit Sub
    End If
    
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If casestring = "mapdone" Then
        'Map = SaveMap
        
        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i
        
        GettingMap = False
        
        ' Play music
        If Trim$(Map(GetPlayerMap(MyIndex)).Music) <> "None" Then
                  Call MapMusic(Map(GetPlayerMap(MyIndex)).Music)
        End If
        
        If GameWeather = WEATHER_RAINING Then
        Call PlayBGS("rain.mp3")
        End If
        If GameWeather = WEATHER_THUNDER Then
        Call PlayBGS("thunder.mp3")
        End If
        Call SendData("mapdone2")
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (casestring = "saymsg") Or (casestring = "broadcastmsg") Or (casestring = "globalmsg") Or (casestring = "playermsg") Or (casestring = "mapmsg") Or (casestring = "adminmsg") Then
        Call AddText(parse$(1), val#(parse$(2)))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If casestring = "spawnitem" Then
        n = val#(parse$(1))
        
        MapItem(n).num = val#(parse$(2))
        MapItem(n).Value = val#(parse$(3))
        MapItem(n).Dur = val#(parse$(4))
        MapItem(n).x = val#(parse$(5))
        MapItem(n).y = val#(parse$(6))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "itemeditor") Then
        InItemsEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_ITEMS
            frmIndex.lstIndex.addItem i & ": " & Trim$(item(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updateitem") Then
        n = val#(parse$(1))
        
        ' Update the item
        item(n).Name = parse$(2)
        item(n).Pic = val#(parse$(3))
        item(n).Type = val#(parse$(4))
        item(n).Data1 = val#(parse$(5))
        item(n).Data2 = val#(parse$(6))
        item(n).Data3 = val#(parse$(7))
        item(n).StrReq = val#(parse$(8))
        item(n).DefReq = val#(parse$(9))
        item(n).SpeedReq = val#(parse$(10))
        item(n).ClassReq = val#(parse$(11))
        item(n).AccessReq = val#(parse$(12))
        
        item(n).AddHP = val#(parse$(13))
        item(n).AddMP = val#(parse$(14))
        item(n).AddSP = val#(parse$(15))
        item(n).AddStr = val#(parse$(16))
        item(n).AddDef = val#(parse$(17))
        item(n).AddMagi = val#(parse$(18))
        item(n).AddSpeed = val#(parse$(19))
        item(n).AddEXP = val#(parse$(20))
        item(n).desc = parse$(21)
        item(n).AttackSpeed = val#(parse$(22))
        item(n).Price = val#(parse$(23))
        item(n).Stackable = val#(parse$(24))
        item(n).Bound = val#(parse$(25))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (casestring = "edititem") Then
        n = val#(parse$(1))
        
        ' Update the item
        item(n).Name = parse$(2)
        item(n).Pic = val#(parse$(3))
        item(n).Type = val#(parse$(4))
        item(n).Data1 = val#(parse$(5))
        item(n).Data2 = val#(parse$(6))
        item(n).Data3 = val#(parse$(7))
        item(n).StrReq = val#(parse$(8))
        item(n).DefReq = val#(parse$(9))
        item(n).SpeedReq = val#(parse$(10))
        item(n).ClassReq = val#(parse$(11))
        item(n).AccessReq = val#(parse$(12))
        
        item(n).AddHP = val#(parse$(13))
        item(n).AddMP = val#(parse$(14))
        item(n).AddSP = val#(parse$(15))
        item(n).AddStr = val#(parse$(16))
        item(n).AddDef = val#(parse$(17))
        item(n).AddMagi = val#(parse$(18))
        item(n).AddSpeed = val#(parse$(19))
        item(n).AddEXP = val#(parse$(20))
        item(n).desc = parse$(21)
        item(n).AttackSpeed = val#(parse$(22))
        item(n).Price = val#(parse$(23))
        item(n).Stackable = val#(parse$(24))
        item(n).Bound = val#(parse$(25))
        
        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: mouse packet  ::
    ' :::::::::::::::::::
    If (casestring = "mouse") Then
        Player(MyIndex).input = 1
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' ::Weather Packet::
    ' ::::::::::::::::::
    If (casestring = "mapweather") Then
        If 0 + val(parse$(1)) <> 0 Then
            Map(val(parse$(1))).Weather = val(parse$(2))
            If val(parse$(1)) = 2 Then
                frmMirage.tmrSnowDrop.Interval = val(parse$(3))
            ElseIf val(parse$(1)) = 1 Then
                frmMirage.tmrRainDrop.Interval = val(parse$(3))
            End If
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If casestring = "spawnnpc" Then
        n = val#(parse$(1))
        
        MapNpc(n).num = val#(parse$(2))
        MapNpc(n).x = val#(parse$(3))
        MapNpc(n).y = val#(parse$(4))
        MapNpc(n).Dir = val#(parse$(5))
        MapNpc(n).Big = val#(parse$(6))
        
        ' Client use only
        MapNpc(n).xOffset = 0
        MapNpc(n).yOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If casestring = "spawnattributenpc" Then
        n = val#(parse$(1))
        
        x = val#(parse$(7))
        y = val#(parse$(8))
        
        MapAttributeNpc(n, x, y).num = val#(parse$(2))
        MapAttributeNpc(n, x, y).x = val#(parse$(3))
        MapAttributeNpc(n, x, y).y = val#(parse$(4))
        MapAttributeNpc(n, x, y).Dir = val#(parse$(5))
        MapAttributeNpc(n, x, y).Big = val#(parse$(6))
        
        ' Client use only
        MapAttributeNpc(n, x, y).xOffset = 0
        MapAttributeNpc(n, x, y).yOffset = 0
        MapAttributeNpc(n, x, y).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If casestring = "npcdead" Then
        n = val#(parse$(1))
        
        MapNpc(n).num = 0
        MapNpc(n).x = 0
        MapNpc(n).y = 0
        MapNpc(n).Dir = 0
        
        ' Client use only
        MapNpc(n).xOffset = 0
        MapNpc(n).yOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If casestring = "attributenpcdead" Then
        n = val#(parse$(1))
        
        MapAttributeNpc(n, val#(parse$(2)), val#(parse$(3))).num = 0
        MapAttributeNpc(n, val#(parse$(2)), val#(parse$(3))).x = 0
        MapAttributeNpc(n, val#(parse$(2)), val#(parse$(3))).y = 0
        MapAttributeNpc(n, val#(parse$(2)), val#(parse$(3))).Dir = 0
        
        ' Client use only
        MapAttributeNpc(n, val#(parse$(2)), val#(parse$(3))).xOffset = 0
        MapAttributeNpc(n, val#(parse$(2)), val#(parse$(3))).yOffset = 0
        MapAttributeNpc(n, val#(parse$(2)), val#(parse$(3))).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    If (casestring = "npceditor") Then
        InNpcEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_NPCS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Npc(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (casestring = "updatenpc") Then
        n = val#(parse$(1))
        
        ' Update the item
        Npc(n).Name = parse$(2)
        Npc(n).AttackSay = vbNullString
        Npc(n).Sprite = val#(parse$(3))
        Npc(n).Spritesize = val#(parse$(4))
        ' That's all well and good but it also resets our NPC - Pickle
        'Npc(n).SpawnSecs = 0
        'Npc(n).Behavior = 0
        'Npc(n).Range = 0
        'For i = 1 To MAX_NPC_DROPS
        '    Npc(n).ItemNPC(i).chance = 0
        '    Npc(n).ItemNPC(i).ItemNum = 0
        '    Npc(n).ItemNPC(i).ItemValue = 0
        'Next i
        'Npc(n).STR = 0
        'Npc(n).DEF = 0
        'Npc(n).speed = 0
        'Npc(n).MAGI = 0
        Npc(n).Big = val#(parse$(5))
        Npc(n).MaxHp = val#(parse$(6))
        'Npc(n).Exp = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet :: <- Used for item editor admins only
    ' :::::::::::::::::::::
    If (casestring = "editnpc") Then
        n = val#(parse$(1))
        
        ' Update the npc
        Npc(n).Name = parse$(2)
        Npc(n).AttackSay = parse$(3)
        Npc(n).Sprite = val#(parse$(4))
        Npc(n).SpawnSecs = val#(parse$(5))
        Npc(n).Behavior = val#(parse$(6))
        Npc(n).Range = val#(parse$(7))
        Npc(n).STR = val#(parse$(8))
        Npc(n).DEF = val#(parse$(9))
        Npc(n).speed = val#(parse$(10))
        Npc(n).MAGI = val#(parse$(11))
        Npc(n).Big = val#(parse$(12))
        Npc(n).MaxHp = val#(parse$(13))
        Npc(n).Exp = val#(parse$(14))
        Npc(n).SpawnTime = val#(parse$(15))
        Npc(n).Element = val#(parse$(16))
        Npc(n).Spritesize = val(parse(17))
        
       ' Call GlobalMsg("At editnpc..." & Npc(n).Element)
        z = 18
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).chance = val#(parse$(z))
            Npc(n).ItemNPC(i).ItemNum = val#(parse$(z + 1))
            Npc(n).ItemNPC(i).ItemValue = val#(parse$(z + 2))
            z = z + 3
        Next i
        
        ' Initialize the npc editor
        Call NpcEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (casestring = "mapkey") Then
        x = val#(parse$(1))
        y = val#(parse$(2))
        n = val#(parse$(3))
                
        TempTile(x, y).DoorOpen = n
        
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
    If (casestring = "editmap") Then
        Call EditorInit
        Exit Sub
    End If
        ' :::::::::::::::::::::
    ' :: Edit house packet ::
    ' :::::::::::::::::::::
    If (casestring = "edithouse") Then
        Call HouseEditorInit
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit main packet ::
    ' :::::::::::::::::::::
    If (casestring = "main") Then
        frmEditor.RT.Text = parse$(1)
        frmEditor.Visible = True
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "shopeditor") Then
        InShopEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SHOPS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Shop(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updateshop") Then
        n = val#(parse$(1))
        
        ' Update the shop name
        Shop(n).Name = parse$(2)
        Shop(n).FixesItems = val(parse(3))
        Shop(n).BuysItems = val(parse(4))
        Shop(n).ShowInfo = val(parse(5))
        Shop(n).currencyItem = val(parse(6))
        
        m = 7
        'Get shop items
        For i = 1 To MAX_SHOP_ITEMS
            Shop(n).ShopItem(i).ItemNum = val(parse(m))
            Shop(n).ShopItem(i).Amount = val(parse(m + 1))
            Shop(n).ShopItem(i).Price = val(parse(m + 2))
            m = m + 3
        Next i
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    If (casestring = "editshop") Then
        
        shopNum = val#(parse$(1))
        
        ' Update the shop
        Shop(shopNum).Name = parse$(2)
        Shop(shopNum).FixesItems = val(parse$(3))
        Shop(shopNum).BuysItems = val(parse$(4))
        Shop(shopNum).ShowInfo = val(parse$(5))
        Shop(shopNum).currencyItem = val(parse$(6))
        
        m = 7
        For i = 1 To 25
            Shop(shopNum).ShopItem(i).ItemNum = val(parse$(m))
            Shop(shopNum).ShopItem(i).Amount = val(parse$(m + 1))
            Shop(shopNum).ShopItem(i).Price = val(parse$(m + 2))
            m = m + 3
        Next i
        
        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "spelleditor") Then
        InSpellEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SPELLS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Spell(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updatespell") Then
        n = val#(parse$(1))
        
        ' Update the spell name
        Spell(n).Name = parse$(2)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (casestring = "editspell") Then
        n = val#(parse$(1))
        
        ' Update the spell
        Spell(n).Name = parse$(2)
        Spell(n).ClassReq = val#(parse$(3))
        Spell(n).LevelReq = val#(parse$(4))
        Spell(n).Type = val#(parse$(5))
        Spell(n).Data1 = val#(parse$(6))
        Spell(n).Data2 = val#(parse$(7))
        Spell(n).Data3 = val#(parse$(8))
        Spell(n).MPCost = val#(parse$(9))
        Spell(n).Sound = val#(parse$(10))
        Spell(n).Range = val#(parse$(11))
        Spell(n).SpellAnim = val#(parse$(12))
        Spell(n).SpellTime = val#(parse$(13))
        Spell(n).SpellDone = val#(parse$(14))
        Spell(n).AE = val#(parse$(15))
        Spell(n).Big = val#(parse$(16))
        Spell(n).Element = val#(parse$(17))
                        
                        
        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (casestring = "goshop") Then
        shopNum = val#(parse$(1))
        'Show the shop
        Call GoShop(shopNum)
    End If

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (casestring = "spells") Then
        
        frmMirage.picPlayerSpells.Visible = True
        frmMirage.lstSpells.Clear
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(i) = val#(parse$(i))
            If Player(MyIndex).Spell(i) <> 0 Then
                frmMirage.lstSpells.addItem i & ": " & Trim$(Spell(Player(MyIndex).Spell(i)).Name)
            Else
                frmMirage.lstSpells.addItem "--- Slot Free ---"
            End If
        Next i
        
        frmMirage.lstSpells.ListIndex = 0
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (casestring = "weather") Then
        If val#(parse$(1)) = WEATHER_RAINING And GameWeather <> WEATHER_RAINING Then
            Call AddText("You see drops of rain falling from the sky above!", BrightGreen)
            Call PlayBGS("rain.mp3")
        End If
        If val#(parse$(1)) = WEATHER_THUNDER And GameWeather <> WEATHER_THUNDER Then
            Call AddText("You see thunder in the sky above!", BrightGreen)
            Call PlayBGS("thunder.mp3")
        End If
        If val#(parse$(1)) = WEATHER_SNOWING And GameWeather <> WEATHER_SNOWING Then
            Call AddText("You see snow falling from the sky above!", BrightGreen)
        End If
        
        If val#(parse$(1)) = WEATHER_NONE Then
            If GameWeather = WEATHER_RAINING Then
                Call AddText("The rain beings to calm.", BrightGreen)
                Call frmMirage.BGSPlayer.StopMedia
            ElseIf GameWeather = WEATHER_SNOWING Then
                Call AddText("The snow is melting away.", BrightGreen)
            ElseIf GameWeather = WEATHER_THUNDER Then
                Call AddText("The thunder begins to disapear.", BrightGreen)
                Call frmMirage.BGSPlayer.StopMedia
            End If
        End If
        GameWeather = val#(parse$(1))
        RainIntensity = val#(parse$(2))
        If MAX_RAINDROPS <> RainIntensity Then
            MAX_RAINDROPS = RainIntensity
            ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
            ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
        End If
    End If

    ' ::::::::::::::::::::::::::::::::
    ' :: playername coloring packet ::
    ' ::::::::::::::::::::::::::::::::
    If (casestring = "namecolor") Then
        Player(MyIndex).color = val(parse$(1))
    Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: image packet      ::
    ' :::::::::::::::::::::::
    If (LCase$(parse(0)) = "fog") Then
        rec.Top = Int(val(parse$(4)))
        rec.Bottom = Int(val(parse$(5)))
        rec.Left = Int(val(parse$(6)))
        rec.Right = Int(val(parse$(7)))
        Call DD_BackBuffer.BltFast(val(parse$(1)), val(parse$(2)), DD_TileSurf(val(parse$(3))), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: light packet   ::
    ' ::::::::::::::::::::
    If casestring = "lights" Then
        Map(val#(parse$(1))).lights = val#(parse$(2))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Get Online List ::
    ' ::::::::::::::::::::::::::
    If casestring = "onlinelist" Then
    frmMirage.lstOnline.Clear
    
        n = 2
        z = val#(parse$(1))
        For x = n To (z + 1)
            frmMirage.lstOnline.addItem Trim$(parse$(n))
            n = n + 2
        Next x
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
    If casestring = "blitplayerdmg" Then
        DmgDamage = val#(parse$(1))
        NPCWho = val#(parse$(2))
        DmgTime = GetTickCount
        iii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
    If casestring = "blitnpcdmg" Then
        NPCDmgDamage = val#(parse$(1))
        NPCDmgTime = GetTickCount
        ii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::
    ' :: Retrieve the player's inventory ::
    ' :::::::::::::::::::::::::::::::::::::
    If casestring = "pptrading" Then
        frmPlayerTrade.Items1.Clear
        frmPlayerTrade.Items2.Clear
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = vbNullString
            Trading2(i).InvNum = 0
            Trading2(i).InvName = vbNullString
            frmPlayerTrade.Items1.addItem i & ": <Nothing>"
            frmPlayerTrade.Items2.addItem i & ": <Nothing>"
        Next i
        
        frmPlayerTrade.Items1.ListIndex = 0
        
        Call UpdateTradeInventory
        frmPlayerTrade.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "qtrade" Then
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = vbNullString
            Trading2(i).InvNum = 0
            Trading2(i).InvName = vbNullString
        Next i
        
        frmPlayerTrade.Command1.ForeColor = &H0&
        frmPlayerTrade.Command2.ForeColor = &H0&
        
        frmPlayerTrade.Visible = False
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "oncanon" Then
        CanonUsed = 1
        Exit Sub
    End If
    
    If casestring = "canonoff" Then
        CanonUsed = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Disable Time ::
    ' ::::::::::::::::::
    If casestring = "dtime" Then
        If parse(1) = True Then
            frmMirage.Label4.Caption = vbNullString
            frmMirage.GameClock.Caption = vbNullString
            frmMirage.Label4.Visible = False
            frmMirage.GameClock.Visible = False
            frmMirage.tmrGameClock.Enabled = False
        Else
            frmMirage.Label4.Visible = True
            frmMirage.GameClock.Visible = True
            frmMirage.Label4.Caption = "It is now:"
            frmMirage.GameClock.Caption = vbNullString
            frmMirage.tmrGameClock.Enabled = True
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "updatetradeitem" Then
            n = val#(parse$(1))
            
            Trading2(n).InvNum = val#(parse$(2))
            Trading2(n).InvName = parse$(3)
            
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
    If casestring = "trading" Then
        n = val#(parse$(1))
            If n = 0 Then frmPlayerTrade.Command2.ForeColor = &H0&
            If n = 1 Then frmPlayerTrade.Command2.ForeColor = &HFF00&
        Exit Sub
    End If
    
' :::::::::::::::::::::::::
' :: Chat System Packets ::
' :::::::::::::::::::::::::
    If casestring = "ppchatting" Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Label1.Caption = "Chatting With: " & Trim$(Player(val(parse$(1))).Name)

        frmPlayerChat.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    If casestring = "qchat" Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Visible = False
        frmPlayerTrade.Command2.ForeColor = &H8000000F
        frmPlayerTrade.Command1.ForeColor = &H8000000F
        Exit Sub
    End If
    
    If casestring = "sendchat" Then
        Dim s As String
  
        s = vbNewLine & GetPlayerName(val(parse$(2))) & "> " & Trim$(parse$(1))
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
    If casestring = "sound" Then
        s = LCase$(parse$(1))
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
                Call PlaySound("magic" & val#(parse$(2)) & ".wav")
            Case "warp"
                If FileExist("SFX\warp.wav") Then Call PlaySound("warp.wav")
            Case "pain"
                    Call PlaySound("pain.wav")
            Case "soundattribute"
                Call PlaySound(Trim$(parse$(2)))
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Sprite Change Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If casestring = "spritechange" Then
        If val#(parse$(1)) = 1 Then
            i = MsgBox("Would you like to buy this sprite?", 4, "Buying Sprite")
            If i = 6 Then
                Call SendData("buysprite" & SEP_CHAR & END_CHAR)
            End If
        Else
            Call SendData("buysprite" & SEP_CHAR & END_CHAR)
        End If
        Exit Sub
    End If
        ' :::::::::::::::::::::::::::::::::::::::
    ' :: House Buy Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If casestring = "housebuy" Then
        If val#(parse$(1)) = 1 Then
            i = MsgBox("Would you like to buy this house?", 4, "Buying House")
            If i = 6 Then
                Call SendData("buyhouse" & SEP_CHAR & END_CHAR)
            End If
        Else
            Call SendData("buyhouse" & SEP_CHAR & END_CHAR)
        End If
        Exit Sub
    End If
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Change Player Direction Packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If casestring = "changedir" Then
        Player(val(parse$(2))).Dir = val#(parse$(1))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Flash Movie Event Packet ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "flashevent" Then
        If LCase(Mid(Trim$(parse$(1)), 1, 7)) = "http://" Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopBGM
            Call StopSound
            frmFlash.Flash.LoadMovie 0, Trim$(parse$(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        ElseIf FileExist("Flashs\" & Trim$(parse$(1))) = True Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopBGM
            Call StopSound
            frmFlash.Flash.LoadMovie 0, App.Path & "\Flashs\" & Trim$(parse$(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If casestring = "prompt" Then
        i = MsgBox(Trim$(parse$(1)), vbYesNo)
        Call SendData("prompt" & SEP_CHAR & i & SEP_CHAR & val#(parse$(2)) & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If casestring = "querybox" Then
        frmQuery.Label1.Caption = Trim$(parse$(1))
        frmQuery.Label2.Caption = parse$(2)
        frmQuery.Show vbModal
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Emoticon editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "emoticoneditor") Then
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
    If (casestring = "elementeditor") Then
        InElementEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        For i = 0 To MAX_ELEMENTS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Element(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If


    ' :::::::::::::::::::::::::
    ' :: Quest editor packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "questeditor") Then
    
        InQuestEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        For i = 1 To MAX_QUESTS
            frmIndex.lstIndex.addItem i & ": " & Quest(i).Name
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    If (casestring = "editquest") Then
        n = val#(parse$(1))

        Quest(n).Name = Trim(parse(2))
        Quest(n).Pictop = val(parse(3))
        Quest(n).Picleft = val(parse(4))
        
        m = 5
        For j = 0 To MAX_QUEST_LENGHT
            Quest(n).Map(j) = val(parse(m))
            Quest(n).x(j) = val(parse(m + 1))
            Quest(n).y(j) = val(parse(m + 2))
            Quest(n).Npc(j) = val(parse(m + 3))
            Quest(n).Script(j) = val(parse(m + 4))
            Quest(n).ItemTake1num(j) = val(parse(m + 5))
            Quest(n).ItemTake2num(j) = val(parse(m + 6))
            Quest(n).ItemTake1val(j) = val(parse(m + 7))
            Quest(n).ItemTake2val(j) = val(parse(m + 8))
            Quest(n).ItemGive1num(j) = val(parse(m + 9))
            Quest(n).ItemGive2num(j) = val(parse(m + 10))
            Quest(n).ItemGive1val(j) = val(parse(m + 11))
            Quest(n).ItemGive2val(j) = val(parse(m + 12))
            Quest(n).ExpGiven(j) = val(parse(m + 13))
            m = m + 14
        Next j

        Call QuestEditorInit
        Exit Sub
    End If

    If (casestring = "updatequest") Then
        n = val#(parse$(1))

        Quest(n).Name = Trim(parse(2))
        Quest(n).Pictop = val(parse(3))
        Quest(n).Picleft = val(parse(4))
        
        m = 5
        For j = 0 To MAX_QUEST_LENGHT
            Quest(n).Map(j) = val(parse(m))
            Quest(n).x(j) = val(parse(m + 1))
            Quest(n).y(j) = val(parse(m + 2))
            Quest(n).Npc(j) = val(parse(m + 3))
            Quest(n).Script(j) = val(parse(m + 4))
            Quest(n).ItemTake1num(j) = val(parse(m + 5))
            Quest(n).ItemTake2num(j) = val(parse(m + 6))
            Quest(n).ItemTake1val(j) = val(parse(m + 7))
            Quest(n).ItemTake2val(j) = val(parse(m + 8))
            Quest(n).ItemGive1num(j) = val(parse(m + 9))
            Quest(n).ItemGive2num(j) = val(parse(m + 10))
            Quest(n).ItemGive1val(j) = val(parse(m + 11))
            Quest(n).ItemGive2val(j) = val(parse(m + 12))
            Quest(n).ExpGiven(j) = val(parse(m + 13))
            m = m + 14
        Next j
        
        Exit Sub
    End If


    ' ::::::::::::::::::::::::::::
    ' :: Skill editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "skilleditor") Then
    
        InSkillEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        For i = 1 To MAX_SKILLS
            frmIndex.lstIndex.addItem i & ": " & skill(i).Name
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    If (casestring = "editskill") Then
        n = val#(parse$(1))

        skill(n).Name = parse$(2)
        skill(n).Action = parse$(3)
        skill(n).Fail = parse$(4)
        skill(n).Succes = parse$(5)
        skill(n).AttemptName = parse$(6)
        skill(n).Pictop = val#(parse$(7))
        skill(n).Picleft = val#(parse$(8))
        
        m = 9
        For j = 1 To MAX_SKILLS_SHEETS
            skill(n).ItemTake1num(j) = val(parse(m))
            skill(n).ItemTake2num(j) = val(parse(m + 1))
            skill(n).ItemGive1num(j) = val(parse(m + 2))
            skill(n).ItemGive2num(j) = val(parse(m + 3))
            skill(n).minlevel(j) = val(parse(m + 4))
            skill(n).ExpGiven(j) = val(parse(m + 5))
            skill(n).base_chance(j) = val(parse(m + 6))
            skill(n).ItemTake1val(j) = val(parse(m + 7))
            skill(n).ItemTake2val(j) = val(parse(m + 8))
            skill(n).ItemGive1val(j) = val(parse(m + 9))
            skill(n).ItemGive2val(j) = val(parse(m + 10))
            skill(n).itemequiped(j) = val(parse(m + 11))
            m = m + 12
        Next j

        Call skillEditorInit
        Exit Sub
    End If

    If (casestring = "updateskill") Then
        n = val#(parse$(1))

        skill(n).Name = parse$(2)
        skill(n).Action = parse$(3)
        skill(n).Fail = parse$(4)
        skill(n).Succes = parse$(5)
        skill(n).AttemptName = parse$(6)
        skill(n).Pictop = val#(parse$(7))
        skill(n).Picleft = val#(parse$(8))
        
        m = 9
        For j = 1 To MAX_SKILLS_SHEETS
            skill(n).ItemTake1num(j) = val(parse(m))
            skill(n).ItemTake2num(j) = val(parse(m + 1))
            skill(n).ItemGive1num(j) = val(parse(m + 2))
            skill(n).ItemGive2num(j) = val(parse(m + 3))
            skill(n).minlevel(j) = val(parse(m + 4))
            skill(n).ExpGiven(j) = val(parse(m + 5))
            skill(n).base_chance(j) = val(parse(m + 6))
            skill(n).ItemTake1val(j) = val(parse(m + 7))
            skill(n).ItemTake2val(j) = val(parse(m + 8))
            skill(n).ItemGive1val(j) = val(parse(m + 9))
            skill(n).ItemGive2val(j) = val(parse(m + 10))
            skill(n).itemequiped(j) = val(parse(m + 11))
            m = m + 12
        Next j
        
        Exit Sub
    End If
    
    'Update skill EXP
    If (casestring = "skillinfo") Then
        n = val#(parse$(1))
        Player(MyIndex).SkilExp(n) = val#(parse$(2))
        Player(MyIndex).SkilLvl(n) = val#(parse$(3))
        'Update skill window
        frmMirage.Exp(n - currentsheet - 1) = Player(MyIndex).SkilExp(n)
        frmMirage.Level(n - currentsheet - 1) = Player(MyIndex).SkilLvl(n)
        Exit Sub
    End If

    If (casestring = "editelement") Then
        n = val#(parse$(1))

        Element(n).Name = parse$(2)
        Element(n).Strong = val#(parse$(3))
        Element(n).Weak = val#(parse$(4))
        
        Call ElementEditorInit
        Exit Sub
    End If
    
    If (casestring = "updateelement") Then
        n = val#(parse$(1))

        Element(n).Name = parse$(2)
        Element(n).Strong = val#(parse$(3))
        Element(n).Weak = val#(parse$(4))
        Exit Sub
    End If

    If (casestring = "editemoticon") Then
        n = val#(parse$(1))

        Emoticons(n).Command = parse$(2)
        Emoticons(n).Pic = val#(parse$(3))
        
        Call EmoticonEditorInit
        Exit Sub
    End If
    
    If (casestring = "updateemoticon") Then
        n = val#(parse$(1))
        
        Emoticons(n).Command = parse$(2)
        Emoticons(n).Pic = val#(parse$(3))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Arrow editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "arroweditor") Then
        InArrowEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        For i = 1 To MAX_ARROWS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Arrows(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    If (casestring = "updatearrow") Then
        n = val#(parse$(1))
        
        Arrows(n).Name = parse$(2)
        Arrows(n).Pic = val#(parse$(3))
        Arrows(n).Range = val#(parse$(4))
        Arrows(n).Amount = val#(parse$(5))
        Exit Sub
    End If

    If (casestring = "editarrow") Then
        n = val#(parse$(1))

        Arrows(n).Name = parse$(2)
        
        Call ArrowEditorInit
        Exit Sub
    End If
    
    If (casestring = "updatearrow") Then
        n = val#(parse$(1))
        
        Arrows(n).Name = parse$(2)
        Arrows(n).Pic = val#(parse$(3))
        Arrows(n).Range = val#(parse$(4))
        Arrows(n).Amount = val#(parse$(5))
        Exit Sub
    End If

    If (casestring = "hookshot") Then
        n = val#(parse$(1))
        i = val#(parse$(3))
        
        Player(n).HookShotAnim = Arrows(val#(parse$(2))).Pic
        Player(n).HookShotTime = GetTickCount
        Player(n).HookShotToX = val#(parse$(4))
        Player(n).HookShotToY = val#(parse$(5))
        Player(n).HookShotX = GetPlayerX(n)
        Player(n).HookShotY = GetPlayerY(n)
        Player(n).HookShotSucces = val#(parse$(6))
        Player(n).HookShotDir = val#(parse$(3))

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
                        Player(n).Arrow(x).Arrow = 0
                        Exit Sub
                    End If
                End If
        Exit Sub
    End If
    
    If (casestring = "checkarrows") Then
        n = val#(parse$(1))
        z = val#(parse$(2))
        i = val#(parse$(3))
        p = val#(parse$(4))
        
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
                Player(n).Arrow(x).ArrowAmount = p
                
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

    If (casestring = "checksprite") Then
        n = val#(parse$(1))
        
        Player(n).Sprite = val#(parse$(2))
        Exit Sub
    End If
    
    If (casestring = "mapreport") Then
        n = 1
        
        frmMapReport.lstIndex.Clear
        For i = 1 To MAX_MAPS
            frmMapReport.lstIndex.addItem i & ": " & Trim$(parse$(n))
            n = n + 1
        Next i
        
        frmMapReport.Show vbModeless, frmMirage
        Exit Sub
    End If

    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (casestring = "actionname") Then
        frmMirage.Action.addItem "None"
        n = 2
        For i = 1 To val#(parse$(1))
            frmMirage.Action.addItem Trim$(parse$(n))
            n = n + 1
        Next i
        Exit Sub
    End If

    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (casestring = "time") Then
        GameTime = val#(parse$(1))
        If GameTime = TIME_DAY Then
            Call AddText("Day has dawned in this realm.", White)
        Else
            Call AddText("Night has fallen upon the weary eyed nightowls.", White)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::
    ' ::WIERD! Packet::
    ' :::::::::::::::::
    If (casestring = "wierd") Then
        Wierd = val#(parse$(1))
        If Wierd = 1 Then
            Call AddText("The world has gone wierd!", White)
        Else
            Call AddText("Normality has returned", White)
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Spell anim packet ::
    ' :::::::::::::::::::::::
    If (casestring = "spellanim") Then
        Dim SpellNum As Long
        SpellNum = val#(parse$(1))
        
        Spell(SpellNum).SpellAnim = val#(parse$(2))
        Spell(SpellNum).SpellTime = val#(parse$(3))
        Spell(SpellNum).SpellDone = val#(parse$(4))
        Spell(SpellNum).Big = val#(parse$(9))
        
        Player(val(parse$(5))).SpellNum = SpellNum
        
        For i = 1 To MAX_SPELL_ANIM
            If Player(val(parse$(5))).SpellAnim(i).CastedSpell = NO Then
                Player(val(parse$(5))).SpellAnim(i).SpellDone = 0
                Player(val(parse$(5))).SpellAnim(i).SpellVar = 0
                Player(val(parse$(5))).SpellAnim(i).SpellTime = GetTickCount
                Player(val(parse$(5))).SpellAnim(i).TargetType = val#(parse$(6))
                Player(val(parse$(5))).SpellAnim(i).Target = val#(parse$(7))
                Player(val(parse$(5))).SpellAnim(i).CastedSpell = YES
                Exit For
            End If
        Next i
        Exit Sub
    End If
    
        ' :::::::::::::::::::::::
    ' :: Script Spell anim packet ::
    ' :::::::::::::::::::::::
    If (casestring = "scriptspellanim") Then
        Spell(val(parse$(1))).SpellAnim = val#(parse$(2))
        Spell(val(parse$(1))).SpellTime = val#(parse$(3))
        Spell(val(parse$(1))).SpellDone = val#(parse$(4))
        Spell(val(parse$(1))).Big = val#(parse$(7))
        
        
        For i = 1 To MAX_SCRIPTSPELLS
            If ScriptSpell(i).CastedSpell = NO Then
                ScriptSpell(i).SpellNum = val#(parse$(1))
                ScriptSpell(i).SpellDone = 0
                ScriptSpell(i).SpellVar = 0
                ScriptSpell(i).SpellTime = GetTickCount
                ScriptSpell(i).x = val#(parse$(5))
                ScriptSpell(i).y = val#(parse$(6))
                ScriptSpell(i).CastedSpell = YES
                Exit For
            End If
        Next i
        Exit Sub
    End If
    
    If (casestring = "checkemoticons") Then
        n = val#(parse$(1))
        
        Player(n).EmoticonNum = val#(parse$(2))
        Player(n).EmoticonTime = GetTickCount
        Player(n).EmoticonVar = 0
        Exit Sub
    End If
    
    
    If casestring = "levelup" Then
        Player(val(parse$(1))).LevelUpT = GetTickCount
        Player(val(parse$(1))).LevelUp = 1
        Exit Sub
    End If
    
    If casestring = "damagedisplay" Then
        For i = 1 To MAX_BLT_LINE
            If val#(parse$(1)) = 0 Then
                If BattlePMsg(i).Index <= 0 Then
                    BattlePMsg(i).Index = 1
                    BattlePMsg(i).Msg = parse$(2)
                    BattlePMsg(i).color = val#(parse$(3))
                    BattlePMsg(i).Time = GetTickCount
                    BattlePMsg(i).Done = 1
                    BattlePMsg(i).y = 0
                    Exit Sub
                Else
                    BattlePMsg(i).y = BattlePMsg(i).y - 15
                End If
            Else
                If BattleMMsg(i).Index <= 0 Then
                    BattleMMsg(i).Index = 1
                    BattleMMsg(i).Msg = parse$(2)
                    BattleMMsg(i).color = val#(parse$(3))
                    BattleMMsg(i).Time = GetTickCount
                    BattleMMsg(i).Done = 1
                    BattleMMsg(i).y = 0
                    Exit Sub
                Else
                    BattleMMsg(i).y = BattleMMsg(i).y - 15
                End If
            End If
        Next i
        
        z = 1
        If val#(parse$(1)) = 0 Then
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then
                    If BattlePMsg(i).y < BattlePMsg(i + 1).y Then z = i
                Else
                    If BattlePMsg(i).y < BattlePMsg(1).y Then z = i
                End If
            Next i
                        
            BattlePMsg(z).Index = 1
            BattlePMsg(z).Msg = parse$(2)
            BattlePMsg(z).color = val#(parse$(3))
            BattlePMsg(z).Time = GetTickCount
            BattlePMsg(z).Done = 1
            BattlePMsg(z).y = 0
        Else
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then
                    If BattleMMsg(i).y < BattleMMsg(i + 1).y Then z = i
                Else
                    If BattleMMsg(i).y < BattleMMsg(1).y Then z = i
                End If
            Next i
                        
            BattleMMsg(z).Index = 1
            BattleMMsg(z).Msg = parse$(2)
            BattleMMsg(z).color = val#(parse$(3))
            BattleMMsg(z).Time = GetTickCount
            BattleMMsg(z).Done = 1
            BattleMMsg(z).y = 0
        End If
        Exit Sub
    End If
    
    If casestring = "itembreak" Then
        ItemDur(val(parse$(1))).item = val#(parse$(2))
        ItemDur(val(parse$(1))).Dur = val#(parse$(3))
        ItemDur(val(parse$(1))).Done = 1
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::::::
    ' :: Index player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::::::::
    If casestring = "itemworn" Then
        Player(val(parse$(1))).Armor = val#(parse$(2))
        Player(val(parse$(1))).Weapon = val#(parse$(3))
        Player(val(parse$(1))).Helmet = val#(parse$(4))
        Player(val(parse$(1))).Shield = val#(parse$(5))
        Player(val(parse$(1))).legs = val#(parse$(6))
        Player(val(parse$(1))).Ring = val#(parse$(7))
        Player(val(parse$(1))).Necklace = val#(parse$(8))
        Exit Sub
    End If
    
    If casestring = "scripttile" Then
        frmScript.lblScript.Caption = parse$(1)
        Exit Sub
    End If
    
    If (casestring = "forceclosehouse") Then
        Call HouseEditorCancel
    End If
    
    ' ::::::::::::::::::::::
    ' :: Set player speed ::
    ' ::::::::::::::::::::::
    If casestring = "setspeed" Then
        SetSpeed parse(1), val#(parse(2))
        Exit Sub
    End If
    
    '::::::::::::::::::
    ':: Custom Menu  ::
    '::::::::::::::::::
    If (casestring = "showcustommenu") Then
        'Error handling
        If Not FileExist(parse$(2)) Then
            Call MsgBox(parse(2) & " not found. Menu loading aborted. Please contact a GM to fix this problem.", vbExclamation)
            Exit Sub
        End If
        
        CUSTOM_TITLE = parse$(1)
        CUSTOM_IS_CLOSABLE = val(parse(3))
        
        frmCustom1.picBackground.Top = 0
        frmCustom1.picBackground.Left = 0
        frmCustom1.picBackground = LoadPicture(App.Path & parse$(2))
        frmCustom1.Height = PixelsToTwips(24 + frmCustom1.picBackground.Height, 1)
        frmCustom1.Width = PixelsToTwips(6 + frmCustom1.picBackground.Width, 0)
        frmCustom1.Visible = True
        
        Exit Sub
    End If
    
    If (casestring = "closecustommenu") Then
        
        CUSTOM_TITLE = "CLOSED"
        frmCustom1.Visible = False
        
        Exit Sub
    End If
    
    If (casestring = "loadpiccustommenu") Then
    
        CustomIndex = parse$(1)
        strfilename = parse$(2)
        CustomX = val(parse$(3))
        CustomY = val(parse$(4))
        
        If strfilename = vbNullString Then
            strfilename = "MEGAUBERBLANKNESSOFUNHOLYPOWER" 'smooth :\    -Pickle
        End If
        
        If FileExist(strfilename) = True Then
        frmCustom1.picCustom(CustomIndex) = LoadPicture(App.Path & strfilename)
        frmCustom1.picCustom(CustomIndex).Top = CustomY
        frmCustom1.picCustom(CustomIndex).Left = CustomX
        frmCustom1.picCustom(CustomIndex).Visible = True
        Else
        frmCustom1.picCustom(CustomIndex).Picture = LoadPicture()
        frmCustom1.picCustom(CustomIndex).Visible = False
        End If
        
        Exit Sub
    End If
    
    If (casestring = "loadlabelcustommenu") Then
    
        CustomIndex = parse$(1)
        strfilename = parse$(2)
        CustomX = val(parse$(3))
        CustomY = val(parse$(4))
        customsize = val(parse$(5))
        customcolour = val(parse$(6))
        
        frmCustom1.BtnCustom(CustomIndex).Caption = strfilename
        frmCustom1.BtnCustom(CustomIndex).Top = CustomY
        frmCustom1.BtnCustom(CustomIndex).Left = CustomX
        frmCustom1.BtnCustom(CustomIndex).Font.Bold = True
        frmCustom1.BtnCustom(CustomIndex).Font.Size = customsize
        frmCustom1.BtnCustom(CustomIndex).ForeColor = QBColor(customcolour)
        frmCustom1.BtnCustom(CustomIndex).Visible = True
        frmCustom1.BtnCustom(CustomIndex).Alignment = parse(7)
        
        If parse(8) <= 0 Or parse(9) <= 0 Then
        frmCustom1.BtnCustom(CustomIndex).AutoSize = True
        Else
        frmCustom1.BtnCustom(CustomIndex).AutoSize = False
        frmCustom1.BtnCustom(CustomIndex).Width = parse(8)
        frmCustom1.BtnCustom(CustomIndex).Height = parse(9)
        End If
        
        Exit Sub
    End If
    
    If (casestring = "loadtextboxcustommenu") Then
    
        CustomIndex = parse$(1)
        strfilename = parse$(2)
        CustomX = val(parse$(3))
        CustomY = val(parse$(4))
        customtext = parse$(5)
        
        frmCustom1.txtCustom(CustomIndex).Text = customtext
        frmCustom1.txtCustom(CustomIndex).Top = CustomY
        frmCustom1.txtCustom(CustomIndex).Left = strfilename
        frmCustom1.txtCustom(CustomIndex).Width = CustomX - 32
        frmCustom1.txtcustomOK(CustomIndex).Top = CustomY
        frmCustom1.txtcustomOK(CustomIndex).Left = frmCustom1.txtCustom(CustomIndex).Left + frmCustom1.txtCustom(CustomIndex).Width
        frmCustom1.txtcustomOK(CustomIndex).Visible = True
        frmCustom1.txtCustom(CustomIndex).Visible = True
        
        Exit Sub
    End If
    
    If (casestring = "loadinternetwindow") Then
        customtext = parse$(1)
        'DEBUG STRING
        'Call AddText(customtext, 15)
        ShellExecute 1, "open", Trim(customtext), vbNullString, vbNullString, 1
        Exit Sub
    End If
    
    If (casestring = "returncustomboxmsg") Then
        customsize = parse$(1)
        
        packet = "returningcustomboxmsg" & SEP_CHAR & frmCustom1.txtCustom(customsize).Text & SEP_CHAR & END_CHAR
        Call SendData(packet)
        
        Exit Sub
    End If
    
End Sub

Function ConnectToServer() As Boolean
Dim Wait As Long
    'Let's say that we're not connected, by default
    ConnectToServer = False

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
    If GetPlayerName(Index) <> vbNullString Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Sub SendData(ByVal Data As String)
    If IsConnected Then
        frmMirage.Socket.SendData Data
    End If
End Sub

Sub SendNewAccount(ByVal Name As String, ByVal Password As String, ByVal Email As String)
Dim packet As String

    packet = "newfaccountied" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & Trim$(Email) & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
Dim packet As String
    
    packet = "delimaccounted" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendLogin(ByVal Name As String, ByVal Password As String)
Dim packet As String

    packet = "logination" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & SEC_CODE1 & SEP_CHAR & SEC_CODE2 & SEP_CHAR & SEC_CODE3 & SEP_CHAR & SEC_CODE4 & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal slot As Long, ByVal headc As Long, ByVal bodyc As Long, ByVal legc As Long)
Dim packet As String

    packet = "addachara" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & slot & SEP_CHAR & headc & SEP_CHAR & bodyc & SEP_CHAR & legc & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendDelChar(ByVal slot As Long)
Dim packet As String
    
    packet = "delimbocharu" & SEP_CHAR & slot & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendGetClasses()
Dim packet As String

    packet = "gatglasses" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendUseChar(ByVal CharSlot As Long)
Dim packet As String

    packet = "usagakarim" & SEP_CHAR & CharSlot & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SayMsg(ByVal Text As String)
Dim packet As String

    packet = "saymsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub GlobalMsg(ByVal Text As String)
Dim packet As String

    packet = "globalmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub BroadcastMsg(ByVal Text As String)
Dim packet As String

    packet = "broadcastmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub EmoteMsg(ByVal Text As String)
Dim packet As String

    packet = "emotemsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub MapMsg(ByVal Text As String)
Dim packet As String

    packet = "mapmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim packet As String

    packet = "playermsg" & SEP_CHAR & MsgTo & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub AdminMsg(ByVal Text As String)
Dim packet As String

    packet = "adminmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendPlayerMove()
Dim packet As String
        
packet = "playermove" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & SEP_CHAR & END_CHAR
Call SendData(packet)

End Sub

Sub SendPlayerDir()
Dim packet As String

    packet = "playerdir" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendPlayerRequestNewMap(ByVal Cancel As Long)
Dim packet As String
    
    packet = "requestnewmap" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Cancel & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendMap()
Dim packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long
Dim parse() As String

    packet = "MAPDATA" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim$(Map(GetPlayerMap(MyIndex)).Name) & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Revision & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Moral & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Up & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Down & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Left & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Right & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Music & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootMap & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootX & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootY & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Indoors & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Weather & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map(GetPlayerMap(MyIndex)).Tile(x, y)
               packet = packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .light & SEP_CHAR
                packet = packet & .GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR
            End With
        Next x
    Next y
    
    For x = 1 To MAX_MAP_NPCS
        packet = packet & Map(GetPlayerMap(MyIndex)).Npc(x) & SEP_CHAR
    Next x
    
    packet = packet & Map(GetPlayerMap(MyIndex)).Owner & SEP_CHAR & END_CHAR
    
    x = Int(Len(packet) / 2)
    P1 = Mid$(packet, 1, x)
    P2 = Mid$(packet, x + 1, Len(packet) - x)
    parse = Split(packet, SEP_CHAR)
    Call SendData(packet)
End Sub

Sub WarpMeTo(ByVal Name As String)
Dim packet As String

    packet = "WARPMETO" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub WarpToMe(ByVal Name As String)
Dim packet As String

    packet = "WARPTOME" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub WarpTo(ByVal MapNum As Long)
Dim packet As String
    
    packet = "WARPTO" & SEP_CHAR & MapNum & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
Dim packet As String

    packet = "SETACCESS" & SEP_CHAR & Name & SEP_CHAR & Access & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendSetSprite(ByVal SpriteNum As Long)
Dim packet As String

    packet = "SETSPRITE" & SEP_CHAR & SpriteNum & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendKick(ByVal Name As String)
Dim packet As String

    packet = "KICKPLAYER" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendBan(ByVal Name As String)
Dim packet As String

    packet = "BANPLAYER" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendBanList()
Dim packet As String

    packet = "BANLIST" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendRequestEditItem()
Dim packet As String

    packet = "REQUESTEDITITEM" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendSaveItem(ByVal ItemNum As Long)
Dim packet As String

    packet = "SAVEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(item(ItemNum).Name) & SEP_CHAR & item(ItemNum).Pic & SEP_CHAR & item(ItemNum).Type & SEP_CHAR & item(ItemNum).Data1 & SEP_CHAR & item(ItemNum).Data2 & SEP_CHAR & item(ItemNum).Data3 & SEP_CHAR & item(ItemNum).StrReq & SEP_CHAR & item(ItemNum).DefReq & SEP_CHAR & item(ItemNum).SpeedReq & SEP_CHAR & item(ItemNum).ClassReq & SEP_CHAR & item(ItemNum).AccessReq & SEP_CHAR
    packet = packet & item(ItemNum).AddHP & SEP_CHAR & item(ItemNum).AddMP & SEP_CHAR & item(ItemNum).AddSP & SEP_CHAR & item(ItemNum).AddStr & SEP_CHAR & item(ItemNum).AddDef & SEP_CHAR & item(ItemNum).AddMagi & SEP_CHAR & item(ItemNum).AddSpeed & SEP_CHAR & item(ItemNum).AddEXP & SEP_CHAR & item(ItemNum).desc & SEP_CHAR & item(ItemNum).AttackSpeed & SEP_CHAR & item(ItemNum).Price & SEP_CHAR & item(ItemNum).Stackable & SEP_CHAR & item(ItemNum).Bound
    packet = packet & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
                
Sub SendRequestEditEmoticon()
Dim packet As String

    packet = "REQUESTEDITEMOTICON" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub SendRequestEditElement()
Dim packet As String

    packet = "REQUESTEDITELEMENT" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub


Sub SendRequestEditQuest()
Dim packet As String

    packet = "REQUESTEDITQUEST" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendSaveQuest(ByVal QuestNum As Long)
Dim packet As String
Dim j As Long

    packet = "SAVEQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim$(Quest(QuestNum).Name) & SEP_CHAR & val(Quest(QuestNum).Pictop) & SEP_CHAR & val(Quest(QuestNum).Picleft)
    
    For j = 0 To MAX_QUEST_LENGHT
        packet = packet & SEP_CHAR & Quest(QuestNum).Map(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).x(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).y(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).Npc(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).Script(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake1num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake2num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake1val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemTake2val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive1num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive2num(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive1val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ItemGive2val(j)
        packet = packet & SEP_CHAR & Quest(QuestNum).ExpGiven(j)
    Next j
    
    packet = packet & SEP_CHAR & END_CHAR
    
    Call SendData(packet)
End Sub

Sub SendRequestEditSkill()
Dim packet As String

    packet = "REQUESTEDITSKILL" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendSaveSkill(ByVal SkillNum As Long)
Dim packet As String
Dim j As Long

    packet = "SAVESKILL" & SEP_CHAR & SkillNum & SEP_CHAR & Trim$(skill(SkillNum).Name) & SEP_CHAR & Trim$(skill(SkillNum).Action) & SEP_CHAR & Trim$(skill(SkillNum).Fail) & SEP_CHAR & Trim$(skill(SkillNum).Succes) & SEP_CHAR & Trim$(skill(SkillNum).AttemptName) & SEP_CHAR & val(skill(SkillNum).Pictop) & SEP_CHAR & val(skill(SkillNum).Picleft)
    
    For j = 1 To MAX_SKILLS_SHEETS
        packet = packet & SEP_CHAR & skill(SkillNum).ItemTake1num(j)
        packet = packet & SEP_CHAR & skill(SkillNum).ItemTake2num(j)
        packet = packet & SEP_CHAR & skill(SkillNum).ItemGive1num(j)
        packet = packet & SEP_CHAR & skill(SkillNum).ItemGive2num(j)
        packet = packet & SEP_CHAR & skill(SkillNum).minlevel(j)
        packet = packet & SEP_CHAR & skill(SkillNum).ExpGiven(j)
        packet = packet & SEP_CHAR & skill(SkillNum).base_chance(j)
        packet = packet & SEP_CHAR & skill(SkillNum).ItemTake1val(j)
        packet = packet & SEP_CHAR & skill(SkillNum).ItemTake2val(j)
        packet = packet & SEP_CHAR & skill(SkillNum).ItemGive1val(j)
        packet = packet & SEP_CHAR & skill(SkillNum).ItemGive2val(j)
        packet = packet & SEP_CHAR & skill(SkillNum).itemequiped(j)
    Next j
    
    packet = packet & SEP_CHAR & END_CHAR
    
    Call SendData(packet)
End Sub

Sub SendSaveEmoticon(ByVal EmoNum As Long)
Dim packet As String

    packet = "SAVEEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub SendSaveElement(ByVal ElementNum As Long)
Dim packet As String

    packet = "SAVEELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim$(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendRequestEditArrow()
Dim packet As String

    packet = "REQUESTEDITARROW" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendSaveArrow(ByVal ArrowNum As Long)
Dim packet As String

    packet = "SAVEARROW" & SEP_CHAR & ArrowNum & SEP_CHAR & Trim$(Arrows(ArrowNum).Name) & SEP_CHAR & Arrows(ArrowNum).Pic & SEP_CHAR & Arrows(ArrowNum).Range & SEP_CHAR & Arrows(ArrowNum).Amount & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
                
Sub SendRequestEditNpc()
Dim packet As String

    packet = "REQUESTEDITNPC" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendSaveNpc(ByVal NpcNum As Long)
Dim packet As String
Dim i As Long
    
    'Call GlobalMsg("At sendsavenpc..." & Npc(NpcNum).Element)
    
    packet = "SAVENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).speed & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR & Npc(NpcNum).Element & SEP_CHAR & Npc(NpcNum).Spritesize & SEP_CHAR
    For i = 1 To MAX_NPC_DROPS
        packet = packet & Npc(NpcNum).ItemNPC(i).chance
        packet = packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum
        packet = packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue & SEP_CHAR
    Next i
    packet = packet & END_CHAR
    Call SendData(packet)
End Sub

Sub SendMapRespawn()
Dim packet As String

    packet = "MAPRESPAWN" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendUseItem(ByVal InvNum As Long)
Dim packet As String

    packet = "USEITEM" & SEP_CHAR & InvNum & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendScript(ByVal num As Long)
Dim packet As String

    packet = "scriptedaction" & SEP_CHAR & num & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendDropItem(ByVal InvNum, ByVal Ammount As Long)
Dim packet As String

    packet = "MAPDROPITEM" & SEP_CHAR & InvNum & SEP_CHAR & Ammount & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendWhosOnline()
Dim packet As String

    packet = "WHOSONLINE" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub SendOnlineList()
Dim packet As String

packet = "ONLINELIST" & SEP_CHAR & END_CHAR
Call SendData(packet)
End Sub
            
Sub SendMOTDChange(ByVal MOTD As String)
Dim packet As String

    packet = "SETMOTD" & SEP_CHAR & MOTD & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendRequestEditShop()
Dim packet As String

    packet = "REQUESTEDITSHOP" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendSaveShop(ByVal shopNum As Long)
Dim packet As String
Dim i As Integer

    packet = "SAVESHOP" & SEP_CHAR & shopNum & SEP_CHAR & Trim$(Shop(shopNum).Name) & SEP_CHAR & Shop(shopNum).FixesItems & SEP_CHAR & Shop(shopNum).BuysItems & SEP_CHAR & Shop(shopNum).ShowInfo & SEP_CHAR & Shop(shopNum).currencyItem & SEP_CHAR
    For i = 1 To MAX_SHOP_ITEMS
        packet = packet & Shop(shopNum).ShopItem(i).ItemNum & SEP_CHAR & Shop(shopNum).ShopItem(i).Amount & SEP_CHAR & Shop(shopNum).ShopItem(i).Price & SEP_CHAR
    Next i
    
    packet = packet & END_CHAR
    
    Call SendData(packet)
    
End Sub

Sub SendRequestEditSpell()
Dim packet As String

    packet = "REQUESTEDITSPELL" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub SendReloadScripts()
Dim packet As String

    packet = "reloadscripts" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub SendSaveSpell(ByVal SpellNum As Long)
Dim packet As String

    packet = "SAVESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Trim$(Spell(SpellNum).Sound) & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Spell(SpellNum).Element & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendRequestEditMap()
Dim packet As String

    packet = "REQUESTEDITMAP" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub SendRequestEditHouse()
Dim packet As String

    packet = "REQUESTEDITHOUSE" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Private Sub optHouse_Click()
    frmHouse.scrlItem.Max = MAX_ITEMS
    frmHouse.Show vbModal
End Sub

Sub SendTradeRequest(ByVal Name As String)
Dim packet As String

    packet = "PPTRADE" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendAcceptTrade()
Dim packet As String

    packet = "ATRADE" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendDeclineTrade()
Dim packet As String

    packet = "DTRADE" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendPartyRequest(ByVal Name As String)
Dim packet As String

    packet = "PARTY" & SEP_CHAR & Name & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendJoinParty()
Dim packet As String

    packet = "JOINPARTY" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendLeaveParty()
Dim packet As String

    packet = "LEAVEPARTY" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendBanDestroy()
Dim packet As String
    
    packet = "BANDESTROY" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendRequestLocation()
Dim packet As String

    packet = "REQUESTLOCATION" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub SendSetPlayerSprite(ByVal Name As String, ByVal SpriteNum As Byte)
Dim packet As String

    packet = "SETPLAYERSPRITE" & SEP_CHAR & Name & SEP_CHAR & SpriteNum & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub SendHotScript1()
Dim packet As String

    packet = "hotscript1" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendHotScript2()
Dim packet As String

    packet = "hotscript2" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub SendHotScript3()
Dim packet As String

    packet = "hotscript3" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendHotScript4()
Dim packet As String

    packet = "hotscript4" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub SendEasterEgg()
Dim packet As String

    packet = "easteregg" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub SendScriptTile(ByVal Text As String)
Dim packet As String

    packet = "SCRIPTTILE" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub editmain()
Dim packet As String

    packet = "editmain" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
Sub savemain(ByVal Text As String)
Dim packet As String

    packet = "savemain" & SEP_CHAR & Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendPlayerMovemouse()
Dim packet As String

    packet = "playermovemouse" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub warp()
Dim packet As String

    packet = "warp" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub mail()
Dim packet As String

    packet = "mail" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
