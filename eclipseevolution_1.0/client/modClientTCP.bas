Attribute VB_Name = "modClientTCP"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public ServerIP As String
Public PlayerBuffer As String
Public InGame As Boolean
Public TradePlayer As Long
Public currentmp3 As musictracker

Sub TcpInit()
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    PlayerBuffer = ""
    
    Dim FileName As String
    FileName = App.Path & "\config.ini"

    frmMirage.Socket.RemoteHost = ReadINI("IPCONFIG", "IP", FileName)
    frmMirage.Socket.RemotePort = Val#(ReadINI("IPCONFIG", "PORT", FileName))
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
Dim Top As String * 3
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
Dim i As Long, n As Long, x As Long, y As Long, p As Long
Dim ShopNum As Long, GiveItem As Long, GiveValue As Long, GetItem As Long, GetValue As Long
Dim z As Long
Dim Stuff As String
Dim Stuff2 As String
Dim Stuff3 As String
Dim ThisIsANumber As Long
Dim strfilename As String
Dim CustomX As Long
Dim CustomY As Long
Dim CustomIndex As Long
Dim customcolour As Long
Dim customsize As Long
Dim customtext As String
Dim casestring As String
Dim packet As String

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
    
    casestring = LCase(parse$(0))
    
    If casestring = "maxinfo" Then
        GAME_NAME = Trim$(parse$(1))
        MAX_PLAYERS = Val#(parse$(2))
        MAX_ITEMS = Val#(parse$(3))
        MAX_NPCS = Val#(parse$(4))
        MAX_SHOPS = Val#(parse$(5))
        MAX_SPELLS = Val#(parse$(6))
        MAX_MAPS = Val#(parse$(7))
        MAX_MAP_ITEMS = Val#(parse$(8))
        MAX_MAPX = Val#(parse$(9))
        MAX_MAPY = Val#(parse$(10))
        MAX_EMOTICONS = Val#(parse$(11))
        MAX_ELEMENTS = Val#(parse$(12))
        PAPERDOLL = Val#(parse$(13))
        SPRITESIZE = Val#(parse$(14))
        MAX_SCRIPTSPELLS = Val#(parse$(15))
        ENCRYPT_PASS = Trim$(parse$(16))
        ENCRYPT_TYPE = Trim$(parse$(17))
        
        Call NewsUpdate(Trim$(parse$(18)))
        
        ReDim Map(1 To MAX_MAPS) As MapRec
        'ReDim MapAttributeNpc(1 To MAX_ATTRIBUTE_NPCS, 0 To MAX_MAPX, 0 To MAX_MAPY) As MapNpcRec
        'ReDim SaveMapAttributeNpc(1 To MAX_ATTRIBUTE_NPCS, 0 To MAX_MAPX, 0 To MAX_MAPY) As MapNpcRec
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
            'ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
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
        Next i
        
        For i = 0 To MAX_EMOTICONS
            Emoticons(i).Pic = 0
            Emoticons(i).Command = ""
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
 
        currentmp3.Song = "None"
 
        Exit Sub
    End If
        
    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    If casestring = "npchp" Then
        n = Val#(parse$(1))
 
        MapNpc(n).HP = Val#(parse$(2))
        MapNpc(n).MaxHp = Val#(parse$(3))
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    If casestring = "attributenpchp" Then
        n = Val#(parse$(1))
 
        MapAttributeNpc(n, Val#(parse$(4)), Val#(parse$(5))).HP = Val#(parse$(2))
        MapAttributeNpc(n, Val#(parse$(4)), Val#(parse$(5))).MaxHp = Val#(parse$(3))
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
        n = Val#(parse$(2))
        
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
            Level = Val#(parse$(n + 2))
            
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
    If casestring = "loginok" Then
        ' Now we can receive game data
        MyIndex = Val#(parse$(1))
        
        frmSendGetData.Visible = True
        frmChars.Visible = False
        
        Call SetStatus("Receiving game data...")
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::
    ' ::     News Recieved packet    ::
    ' :::::::::::::::::::::::::::::::::
    If casestring = "news" Then
        Call WriteINI("DATA", "News", parse$(1), (App.Path & "\News.ini"))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If casestring = "newcharclasses" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val#(parse$(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1

        For i = 0 To Max_Classes
            Class(i).Name = parse$(n)
            
            Class(i).HP = Val#(parse$(n + 1))
            Class(i).MP = Val#(parse$(n + 2))
            Class(i).SP = Val#(parse$(n + 3))
            
            Class(i).STR = Val#(parse$(n + 4))
            Class(i).DEF = Val#(parse$(n + 5))
            Class(i).Speed = Val#(parse$(n + 6))
            Class(i).MAGI = Val#(parse$(n + 7))
            'Class(i).INTEL = val#(parse$(n + 8))
            Class(i).MaleSprite = Val#(parse$(n + 8))
            Class(i).FemaleSprite = Val#(parse$(n + 9))
            Class(i).Locked = Val#(parse$(n + 10))
        
        n = n + 11
        Next i
        
        ' Used for if the player is creating a new character
        frmNewChar.Visible = True
        frmSendGetData.Visible = False

        frmNewChar.cmbClass.Clear
        For i = 0 To Max_Classes
            If Class(i).Locked = 0 Then
                frmNewChar.cmbClass.AddItem Trim$(Class(i).Name)
            End If
        Next i
        
        frmNewChar.cmbClass.ListIndex = 0
        frmNewChar.lblHP.Caption = STR(Class(0).HP)
        frmNewChar.lblMP.Caption = STR(Class(0).MP)
        frmNewChar.lblSP.Caption = STR(Class(0).SP)
    
        frmNewChar.lblSTR.Caption = STR(Class(0).STR)
        frmNewChar.lblDEF.Caption = STR(Class(0).DEF)
        frmNewChar.lblSPEED.Caption = STR(Class(0).Speed)
        frmNewChar.lblMAGI.Caption = STR(Class(0).MAGI)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If casestring = "classesdata" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val#(parse$(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For i = 0 To Max_Classes
            Class(i).Name = parse$(n)
            
            Class(i).HP = Val#(parse$(n + 1))
            Class(i).MP = Val#(parse$(n + 2))
            Class(i).SP = Val#(parse$(n + 3))
            
            Class(i).STR = Val#(parse$(n + 4))
            Class(i).DEF = Val#(parse$(n + 5))
            Class(i).Speed = Val#(parse$(n + 6))
            Class(i).MAGI = Val#(parse$(n + 7))
            
            Class(i).Locked = Val#(parse$(n + 8))
            
            n = n + 9
        Next i
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' ::  Game Clock (Time)  ::
    ' :::::::::::::::::::::::::
    If casestring = "gameclock" Then
        frmMirage.GameClock.Caption = parse$(1)
        frmMirage.Label4.Caption = "It is now:"
    End If
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If casestring = "ingame" Then
        InGame = True
        Call GameInit
        Call GameLoop
        If parse$(1) = END_CHAR Then
            MsgBox ("here")
            End
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    If casestring = "playerinv" Then
        n = 2
        z = Val#(parse$(1))
        
        For i = 1 To MAX_INV
            Call SetPlayerInvItemNum(z, i, Val#(parse$(n)))
            Call SetPlayerInvItemValue(z, i, Val#(parse$(n + 1)))
            Call SetPlayerInvItemDur(z, i, Val#(parse$(n + 2)))
            
            n = n + 3
        Next i
        
        If z = MyIndex Then Call UpdateVisInv
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If casestring = "playerinvupdate" Then
        n = Val#(parse$(1))
        z = Val#(parse$(2))
        
        Call SetPlayerInvItemNum(z, n, Val#(parse$(3)))
        Call SetPlayerInvItemValue(z, n, Val#(parse$(4)))
        Call SetPlayerInvItemDur(z, n, Val#(parse$(5)))
        If z = MyIndex Then Call UpdateVisInv
        Exit Sub
    End If
    ' ::::::::::::::::::::::::
   ' :: Player bank packet ::
   ' ::::::::::::::::::::::::
   If casestring = "playerbank" Then
       n = 1
       For i = 1 To MAX_BANK
           Call SetPlayerBankItemNum(MyIndex, i, Val#(parse$(n)))
           Call SetPlayerBankItemValue(MyIndex, i, Val#(parse$(n + 1)))
           Call SetPlayerBankItemDur(MyIndex, i, Val#(parse$(n + 2)))
           
           n = n + 3
       Next i
       
       If frmBank.Visible = True Then Call UpdateBank
       Exit Sub
   End If
   
   ' :::::::::::::::::::::::::::::::
   ' :: Player bank update packet ::
   ' :::::::::::::::::::::::::::::::
   If casestring = "playerbankupdate" Then
       n = Val#(parse$(1))
       
       Call SetPlayerBankItemNum(MyIndex, n, Val#(parse$(2)))
       Call SetPlayerBankItemValue(MyIndex, n, Val#(parse$(3)))
       Call SetPlayerBankItemDur(MyIndex, n, Val#(parse$(4)))
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
               If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
                   frmBank.lstInventory.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
               Else
                   If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerNecklaceSlot(MyIndex) = i Then
                       frmBank.lstInventory.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                   Else
                       frmBank.lstInventory.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                   End If
               End If
           Else
               frmBank.lstInventory.AddItem i & "> Empty"
           End If
           DoEvents
       Next i
       
       For i = 1 To MAX_BANK
           If GetPlayerBankItemNum(MyIndex, i) > 0 Then
               If Item(GetPlayerBankItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(MyIndex, i)).Stackable = 1 Then
                   frmBank.lstBank.AddItem i & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, i)).Name) & " (" & GetPlayerBankItemValue(MyIndex, i) & ")"
               Else
                   If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerNecklaceSlot(MyIndex) = i Then
                       frmBank.lstBank.AddItem i & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, i)).Name) & " (worn)"
                   Else
                       frmBank.lstBank.AddItem i & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, i)).Name)
                   End If
               End If
           Else
               frmBank.lstBank.AddItem i & "> Empty"
           End If
           DoEvents
       Next i
       frmBank.lstBank.ListIndex = 0
       frmBank.lstInventory.ListIndex = 0
       
       frmBank.Show vbModal
       Exit Sub
   End If
   
   If LCase(parse$(0)) = "bankmsg" Then
       frmBank.lblMsg.Caption = Trim$(parse$(1))
       Exit Sub
   End If
   
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    If casestring = "playerworneq" Then
 
        z = Val#(parse$(1))
        If z <= 0 Then Exit Sub
        Call SetPlayerArmorSlot(z, Val#(parse$(2)))
        Call SetPlayerWeaponSlot(z, Val#(parse$(3)))
        Call SetPlayerHelmetSlot(z, Val#(parse$(4)))
        Call SetPlayerShieldSlot(z, Val#(parse$(5)))
        Call SetPlayerLegsSlot(z, Val#(parse$(6)))
        Call SetPlayerRingSlot(z, Val#(parse$(7)))
        Call SetPlayerNecklaceSlot(z, Val#(parse$(8)))
        
        If z = MyIndex Then
            Call UpdateVisInv
        End If
        
        'Call AddText("index:" & Val(parse$(1)) & " armor:" & Val#(parse$(2)) & " wep:" & Val#(parse$(3)) & " helm:" & Val#(parse$(4)) & " shield:" & Val#(parse$(5)) & " legs:" & Val#(parse$(6)) & " ring:" & Val#(parse$(7)) & " necklace:" & Val#(parse$(8)), AlertColor)
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player points packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "playerpoints" Then
        Player(MyIndex).POINTS = Val#(parse$(1))
        frmMirage.lblPoints.Caption = Val#(parse$(1))
        Exit Sub
    End If
        
    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    If casestring = "playerhp" Then
        Player(MyIndex).MaxHp = Val#(parse$(1))
        Call SetPlayerHP(MyIndex, Val#(parse$(2)))
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
        Player(MyIndex).MaxMP = Val#(parse$(1))
        Call SetPlayerMP(MyIndex, Val#(parse$(2)))
        If GetPlayerMaxMP(MyIndex) > 0 Then
            'frmMirage.shpMP.FillColor = RGB(208, 11, 0)
            frmMirage.shpMP.Width = ((GetPlayerMP(MyIndex)) / (GetPlayerMaxMP(MyIndex))) * 150
            frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
        End If
        Exit Sub
    End If
    
    ' speech bubble parse
    If (casestring = "mapmsg2") Then
        Bubble(Val(parse$(2))).Text = parse$(1)
        Bubble(Val(parse$(2))).Created = GetTickCount()
        Exit Sub
    End If
    
    ' scriptbubble parse
    If (casestring = "scriptbubble") Then
        ScriptBubble(Val(parse$(1))).Text = Trim(parse$(2))
        ScriptBubble(Val(parse$(1))).Map = Val(parse$(3))
        ScriptBubble(Val(parse$(1))).x = Val(parse$(4))
        ScriptBubble(Val(parse$(1))).y = Val(parse$(5))
        ScriptBubble(Val(parse$(1))).Colour = Val(parse$(6))
        ScriptBubble(Val(parse$(1))).Created = GetTickCount()
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    If casestring = "playersp" Then
       ' Player(MyIndex).MaxSP = val#(parse$(1))
        'Call SetPlayerSP(MyIndex, val#(parse$(2)))
        'If GetPlayerMaxSP(MyIndex) > 0 Then
            'frmMirage.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
        'End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Player Stats Packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "playerstatspacket") Then
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
        If GetPlayerLegsSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerRingSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerNecklaceSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddSpeed
        End If
        
        If SubStr > 0 Then
            frmMirage.lblSTR.Caption = Val#(parse$(1)) - SubStr & " (+" & SubStr & ")"
        Else
            frmMirage.lblSTR.Caption = Val#(parse$(1))
        End If
        If SubDef > 0 Then
            frmMirage.lblDEF.Caption = Val#(parse$(2)) - SubDef & " (+" & SubDef & ")"
        Else
            frmMirage.lblDEF.Caption = Val#(parse$(2))
        End If
        If SubMagi > 0 Then
            frmMirage.lblMAGI.Caption = Val#(parse$(4)) - SubMagi & " (+" & SubMagi & ")"
        Else
            frmMirage.lblMAGI.Caption = Val#(parse$(4))
        End If
        If SubSpeed > 0 Then
            frmMirage.lblSPEED.Caption = Val#(parse$(3)) - SubSpeed & " (+" & SubSpeed & ")"
        Else
            frmMirage.lblSPEED.Caption = Val#(parse$(3))
        End If
        frmMirage.lblEXP.Caption = Val#(parse$(6)) & " / " & Val#(parse$(5))
        
        frmMirage.shpTNL.Width = (((Val(parse$(6))) / (Val(parse$(5)))) * 150)
        frmMirage.lblLevel.Caption = Val#(parse$(7))
        
        Exit Sub
    End If
                

    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If casestring = "playerdata" Then
        i = Val#(parse$(1))
        Call SetPlayerName(i, parse$(2))
        Call SetPlayerSprite(i, Val#(parse$(3)))
        Call SetPlayerMap(i, Val#(parse$(4)))
        Call SetPlayerX(i, Val#(parse$(5)))
        Call SetPlayerY(i, Val#(parse$(6)))
        Call SetPlayerDir(i, Val#(parse$(7)))
        Call SetPlayerAccess(i, Val#(parse$(8)))
        Call SetPlayerPK(i, Val#(parse$(9)))
        Call SetPlayerGuild(i, parse$(10))
        Call SetPlayerGuildAccess(i, Val#(parse$(11)))
        Call SetPlayerClass(i, Val#(parse$(12)))

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
    
    
    ' ::::::::::::::::::::::::
    ' :: Update Sprite Packet ::
    ' ::::::::::::::::::::::::
    If casestring = "updatesprite" Then
        i = Val#(parse$(1))
        Call SetPlayerSprite(i, Val#(parse$(1)))
    End If
    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "playermove") Then
        i = Val#(parse$(1))
        x = Val#(parse$(2))
        y = Val#(parse$(3))
        Dir = Val#(parse$(4))
        n = Val#(parse$(5))

        If Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Sub

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
    If (casestring = "npcmove") Then
        i = Val#(parse$(1))
        x = Val#(parse$(2))
        y = Val#(parse$(3))
        Dir = Val#(parse$(4))
        n = Val#(parse$(5))

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
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "attributenpcmove") Then
        i = Val#(parse$(1))
        x = Val#(parse$(2))
        y = Val#(parse$(3))
        Dir = Val#(parse$(4))
        n = Val#(parse$(5))

        MapAttributeNpc(i, Val#(parse$(6)), Val#(parse$(7))).x = x
        MapAttributeNpc(i, Val#(parse$(6)), Val#(parse$(7))).y = y
        MapAttributeNpc(i, Val#(parse$(6)), Val#(parse$(7))).Dir = Dir
        MapAttributeNpc(i, Val#(parse$(6)), Val#(parse$(7))).XOffset = 0
        MapAttributeNpc(i, Val#(parse$(6)), Val#(parse$(7))).YOffset = 0
        MapAttributeNpc(i, Val#(parse$(6)), Val#(parse$(7))).Moving = n
        
        Select Case MapAttributeNpc(i, Val#(parse$(6)), Val#(parse$(7))).Dir
            Case DIR_UP
                MapAttributeNpc(i, Val#(parse$(6)), Val#(parse$(7))).YOffset = PIC_Y
            Case DIR_DOWN
                MapAttributeNpc(i, Val#(parse$(6)), Val#(parse$(7))).YOffset = PIC_Y * -1
            Case DIR_LEFT
                MapAttributeNpc(i, Val#(parse$(6)), Val#(parse$(7))).XOffset = PIC_X
            Case DIR_RIGHT
                MapAttributeNpc(i, Val#(parse$(6)), Val#(parse$(7))).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (casestring = "playerdir") Then
        i = Val#(parse$(1))
        Dir = Val#(parse$(2))
        
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Sub

        Call SetPlayerDir(i, Dir)
        
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "npcdir") Then
        i = Val#(parse$(1))
        Dir = Val#(parse$(2))
        MapNpc(i).Dir = Dir
        
        MapNpc(i).XOffset = 0
        MapNpc(i).YOffset = 0
        MapNpc(i).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "attributenpcdir") Then
        i = Val#(parse$(1))
        Dir = Val#(parse$(2))
        MapAttributeNpc(i, Val#(parse$(3)), Val#(parse$(4))).Dir = Dir
        
        MapAttributeNpc(i, Val#(parse$(3)), Val#(parse$(4))).XOffset = 0
        MapAttributeNpc(i, Val#(parse$(3)), Val#(parse$(4))).YOffset = 0
        MapAttributeNpc(i, Val#(parse$(3)), Val#(parse$(4))).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (casestring = "playerxy") Then
        x = Val#(parse$(1))
        y = Val#(parse$(2))
        
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
    If (casestring = "attack") Then
        i = Val#(parse$(1))
        
        ' Set player to attacking
        Player(i).Attacking = 1
        Player(i).AttackTimer = GetTickCount
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (casestring = "npcattack") Then
        i = Val#(parse$(1))
        
        ' Set player to attacking
        MapNpc(i).Attacking = 1
        MapNpc(i).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (casestring = "attributenpcattack") Then
        i = Val#(parse$(1))
        
        ' Set player to attacking
        MapAttributeNpc(i, Val#(parse$(2)), Val#(parse$(3))).Attacking = 1
        MapAttributeNpc(i, Val#(parse$(2)), Val#(parse$(3))).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "checkformap") Then
        ' Erase all players except self
        For i = 1 To MAX_PLAYERS
            If i <> MyIndex Then
                Call SetPlayerMap(i, 0)
            End If
        Next i
        
        ' Erase all temporary tile values
        Call ClearTempTile
        '!!!

        ' Get map num
        x = Val#(parse$(1))
        
        ' Get revision
        y = Val#(parse$(2))
        
        If FileExist("maps\map" & x & ".dat") Then
            ' Check to see if the revisions match
            If GetMapRevision(x) = y Then
                ' We do so we dont need the map
                
                ' Load the map
                'Call LoadMap(X)
                
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
        
        Map(Val(parse$(1))).Name = parse$(n + 1)
        Map(Val(parse$(1))).Revision = Val#(parse$(n + 2))
        Map(Val(parse$(1))).Moral = Val#(parse$(n + 3))
        Map(Val(parse$(1))).Up = Val#(parse$(n + 4))
        Map(Val(parse$(1))).Down = Val#(parse$(n + 5))
        Map(Val(parse$(1))).Left = Val#(parse$(n + 6))
        Map(Val(parse$(1))).Right = Val#(parse$(n + 7))
        Map(Val(parse$(1))).Music = parse$(n + 8)
        Map(Val(parse$(1))).BootMap = Val#(parse$(n + 9))
        Map(Val(parse$(1))).BootX = Val#(parse$(n + 10))
        Map(Val(parse$(1))).BootY = Val#(parse$(n + 11))
        Map(Val(parse$(1))).Indoors = Val#(parse$(n + 12))
        
        n = n + 13
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(Val(parse$(1))).Tile(x, y).Ground = Val#(parse$(n))
                Map(Val(parse$(1))).Tile(x, y).Mask = Val#(parse$(n + 1))
                Map(Val(parse$(1))).Tile(x, y).Anim = Val#(parse$(n + 2))
                Map(Val(parse$(1))).Tile(x, y).Mask2 = Val#(parse$(n + 3))
                Map(Val(parse$(1))).Tile(x, y).M2Anim = Val#(parse$(n + 4))
                Map(Val(parse$(1))).Tile(x, y).Fringe = Val#(parse$(n + 5))
                Map(Val(parse$(1))).Tile(x, y).FAnim = Val#(parse$(n + 6))
                Map(Val(parse$(1))).Tile(x, y).Fringe2 = Val#(parse$(n + 7))
                Map(Val(parse$(1))).Tile(x, y).F2Anim = Val#(parse$(n + 8))
                Map(Val(parse$(1))).Tile(x, y).Type = Val#(parse$(n + 9))
                Map(Val(parse$(1))).Tile(x, y).Data1 = Val#(parse$(n + 10))
                Map(Val(parse$(1))).Tile(x, y).Data2 = Val#(parse$(n + 11))
                Map(Val(parse$(1))).Tile(x, y).Data3 = Val#(parse$(n + 12))
                Map(Val(parse$(1))).Tile(x, y).String1 = parse$(n + 13)
                Map(Val(parse$(1))).Tile(x, y).String2 = parse$(n + 14)
                Map(Val(parse$(1))).Tile(x, y).String3 = parse$(n + 15)
                Map(Val(parse$(1))).Tile(x, y).Light = Val#(parse$(n + 16))
                Map(Val(parse$(1))).Tile(x, y).GroundSet = Val#(parse$(n + 17))
                Map(Val(parse$(1))).Tile(x, y).MaskSet = Val#(parse$(n + 18))
                Map(Val(parse$(1))).Tile(x, y).AnimSet = Val#(parse$(n + 19))
                Map(Val(parse$(1))).Tile(x, y).Mask2Set = Val#(parse$(n + 20))
                Map(Val(parse$(1))).Tile(x, y).M2AnimSet = Val#(parse$(n + 21))
                Map(Val(parse$(1))).Tile(x, y).FringeSet = Val#(parse$(n + 22))
                Map(Val(parse$(1))).Tile(x, y).FAnimSet = Val#(parse$(n + 23))
                Map(Val(parse$(1))).Tile(x, y).Fringe2Set = Val#(parse$(n + 24))
                Map(Val(parse$(1))).Tile(x, y).F2AnimSet = Val#(parse$(n + 25))
                
                n = n + 26
            Next x
        Next y
        
        For x = 1 To MAX_MAP_NPCS
            Map(Val(parse$(1))).Npc(x) = Val#(parse$(n))
            n = n + 1
        Next x
    
        ' Save the map
        Call SaveLocalMap(Val(parse$(1)))
        
        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmMapEditor.Visible = False
            frmAttributes.Visible = False
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
     x = Val#(parse$(2))
     y = Val#(parse$(3))
     
     Select Case Val#(parse$(4))
     Case 0
                Map(Val(parse$(1))).Tile(x, y).Ground = Val#(parse$(n))
                Map(Val(parse$(1))).Tile(x, y).GroundSet = Val#(parse$(n + 1))
     Case 1
                Map(Val(parse$(1))).Tile(x, y).Mask = Val#(parse$(n))
                Map(Val(parse$(1))).Tile(x, y).MaskSet = Val#(parse$(n + 1))
     Case 2
                Map(Val(parse$(1))).Tile(x, y).Anim = Val#(parse$(n))
                Map(Val(parse$(1))).Tile(x, y).AnimSet = Val#(parse$(n + 1))
     Case 3
                Map(Val(parse$(1))).Tile(x, y).Mask2 = Val#(parse$(n))
                Map(Val(parse$(1))).Tile(x, y).Mask2Set = Val#(parse$(n + 1))
     Case 4
                Map(Val(parse$(1))).Tile(x, y).M2Anim = Val#(parse$(n))
                Map(Val(parse$(1))).Tile(x, y).M2AnimSet = Val#(parse$(n + 1))
     Case 5
                Map(Val(parse$(1))).Tile(x, y).Fringe = Val#(parse$(n))
                Map(Val(parse$(1))).Tile(x, y).FringeSet = Val#(parse$(n + 1))
     Case 6
                Map(Val(parse$(1))).Tile(x, y).FAnim = Val#(parse$(n))
                Map(Val(parse$(1))).Tile(x, y).FAnimSet = Val#(parse$(n + 1))
     Case 7
                Map(Val(parse$(1))).Tile(x, y).Fringe2 = Val#(parse$(n))
                Map(Val(parse$(1))).Tile(x, y).Fringe2Set = Val#(parse$(n + 1))
     Case 8
                Map(Val(parse$(1))).Tile(x, y).F2Anim = Val#(parse$(n))
                Map(Val(parse$(1))).Tile(x, y).F2AnimSet = Val#(parse$(n + 1))
     End Select
        Call SaveLocalMap(Val(parse$(1)))
    End If
    
    If casestring = "tilecheckattribute" Then
     n = 5
     x = Val#(parse$(2))
     y = Val#(parse$(3))
     
                Map(Val(parse$(1))).Tile(x, y).Type = Val#(parse$(n - 1))
                Map(Val(parse$(1))).Tile(x, y).Data1 = Val#(parse$(n))
                Map(Val(parse$(1))).Tile(x, y).Data2 = Val#(parse$(n + 1))
                Map(Val(parse$(1))).Tile(x, y).Data3 = Val#(parse$(n + 2))
                Map(Val(parse$(1))).Tile(x, y).String1 = parse$(n + 3)
                Map(Val(parse$(1))).Tile(x, y).String2 = parse$(n + 4)
                Map(Val(parse$(1))).Tile(x, y).String3 = parse$(n + 5)
        Call SaveLocalMap(Val(parse$(1)))
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If casestring = "mapitemdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).num = Val#(parse$(n))
            SaveMapItem(i).Value = Val#(parse$(n + 1))
            SaveMapItem(i).Dur = Val#(parse$(n + 2))
            SaveMapItem(i).x = Val#(parse$(n + 3))
            SaveMapItem(i).y = Val#(parse$(n + 4))
            
            n = n + 5
        Next i
        
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If casestring = "mapnpcdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_NPCS
            SaveMapNpc(i).num = Val#(parse$(n))
            SaveMapNpc(i).x = Val#(parse$(n + 1))
            SaveMapNpc(i).y = Val#(parse$(n + 2))
            SaveMapNpc(i).Dir = Val#(parse$(n + 3))
            
            n = n + 4
        Next i
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If casestring = "mapattributenpcdata" Then
        n = 3
        
        x = Val#(parse$(1))
        y = Val#(parse$(2))
        
        For i = 1 To MAX_ATTRIBUTE_NPCS
            'If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                'If i <= Map(GetPlayerMap(MyIndex)).Tile(x, y).Data2 Then
                    SaveMapAttributeNpc(i, x, y).num = Val#(parse$(n))
                    SaveMapAttributeNpc(i, x, y).x = Val#(parse$(n + 1))
                    SaveMapAttributeNpc(i, x, y).y = Val#(parse$(n + 2))
                    SaveMapAttributeNpc(i, x, y).Dir = Val#(parse$(n + 3))
    
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
        
        'For y = 0 To MAX_MAPY
        '    For x = 0 To MAX_MAPX
        '        If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
        '            For i = 1 To MAX_ATTRIBUTE_NPCS
        '                If i <= Map(GetPlayerMap(MyIndex)).Tile(x, y).Data2 Then
        '                    MapAttributeNpc(i, x, y) = SaveMapAttributeNpc(i, x, y)
        '                End If
        '            Next i
        '        End If
        '    Next x
        'Next y
        
        GettingMap = False
        
        ' Play music
        If Trim$(Map(GetPlayerMap(MyIndex)).Music) <> "None" Then
                        
            Select Case Right(Trim$(Map(GetPlayerMap(MyIndex)).Music), 4)
            
                Case ".mid"
                If currentmp3.Song <> Trim$(Map(GetPlayerMap(MyIndex)).Music) Then
                
                    If Trim$(Map(GetPlayerMap(MyIndex)).Music) = "None" Then
                        Call StopMidi
                    Else
                        Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
                    End If

                    currentmp3.Song = Trim$(Map(GetPlayerMap(MyIndex)).Music)
                    frmMirage.Mp3musicplayer.currentPlaylist.Clear
                    frmMirage.Mp3musicplayer.Controls.Stop
                    frmMirage.Mp3musicplayer.URL = ""
                    frmMirage.Mp3musicplayer.settings.playCount = 0
                    frmMirage.Mp3timer.Enabled = False
                End If
            
                Case ".mp3"
                If frmMirage.mp3player.playState = 0 Or frmMirage.mp3player.playState = 1 Then
                    If currentmp3.Song <> Trim$(Map(GetPlayerMap(MyIndex)).Music) Then
                        If Trim$(Map(GetPlayerMap(MyIndex)).Music) <> "None" Then
                            Call StopMidi
                            frmMirage.Mp3musicplayer.currentPlaylist.Clear
                            frmMirage.Mp3musicplayer.Controls.Stop
                            frmMirage.Mp3musicplayer.URL = App.Path & "\Music\" & Trim$(Map(GetPlayerMap(MyIndex)).Music) & ""
                            frmMirage.Mp3musicplayer.settings.playCount = 500
                            frmMirage.Mp3timer.Enabled = True
                            currentmp3.Song = Trim$(Map(GetPlayerMap(MyIndex)).Music)
                        Else
                            Call StopMidi
                            frmMirage.Mp3musicplayer.currentPlaylist.Clear
                            frmMirage.Mp3musicplayer.Controls.Stop
                            frmMirage.Mp3musicplayer.URL = ""
                            frmMirage.Mp3musicplayer.settings.playCount = 0
                            frmMirage.Mp3timer.Enabled = False
                            currentmp3.Song = "None"
                        End If
                    End If
                End If
                
                Case ".wma"
                If frmMirage.mp3player.playState = 0 Or frmMirage.mp3player.playState = 1 Then
                    If currentmp3.Song <> Trim$(Map(GetPlayerMap(MyIndex)).Music) Then
                        If Trim$(Map(GetPlayerMap(MyIndex)).Music) <> "None" Then
                            Call StopMidi
                            frmMirage.Mp3musicplayer.currentPlaylist.Clear
                            frmMirage.Mp3musicplayer.Controls.Stop
                            frmMirage.Mp3musicplayer.URL = App.Path & "\Music\" & Trim$(Map(GetPlayerMap(MyIndex)).Music) & ""
                            frmMirage.Mp3musicplayer.settings.playCount = 500
                            frmMirage.Mp3timer.Enabled = True
                            currentmp3.Song = Trim$(Map(GetPlayerMap(MyIndex)).Music)
                        Else
                            Call StopMidi
                            frmMirage.Mp3musicplayer.currentPlaylist.Clear
                            frmMirage.Mp3musicplayer.Controls.Stop
                            frmMirage.Mp3musicplayer.URL = ""
                            frmMirage.Mp3musicplayer.settings.playCount = 0
                            frmMirage.Mp3timer.Enabled = False
                            currentmp3.Song = "None"
                        End If
                    End If
                End If
                
            End Select
        
        Else
        
        'Call StopMidi
        
        'currentmp3.song = "None"
        'frmMirage.Mp3musicplayer.currentPlaylist.Clear
        'frmMirage.Mp3musicplayer.Controls.Stop
        'frmMirage.Mp3musicplayer.URL = ""
        'frmMirage.Mp3musicplayer.settings.playCount = 0
        'frmMirage.Mp3timer.Enabled = False
        
        End If
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (casestring = "saymsg") Or (casestring = "broadcastmsg") Or (casestring = "globalmsg") Or (casestring = "playermsg") Or (casestring = "mapmsg") Or (casestring = "adminmsg") Then
        Call AddText(parse$(1), Val#(parse$(2)))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If casestring = "spawnitem" Then
        n = Val#(parse$(1))
        
        MapItem(n).num = Val#(parse$(2))
        MapItem(n).Value = Val#(parse$(3))
        MapItem(n).Dur = Val#(parse$(4))
        MapItem(n).x = Val#(parse$(5))
        MapItem(n).y = Val#(parse$(6))
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
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updateitem") Then
        n = Val#(parse$(1))
        
        ' Update the item
        Item(n).Name = parse$(2)
        Item(n).Pic = Val#(parse$(3))
        Item(n).Type = Val#(parse$(4))
        Item(n).Data1 = Val#(parse$(5))
        Item(n).Data2 = Val#(parse$(6))
        Item(n).Data3 = Val#(parse$(7))
        Item(n).StrReq = Val#(parse$(8))
        Item(n).DefReq = Val#(parse$(9))
        Item(n).SpeedReq = Val#(parse$(10))
        Item(n).ClassReq = Val#(parse$(11))
        Item(n).AccessReq = Val#(parse$(12))
        
        Item(n).AddHP = Val#(parse$(13))
        Item(n).AddMP = Val#(parse$(14))
        Item(n).AddSP = Val#(parse$(15))
        Item(n).AddStr = Val#(parse$(16))
        Item(n).AddDef = Val#(parse$(17))
        Item(n).AddMagi = Val#(parse$(18))
        Item(n).AddSpeed = Val#(parse$(19))
        Item(n).AddEXP = Val#(parse$(20))
        Item(n).desc = parse$(21)
        Item(n).AttackSpeed = Val#(parse$(22))
        Item(n).Price = Val#(parse$(23))
        Item(n).Stackable = Val#(parse$(24))
        Item(n).Bound = Val#(parse$(25))
        Exit Sub
    End If
       
    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (casestring = "edititem") Then
        n = Val#(parse$(1))
        
        ' Update the item
        Item(n).Name = parse$(2)
        Item(n).Pic = Val#(parse$(3))
        Item(n).Type = Val#(parse$(4))
        Item(n).Data1 = Val#(parse$(5))
        Item(n).Data2 = Val#(parse$(6))
        Item(n).Data3 = Val#(parse$(7))
        Item(n).StrReq = Val#(parse$(8))
        Item(n).DefReq = Val#(parse$(9))
        Item(n).SpeedReq = Val#(parse$(10))
        Item(n).ClassReq = Val#(parse$(11))
        Item(n).AccessReq = Val#(parse$(12))
        
        Item(n).AddHP = Val#(parse$(13))
        Item(n).AddMP = Val#(parse$(14))
        Item(n).AddSP = Val#(parse$(15))
        Item(n).AddStr = Val#(parse$(16))
        Item(n).AddDef = Val#(parse$(17))
        Item(n).AddMagi = Val#(parse$(18))
        Item(n).AddSpeed = Val#(parse$(19))
        Item(n).AddEXP = Val#(parse$(20))
        Item(n).desc = parse$(21)
        Item(n).AttackSpeed = Val#(parse$(22))
        Item(n).Price = Val#(parse$(23))
        Item(n).Stackable = Val#(parse$(24))
        Item(n).Bound = Val#(parse$(25))
        
        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If casestring = "spawnnpc" Then
        n = Val#(parse$(1))
        
        MapNpc(n).num = Val#(parse$(2))
        MapNpc(n).x = Val#(parse$(3))
        MapNpc(n).y = Val#(parse$(4))
        MapNpc(n).Dir = Val#(parse$(5))
        MapNpc(n).Big = Val#(parse$(6))
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If casestring = "spawnattributenpc" Then
        n = Val#(parse$(1))
        
        x = Val#(parse$(7))
        y = Val#(parse$(8))
        
        MapAttributeNpc(n, x, y).num = Val#(parse$(2))
        MapAttributeNpc(n, x, y).x = Val#(parse$(3))
        MapAttributeNpc(n, x, y).y = Val#(parse$(4))
        MapAttributeNpc(n, x, y).Dir = Val#(parse$(5))
        MapAttributeNpc(n, x, y).Big = Val#(parse$(6))
        
        ' Client use only
        MapAttributeNpc(n, x, y).XOffset = 0
        MapAttributeNpc(n, x, y).YOffset = 0
        MapAttributeNpc(n, x, y).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If casestring = "npcdead" Then
        n = Val#(parse$(1))
        
        MapNpc(n).num = 0
        MapNpc(n).x = 0
        MapNpc(n).y = 0
        MapNpc(n).Dir = 0
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If casestring = "attributenpcdead" Then
        n = Val#(parse$(1))
        
        MapAttributeNpc(n, Val#(parse$(2)), Val#(parse$(3))).num = 0
        MapAttributeNpc(n, Val#(parse$(2)), Val#(parse$(3))).x = 0
        MapAttributeNpc(n, Val#(parse$(2)), Val#(parse$(3))).y = 0
        MapAttributeNpc(n, Val#(parse$(2)), Val#(parse$(3))).Dir = 0
        
        ' Client use only
        MapAttributeNpc(n, Val#(parse$(2)), Val#(parse$(3))).XOffset = 0
        MapAttributeNpc(n, Val#(parse$(2)), Val#(parse$(3))).YOffset = 0
        MapAttributeNpc(n, Val#(parse$(2)), Val#(parse$(3))).Moving = 0
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
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (casestring = "updatenpc") Then
        n = Val#(parse$(1))
        
        ' Update the item
        Npc(n).Name = parse$(2)
        Npc(n).AttackSay = ""
        Npc(n).Sprite = Val#(parse$(3))
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
        Npc(n).Big = Val#(parse$(4))
        Npc(n).MaxHp = Val#(parse$(5))
        Npc(n).EXP = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet :: <- Used for item editor admins only
    ' :::::::::::::::::::::
    If (casestring = "editnpc") Then
        n = Val#(parse$(1))
        
        ' Update the npc
        Npc(n).Name = parse$(2)
        Npc(n).AttackSay = parse$(3)
        Npc(n).Sprite = Val#(parse$(4))
        Npc(n).SpawnSecs = Val#(parse$(5))
        Npc(n).Behavior = Val#(parse$(6))
        Npc(n).Range = Val#(parse$(7))
        Npc(n).STR = Val#(parse$(8))
        Npc(n).DEF = Val#(parse$(9))
        Npc(n).Speed = Val#(parse$(10))
        Npc(n).MAGI = Val#(parse$(11))
        Npc(n).Big = Val#(parse$(12))
        Npc(n).MaxHp = Val#(parse$(13))
        Npc(n).EXP = Val#(parse$(14))
        Npc(n).SpawnTime = Val#(parse$(15))
        Npc(n).Element = Val#(parse$(16))
        
       ' Call GlobalMsg("At editnpc..." & Npc(n).Element)
        z = 17
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = Val#(parse$(z))
            Npc(n).ItemNPC(i).ItemNum = Val#(parse$(z + 1))
            Npc(n).ItemNPC(i).ItemValue = Val#(parse$(z + 2))
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
        x = Val#(parse$(1))
        y = Val#(parse$(2))
        n = Val#(parse$(3))
                
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
    
    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "shopeditor") Then
        InShopEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SHOPS
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updateshop") Then
        n = Val#(parse$(1))
        
        ' Update the shop name
        Shop(n).Name = parse$(2)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    If (casestring = "editshop") Then
        ShopNum = Val#(parse$(1))
        
        ' Update the shop
        Shop(ShopNum).Name = parse$(2)
        Shop(ShopNum).JoinSay = parse$(3)
        Shop(ShopNum).LeaveSay = parse$(4)
        Shop(ShopNum).FixesItems = Val#(parse$(5))
        
        n = 6
        For z = 1 To 7
            For i = 1 To MAX_TRADES
                
                GiveItem = Val#(parse$(n))
                GiveValue = Val#(parse$(n + 1))
                GetItem = Val#(parse$(n + 2))
                GetValue = Val#(parse$(n + 3))
                
                Shop(ShopNum).TradeItem(z).Value(i).GiveItem = GiveItem
                Shop(ShopNum).TradeItem(z).Value(i).GiveValue = GiveValue
                Shop(ShopNum).TradeItem(z).Value(i).GetItem = GetItem
                Shop(ShopNum).TradeItem(z).Value(i).GetValue = GetValue
                
                n = n + 4
            Next i
        Next z
        
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
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updatespell") Then
        n = Val#(parse$(1))
        
        ' Update the spell name
        Spell(n).Name = parse$(2)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (casestring = "editspell") Then
        n = Val#(parse$(1))
        
        ' Update the spell
        Spell(n).Name = parse$(2)
        Spell(n).ClassReq = Val#(parse$(3))
        Spell(n).LevelReq = Val#(parse$(4))
        Spell(n).Type = Val#(parse$(5))
        Spell(n).Data1 = Val#(parse$(6))
        Spell(n).Data2 = Val#(parse$(7))
        Spell(n).Data3 = Val#(parse$(8))
        Spell(n).MPCost = Val#(parse$(9))
        Spell(n).Sound = Val#(parse$(10))
        Spell(n).Range = Val#(parse$(11))
        Spell(n).SpellAnim = Val#(parse$(12))
        Spell(n).SpellTime = Val#(parse$(13))
        Spell(n).SpellDone = Val#(parse$(14))
        Spell(n).AE = Val#(parse$(15))
        Spell(n).Big = Val#(parse$(16))
        Spell(n).Element = Val#(parse$(17))
                        
                        
        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (casestring = "trade") Then
        ShopNum = Val#(parse$(1))
        If Val#(parse$(2)) = 1 Then
            frmTrade.picFixItems.Visible = True
        Else
            frmTrade.picFixItems.Visible = False
        End If
        
        n = 3
        For z = 1 To 7
            For i = 1 To MAX_TRADES
                GiveItem = Val#(parse$(n))
                GiveValue = Val#(parse$(n + 1))
                GetItem = Val#(parse$(n + 2))
                GetValue = Val#(parse$(n + 3))
                
                Trade(z).Items(i).ItemGetNum = GetItem
                Trade(z).Items(i).ItemGiveNum = GiveItem
                Trade(z).Items(i).ItemGetVal = GetValue
                Trade(z).Items(i).ItemGiveVal = GiveValue
                
                n = n + 4
            Next i
        Next z
        
        Dim xx As Long
        For xx = 1 To 7
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

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (casestring = "spells") Then
        
        frmMirage.picPlayerSpells.Visible = True
        frmMirage.lstSpells.Clear
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(i) = Val#(parse$(i))
            If Player(MyIndex).Spell(i) <> 0 Then
                frmMirage.lstSpells.AddItem i & ": " & Trim$(Spell(Player(MyIndex).Spell(i)).Name)
            Else
                frmMirage.lstSpells.AddItem "--- Slot Free ---"
            End If
        Next i
        
        frmMirage.lstSpells.ListIndex = 0
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (casestring = "weather") Then
        If Val#(parse$(1)) = WEATHER_RAINING And GameWeather <> WEATHER_RAINING Then
            Call AddText("You see drops of rain falling from the sky above!", BrightGreen)
        End If
        If Val#(parse$(1)) = WEATHER_THUNDER And GameWeather <> WEATHER_THUNDER Then
            Call AddText("You see thunder in the sky above!", BrightGreen)
        End If
        If Val#(parse$(1)) = WEATHER_SNOWING And GameWeather <> WEATHER_SNOWING Then
            Call AddText("You see snow falling from the sky above!", BrightGreen)
        End If
        
        If Val#(parse$(1)) = WEATHER_NONE Then
            If GameWeather = WEATHER_RAINING Then
                Call AddText("The rain beings to calm.", BrightGreen)
            ElseIf GameWeather = WEATHER_SNOWING Then
                Call AddText("The snow is melting away.", BrightGreen)
            ElseIf GameWeather = WEATHER_THUNDER Then
                Call AddText("The thunder begins to disapear.", BrightGreen)
            End If
        End If
        GameWeather = Val#(parse$(1))
        RainIntensity = Val#(parse$(2))
        If MAX_RAINDROPS <> RainIntensity Then
            MAX_RAINDROPS = RainIntensity
            ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
            ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
        End If
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Get Online List ::
    ' ::::::::::::::::::::::::::
    If casestring = "onlinelist" Then
    frmMirage.lstOnline.Clear
    
        n = 2
        z = Val#(parse$(1))
        For x = n To (z + 1)
            frmMirage.lstOnline.AddItem Trim$(parse$(n))
            n = n + 2
        Next x
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
    If casestring = "blitplayerdmg" Then
        DmgDamage = Val#(parse$(1))
        NPCWho = Val#(parse$(2))
        DmgTime = GetTickCount
        iii = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
    If casestring = "blitnpcdmg" Then
        NPCDmgDamage = Val#(parse$(1))
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
            Trading(i).InvName = ""
            Trading2(i).InvNum = 0
            Trading2(i).InvName = ""
            frmPlayerTrade.Items1.AddItem i & ": <Nothing>"
            frmPlayerTrade.Items2.AddItem i & ": <Nothing>"
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
            Trading(i).InvName = ""
            Trading2(i).InvNum = 0
            Trading2(i).InvName = ""
        Next i
        
        frmPlayerTrade.Command1.ForeColor = &H0&
        frmPlayerTrade.Command2.ForeColor = &H0&
        
        frmPlayerTrade.Visible = False
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Disable Time ::
    ' ::::::::::::::::::
    If casestring = "dtime" Then
        If Val#(parse$(1)) = 1 Then
            frmMirage.Label4.Caption = ""
            frmMirage.GameClock.Caption = ""
            frmMirage.Label4.Visible = False
            frmMirage.GameClock.Visible = False
        Else
            frmMirage.Label4.Visible = True
            frmMirage.GameClock.Visible = True
            frmMirage.Label4.Caption = ""
            frmMirage.GameClock.Caption = ""
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "updatetradeitem" Then
            n = Val#(parse$(1))
            
            Trading2(n).InvNum = Val#(parse$(2))
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
        n = Val#(parse$(1))
            If n = 0 Then frmPlayerTrade.Command2.ForeColor = &H0&
            If n = 1 Then frmPlayerTrade.Command2.ForeColor = &HFF00&
        Exit Sub
    End If
    
' :::::::::::::::::::::::::
' :: Chat System Packets ::
' :::::::::::::::::::::::::
    If casestring = "ppchatting" Then
        frmPlayerChat.txtChat.Text = ""
        frmPlayerChat.txtSay.Text = ""
        frmPlayerChat.Label1.Caption = "Chatting With: " & Trim$(Player(Val(parse$(1))).Name)

        frmPlayerChat.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    If casestring = "qchat" Then
        frmPlayerChat.txtChat.Text = ""
        frmPlayerChat.txtSay.Text = ""
        frmPlayerChat.Visible = False
        frmPlayerTrade.Command2.ForeColor = &H8000000F
        frmPlayerTrade.Command1.ForeColor = &H8000000F
        Exit Sub
    End If
    
    If casestring = "sendchat" Then
        Dim s As String
  
        s = vbNewLine & GetPlayerName(Val(parse$(2))) & "> " & Trim$(parse$(1))
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
        s = LCase(parse$(1))
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
                Call PlaySound("magic" & Val#(parse$(2)) & ".wav")
            Case "warp"
                Call PlaySound("warp.wav")
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
        If Val#(parse$(1)) = 1 Then
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
        If Val#(parse$(1)) = 1 Then
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
        Player(Val(parse$(2))).Dir = Val#(parse$(1))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Flash Movie Event Packet ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "flashevent" Then
        If LCase(Mid(Trim$(parse$(1)), 1, 7)) = "http://" Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopMidi
            Call StopSound
            frmFlash.Flash.LoadMovie 0, Trim$(parse$(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        ElseIf FileExist("Flashs\" & Trim$(parse$(1))) = True Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopMidi
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
        Call SendData("prompt" & SEP_CHAR & i & SEP_CHAR & Val#(parse$(2)) & SEP_CHAR & END_CHAR)
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
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Emoticons(i).Command)
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
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Element(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    If (casestring = "editelement") Then
        n = Val#(parse$(1))

        Element(n).Name = parse$(2)
        Element(n).Strong = Val#(parse$(3))
        Element(n).Weak = Val#(parse$(4))
        
        Call ElementEditorInit
        Exit Sub
    End If
    
    If (casestring = "updateelement") Then
        n = Val#(parse$(1))

        Element(n).Name = parse$(2)
        Element(n).Strong = Val#(parse$(3))
        Element(n).Weak = Val#(parse$(4))
        Exit Sub
    End If

    If (casestring = "editemoticon") Then
        n = Val#(parse$(1))

        Emoticons(n).Command = parse$(2)
        Emoticons(n).Pic = Val#(parse$(3))
        
        Call EmoticonEditorInit
        Exit Sub
    End If
    
    If (casestring = "updateemoticon") Then
        n = Val#(parse$(1))
        
        Emoticons(n).Command = parse$(2)
        Emoticons(n).Pic = Val#(parse$(3))
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
            frmIndex.lstIndex.AddItem i & ": " & Trim$(Arrows(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    If (casestring = "updatearrow") Then
        n = Val#(parse$(1))
        
        Arrows(n).Name = parse$(2)
        Arrows(n).Pic = Val#(parse$(3))
        Arrows(n).Range = Val#(parse$(4))
        Arrows(n).Amount = Val#(parse$(5))
        Exit Sub
    End If

    If (casestring = "editarrow") Then
        n = Val#(parse$(1))

        Arrows(n).Name = parse$(2)
        
        Call ArrowEditorInit
        Exit Sub
    End If
    
    If (casestring = "updatearrow") Then
        n = Val#(parse$(1))
        
        Arrows(n).Name = parse$(2)
        Arrows(n).Pic = Val#(parse$(3))
        Arrows(n).Range = Val#(parse$(4))
        Arrows(n).Amount = Val#(parse$(5))
        Exit Sub
    End If

    If (casestring = "checkarrows") Then
        n = Val#(parse$(1))
        z = Val#(parse$(2))
        i = Val#(parse$(3))
        p = Val#(parse$(4))
        
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
        n = Val#(parse$(1))
        
        Player(n).Sprite = Val#(parse$(2))
        Exit Sub
    End If
    
    If (casestring = "mapreport") Then
        n = 1
        
        frmMapReport.lstIndex.Clear
        For i = 1 To MAX_MAPS
            frmMapReport.lstIndex.AddItem i & ": " & Trim$(parse$(n))
            n = n + 1
        Next i
        
        frmMapReport.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (casestring = "time") Then
        GameTime = Val#(parse$(1))
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
        Wierd = Val#(parse$(1))
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
        
        SpellNum = Val#(parse$(1))
        
        Spell(SpellNum).SpellAnim = Val#(parse$(2))
        Spell(SpellNum).SpellTime = Val#(parse$(3))
        Spell(SpellNum).SpellDone = Val#(parse$(4))
        Spell(SpellNum).Big = Val#(parse$(9))
        
        Player(Val(parse$(5))).SpellNum = SpellNum
        
        For i = 1 To MAX_SPELL_ANIM
            If Player(Val(parse$(5))).SpellAnim(i).CastedSpell = NO Then
                Player(Val(parse$(5))).SpellAnim(i).SpellDone = 0
                Player(Val(parse$(5))).SpellAnim(i).SpellVar = 0
                Player(Val(parse$(5))).SpellAnim(i).SpellTime = GetTickCount
                Player(Val(parse$(5))).SpellAnim(i).TargetType = Val#(parse$(6))
                Player(Val(parse$(5))).SpellAnim(i).Target = Val#(parse$(7))
                Player(Val(parse$(5))).SpellAnim(i).CastedSpell = YES
                Exit For
            End If
        Next i
        Exit Sub
    End If
    
        ' :::::::::::::::::::::::
    ' :: Script Spell anim packet ::
    ' :::::::::::::::::::::::
    If (casestring = "scriptspellanim") Then
    
    ' THIS DOESNT NEED THESE WTF??/
        Spell(Val(parse$(1))).SpellAnim = Val#(parse$(2))
        Spell(Val(parse$(1))).SpellTime = Val#(parse$(3))
        Spell(Val(parse$(1))).SpellDone = Val#(parse$(4))
        Spell(Val(parse$(1))).Big = Val#(parse$(7))
        
        
        For i = 1 To MAX_SCRIPTSPELLS
            If ScriptSpell(i).CastedSpell = NO Then
                ScriptSpell(i).SpellNum = Val#(parse$(1))
                ScriptSpell(i).SpellDone = 0
                ScriptSpell(i).SpellVar = 0
                ScriptSpell(i).SpellTime = GetTickCount
                ScriptSpell(i).x = Val#(parse$(5))
                ScriptSpell(i).y = Val#(parse$(6))
                ScriptSpell(i).CastedSpell = YES
                Exit For
            End If
        Next i
        Exit Sub
    End If
    
    If (casestring = "checkemoticons") Then
        n = Val#(parse$(1))
        
        Player(n).EmoticonNum = Val#(parse$(2))
        Player(n).EmoticonTime = GetTickCount
        Player(n).EmoticonVar = 0
        Exit Sub
    End If
    
    If casestring = "updatesell" Then
        frmSellItem.lstSellItem.Clear
        For i = 1 To MAX_INV
          If GetPlayerInvItemNum(MyIndex, i) > 0 Then
                   If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
                    frmSellItem.lstSellItem.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerNecklaceSlot(MyIndex) = i Then
                        frmSellItem.lstSellItem.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                    Else
                        frmSellItem.lstSellItem.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                    End If
                End If
                       Else
           frmSellItem.lstSellItem.AddItem i & "> None"
       End If
   Next i
   frmSellItem.lstSellItem.ListIndex = 0
        Exit Sub
    End If
    
    If casestring = "levelup" Then
        Player(Val(parse$(1))).LevelUpT = GetTickCount
        Player(Val(parse$(1))).LevelUp = 1
        Exit Sub
    End If
    
    If casestring = "damagedisplay" Then
        For i = 1 To MAX_BLT_LINE
            If Val#(parse$(1)) = 0 Then
                If BattlePMsg(i).Index <= 0 Then
                    BattlePMsg(i).Index = 1
                    BattlePMsg(i).Msg = parse$(2)
                    BattlePMsg(i).Color = Val#(parse$(3))
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
                    BattleMMsg(i).Color = Val#(parse$(3))
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
        If Val#(parse$(1)) = 0 Then
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then
                    If BattlePMsg(i).y < BattlePMsg(i + 1).y Then z = i
                Else
                    If BattlePMsg(i).y < BattlePMsg(1).y Then z = i
                End If
            Next i
                        
            BattlePMsg(z).Index = 1
            BattlePMsg(z).Msg = parse$(2)
            BattlePMsg(z).Color = Val#(parse$(3))
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
            BattleMMsg(z).Color = Val#(parse$(3))
            BattleMMsg(z).Time = GetTickCount
            BattleMMsg(z).Done = 1
            BattleMMsg(z).y = 0
        End If
        Exit Sub
    End If
    
    If casestring = "itembreak" Then
        ItemDur(Val(parse$(1))).Item = Val#(parse$(2))
        ItemDur(Val(parse$(1))).Dur = Val#(parse$(3))
        ItemDur(Val(parse$(1))).Done = 1
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::::::
    ' :: Index player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::::::::
    If casestring = "itemworn" Then
        Player(Val(parse$(1))).Armor = Val#(parse$(2))
        Player(Val(parse$(1))).Weapon = Val#(parse$(3))
        Player(Val(parse$(1))).Helmet = Val#(parse$(4))
        Player(Val(parse$(1))).Shield = Val#(parse$(5))
        Player(Val(parse$(1))).Legs = Val#(parse$(6))
        Player(Val(parse$(1))).Ring = Val#(parse$(7))
        Player(Val(parse$(1))).Necklace = Val#(parse$(8))
        
        'Call AddText("index:" & Val(parse$(1)) & " armor:" & Val#(parse$(2)) & " wep:" & Val#(parse$(3)) & " helm:" & Val#(parse$(4)) & " shield:" & Val#(parse$(5)) & " legs:" & Val#(parse$(6)) & " ring:" & Val#(parse$(7)) & " necklace:" & Val#(parse$(8)), AlertColor)
        'Call AddText(Player(parse(1)).Armor, Red)
        'Call AddText(Player(1).Armor, Red)
        
        Exit Sub
        
    End If
    
    If casestring = "scripttile" Then
        frmScript.lblScript.Caption = parse$(1)
        Exit Sub
    End If
    
    If (casestring = "forceclosehouse") Then
        Call HouseEditorCancel
    End If
    
    If (casestring = "showcustommenu") Then
        
        CUSTOM_TITLE = parse$(1)
        CUSTOM_IS_CLOSABLE = Val(parse(3))
        
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
        CustomX = Val(parse$(3))
        CustomY = Val(parse$(4))
        
        If strfilename = "" Then
        strfilename = "MEGAUBERBLANKNESSOFUNHOLYPOWER"
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
        CustomX = Val(parse$(3))
        CustomY = Val(parse$(4))
        customsize = Val(parse$(5))
        customcolour = Val(parse$(6))
        
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
        CustomX = Val(parse$(3))
        CustomY = Val(parse$(4))
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
        Call sleep(1)
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

Sub SendData(ByVal Data As String)
    If IsConnected Then
        frmMirage.Socket.SendData Data
        DoEvents
    End If
End Sub

Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
Dim packet As String

    packet = "newfaccountied" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & END_CHAR
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

Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Slot As Long)
Dim packet As String

    packet = "addachara" & SEP_CHAR & Trim$(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendDelChar(ByVal Slot As Long)
Dim packet As String
    
    packet = "delimbocharu" & SEP_CHAR & Slot & SEP_CHAR & END_CHAR
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

    packet = "MAPDATA" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim$(Map(GetPlayerMap(MyIndex)).Name) & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Revision & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Moral & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Up & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Down & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Left & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Right & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Music & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootMap & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootX & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootY & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Indoors & SEP_CHAR
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                packet = packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .Light & SEP_CHAR
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

    packet = "SAVEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    packet = packet & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
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
    
    packet = "SAVENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).Speed & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).EXP & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR & Npc(NpcNum).Element & SEP_CHAR
    For i = 1 To MAX_NPC_DROPS
        packet = packet & Npc(NpcNum).ItemNPC(i).Chance
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

Sub SendSaveShop(ByVal ShopNum As Long)
Dim packet As String
Dim i As Long, z As Long

    packet = "SAVESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Trim$(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim$(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To 7
        For z = 1 To MAX_TRADES
            packet = packet & Shop(ShopNum).TradeItem(i).Value(z).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).Value(z).GetValue & SEP_CHAR
        Next z
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
