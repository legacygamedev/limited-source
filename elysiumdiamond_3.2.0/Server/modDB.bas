Attribute VB_Name = "modDB"
Option Explicit

Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)

Public Sub InitEditor()

    If Not FileExist("Data.ini") Then
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "GameName", "Elysium"
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "WebSite", vbNullString
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "Port", 4000
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "HPRegen", 1
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "MPRegen", 1
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "SPRegen", 1
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "Scrolling", 1

        'SpecialPutVar App.Path & "\Data.ini", "CONFIG", "AutoTurn", 0
        SpecialPutVar App.Path & "\Data.ini", "CONFIG", "SCRIPTING", 1
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_PLAYERS", 25
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_ITEMS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_NPCS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_SHOPS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_SPELLS", 100
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_MAPS", 200
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS", 20
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_GUILDS", 20
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS", 10
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS", 10
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_LEVEL", 500
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_PARTIES", 20
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS", 4
        SpecialPutVar App.Path & "\Data.ini", "MAX", "MAX_SPEECH", 25
    End If

    If LCase$(Dir$(App.Path & "\accounts", vbDirectory)) <> "accounts" Then
        Call MkDir$(App.Path & "\Accounts")
    End If
    
    frmStart.lblStatus.Caption = "Loading accounts..."
    DoEvents
    
    MAX_PLAYERS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_PLAYERS"))
    MAX_SPELLS = Val(GetVar(App.Path & "\Data.ini", "MAX", "MAX_SPELLS"))
    
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    
    frmStart.lblStatus.Visible = False
    
    frmStart.lblChoose.Visible = True
    frmStart.txtChoose.Visible = True
    frmStart.cmdAccept.Visible = True
    DoEvents

End Sub

Public Sub InitStart()
Dim I As Long

    With frmRun
        .lblAccountName.Caption = "Account Name: " & Player(1).Login
        .lblAccountPassword.Caption = "Account Password: " & Player(1).Password
        
        If Player(1).InGame = YES Then
            .lblPlayerOnOff.Caption = "Note: PLAYER IS ON, ONLY EDIT ACCESS"
        Else
            .lblPlayerOnOff.Caption = "Note: PLAYER IS OFF, DO WHAT YOU WISH"
        End If
        
        For I = 1 To MAX_CHARS
            .lblCName(I).Text = Player(1).Char(I).Name
            .lblCLevel(I).Text = Player(1).Char(I).Class
            .lblCAccess(I).Text = Player(1).Char(I).Access
            .lblCClass(I).Text = Player(1).Char(I).Class
            .lblCSprite(I).Text = Player(1).Char(I).Sprite
            .lblCHP(I).Text = Player(1).Char(I).HP
            .lblCStr(I).Text = Player(1).Char(I).STR
            .lblCDef(I).Text = Player(1).Char(I).DEF
            .lblCSpeed(I).Text = Player(1).Char(I).Speed
        Next I
    End With
    
    DoEvents

    For I = 1 To MAX_CHARS
        With frmRun
            If .lblCName(I).Text = "" Then
                .chkUsed(I).Value = NO
                .lblCName(I).Text = vbNullString
                .lblCLevel(I).Text = vbNullString
                .lblCAccess(I).Text = vbNullString
                .lblCClass(I).Text = vbNullString
                .lblCSprite(I).Text = vbNullString
                .lblCHP(I).Text = vbNullString
                .lblCStr(I).Text = vbNullString
                .lblCDef(I).Text = vbNullString
                .lblCSpeed(I).Text = vbNullString
            Else
                .chkUsed(I).Value = YES
            End If
        End With
    Next I
    
    DoEvents

End Sub

Public Sub DoSave()
Dim I As Long
Dim NOPE As Byte

    NOPE = YES

    If Player(1).InGame = YES Then
        For I = 1 To MAX_CHARS
            With Player(1)
                If .Char(I).Name <> frmRun.lblCName(I).Text Then NOPE = NO
                If .Char(I).Level <> frmRun.lblCLevel(I).Text Then NOPE = NO
                If .Char(I).Class <> frmRun.lblCClass(I).Text Then NOPE = NO
                If .Char(I).Sprite <> frmRun.lblCClass(I).Text Then NOPE = NO
                If .Char(I).HP <> frmRun.lblCHP(I).Text Then NOPE = NO
                If .Char(I).STR <> frmRun.lblCStr(I).Text Then NOPE = NO
                If .Char(I).DEF <> frmRun.lblCDef(I).Text Then NOPE = NO
                If .Char(I).Speed <> frmRun.lblCSpeed(I).Text Then NOPE = NO
            End With
        Next I
    End If
    
    If NOPE = NO Then
        MsgBox "The character is logged in, thus you cannot edit his / her account except for the Access.", vbOKOnly, "Uh oh!"
        Exit Sub
    End If

    For I = 1 To MAX_CHARS
        With Player(1)
            If frmRun.chkUsed(I).Value = YES Then
                .Char(I).Name = frmRun.lblCName(I).Text
                .Char(I).Level = Val(frmRun.lblCLevel(I).Text)
                .Char(I).Access = Val(frmRun.lblCAccess(I).Text)
                .Char(I).Class = Val(frmRun.lblCClass(I).Text)
                .Char(I).Sprite = Val(frmRun.lblCSprite(I).Text)
                .Char(I).HP = Val(frmRun.lblCHP(I).Text)
                .Char(I).STR = Val(frmRun.lblCStr(I).Text)
                .Char(I).DEF = Val(frmRun.lblCDef(I).Text)
                .Char(I).Speed = Val(frmRun.lblCSpeed(I).Text)
            End If
        End With
    Next I

    Call SavePlayer(1)
    
    Call ClearMe

End Sub

Public Sub ClearMe()

    Call ClearPlayer(1)
    frmRun.Hide
    frmStart.Show
    frmStart.txtChoose.Text = vbNullString

End Sub

Function AccountExist(ByVal Name As String) As Boolean
    Dim FileName As String

    FileName = "accounts\" & Trim$(Name) & ".ini"

    If FileExist(FileName) Then
        AccountExist = True
    Else
        AccountExist = False
    End If

End Function

Public Function ExistVar(File As String, Header As String, Var As String) As Boolean

    ExistVar = (GetVar(File, Header, Var) <> "")

End Function

Function FileExist(ByVal FileName As String) As Boolean

    If Dir$(App.Path & "\" & FileName) = vbNullString Then
        FileExist = False
    Else
        FileExist = True
    End If

End Function

Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, _
   Header As String, _
   Var As String, _
   Value As String)

    If Trim$(Value) = "0" Or Trim$(Value) = vbNullString Then
        If ExistVar(File, Header, Var) Then
            Call DelVar(File, Header, Var)
        End If

    Else
        Call WritePrivateProfileString(Header, Var, Value, File)
    End If

End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim I As Long
    Dim N As Long

    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString

    For I = 1 To MAX_CHARS
        Player(Index).Char(I).Name = vbNullString
        Player(Index).Char(I).Class = 1
        Player(Index).Char(I).Level = 0
        Player(Index).Char(I).Sprite = 0
        Player(Index).Char(I).Exp = 0
        Player(Index).Char(I).Access = 0
        Player(Index).Char(I).PK = NO
        Player(Index).Char(I).POINTS = 0
        Player(Index).Char(I).Guild = vbNullString
        Player(Index).Char(I).HP = 0
        Player(Index).Char(I).MP = 0
        Player(Index).Char(I).SP = 0
        Player(Index).Char(I).STR = 0
        Player(Index).Char(I).DEF = 0
        Player(Index).Char(I).Speed = 0
        Player(Index).Char(I).Magi = 0

        For N = 1 To MAX_INV
            Player(Index).Char(I).Inv(N).num = 0
            Player(Index).Char(I).Inv(N).Value = 0
            Player(Index).Char(I).Inv(N).Dur = 0
        Next

        For N = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(I).Spell(N) = 0
        Next

        Player(Index).Char(I).ArmorSlot = 0
        Player(Index).Char(I).WeaponSlot = 0
        Player(Index).Char(I).HelmetSlot = 0
        Player(Index).Char(I).ShieldSlot = 0
        Player(Index).Char(I).Map = 0
        Player(Index).Char(I).x = 0
        Player(Index).Char(I).y = 0
        Player(Index).Char(I).Dir = 0

        For N = 1 To MAX_FRIENDS
            Player(Index).Char(I).Friends(N) = vbNullString
        Next
    Next

    Player(Index).Pet.Alive = NO

    ' Temporary vars
    Player(Index).Buffer = vbNullString
    Player(Index).IncBuffer = vbNullString
    Player(Index).CharNum = 0
    Player(Index).InGame = False
    Player(Index).AttackTimer = 0
    Player(Index).DataTimer = 0
    Player(Index).DataBytes = 0
    Player(Index).DataPackets = 0
    Player(Index).PartyID = 0
    Player(Index).InParty = 0
    Player(Index).Invited = 0
    Player(Index).Target = 0
    Player(Index).TargetType = 0
    Player(Index).CastedSpell = NO
    Player(Index).GettingMap = NO
    Player(Index).Emoticon = -1
    Player(Index).InTrade = 0
    Player(Index).TradePlayer = 0
    Player(Index).TradeOk = 0
    Player(Index).TradeItemMax = 0
    Player(Index).TradeItemMax2 = 0

    'For N = 1 To MAX_PLAYER_TRADES
    '    Player(Index).Trading(N).InvName = vbNullString
    '    Player(Index).Trading(N).InvNum = 0
    'Next

    Player(Index).ChatPlayer = 0
End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim FileName As String
    Dim I As Long
    Dim N As Long

    FileName = App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini"
    Call PutVar(FileName, "GENERAL", "Login", Trim$(Player(Index).Login))
    Call PutVar(FileName, "GENERAL", "Password", Trim$(Player(Index).Password))

    For I = 1 To MAX_CHARS

        ' General
        Call PutVar(FileName, "CHAR" & I, "Name", Trim$(Player(Index).Char(I).Name))
        Call PutVar(FileName, "CHAR" & I, "Class", STR(Player(Index).Char(I).Class))
        Call PutVar(FileName, "CHAR" & I, "Sex", STR(Player(Index).Char(I).Sex))
        Call PutVar(FileName, "CHAR" & I, "Sprite", STR(Player(Index).Char(I).Sprite))
        Call PutVar(FileName, "CHAR" & I, "Level", STR(Player(Index).Char(I).Level))
        Call PutVar(FileName, "CHAR" & I, "Exp", STR(Player(Index).Char(I).Exp))
        Call PutVar(FileName, "CHAR" & I, "Access", STR(Player(Index).Char(I).Access))
        Call PutVar(FileName, "CHAR" & I, "PK", STR(Player(Index).Char(I).PK))
        Call PutVar(FileName, "CHAR" & I, "Guild", Trim$(Player(Index).Char(I).Guild))
        Call PutVar(FileName, "CHAR" & I, "Guildaccess", STR(Player(Index).Char(I).Guildaccess))

        ' Vitals
        Call PutVar(FileName, "CHAR" & I, "HP", STR(Player(Index).Char(I).HP))
        Call PutVar(FileName, "CHAR" & I, "MP", STR(Player(Index).Char(I).MP))
        Call PutVar(FileName, "CHAR" & I, "SP", STR(Player(Index).Char(I).SP))

        ' Stats
        Call PutVar(FileName, "CHAR" & I, "str", STR(Player(Index).Char(I).STR))
        Call PutVar(FileName, "CHAR" & I, "DEF", STR(Player(Index).Char(I).DEF))
        Call PutVar(FileName, "CHAR" & I, "SPEED", STR(Player(Index).Char(I).Speed))
        Call PutVar(FileName, "CHAR" & I, "MAGI", STR(Player(Index).Char(I).Magi))
        Call PutVar(FileName, "CHAR" & I, "POINTS", STR(Player(Index).Char(I).POINTS))

        ' Worn equipment
        Call PutVar(FileName, "CHAR" & I, "ArmorSlot", STR(Player(Index).Char(I).ArmorSlot))
        Call PutVar(FileName, "CHAR" & I, "WeaponSlot", STR(Player(Index).Char(I).WeaponSlot))
        Call PutVar(FileName, "CHAR" & I, "HelmetSlot", STR(Player(Index).Char(I).HelmetSlot))
        Call PutVar(FileName, "CHAR" & I, "ShieldSlot", STR(Player(Index).Char(I).ShieldSlot))

        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(I).Map = 0 Then
            Player(Index).Char(I).Map = START_MAP
            Player(Index).Char(I).x = START_X
            Player(Index).Char(I).y = START_Y
        End If

        ' Position
        Call PutVar(FileName, "CHAR" & I, "Map", STR(Player(Index).Char(I).Map))
        Call PutVar(FileName, "CHAR" & I, "X", STR(Player(Index).Char(I).x))
        Call PutVar(FileName, "CHAR" & I, "Y", STR(Player(Index).Char(I).y))
        Call PutVar(FileName, "CHAR" & I, "Dir", STR(Player(Index).Char(I).Dir))

        ' Inventory
        For N = 1 To MAX_INV
            Call PutVar(FileName, "CHAR" & I, "InvItemNum" & N, STR(Player(Index).Char(I).Inv(N).num))
            Call PutVar(FileName, "CHAR" & I, "InvItemVal" & N, STR(Player(Index).Char(I).Inv(N).Value))
            Call PutVar(FileName, "CHAR" & I, "InvItemDur" & N, STR(Player(Index).Char(I).Inv(N).Dur))
        Next

        ' Spells
        For N = 1 To MAX_PLAYER_SPELLS
            Call PutVar(FileName, "CHAR" & I, "Spell" & N, STR(Player(Index).Char(I).Spell(N)))
        Next

        ' Pet
        If I = Player(Index).CharNum Then
            If Player(Index).Pet.Alive = YES Then
                Call PutVar(FileName, "CHAR" & I, "HasPet", 1)
                Call PutVar(FileName, "CHAR" & I, "Pet", STR(Player(Index).Pet.Sprite))
                Call PutVar(FileName, "CHAR" & I, "PetLevel", STR(Player(Index).Pet.Level))
            Else
                Call PutVar(FileName, "CHAR" & I, "HasPet", 0)
                Call DelVar(FileName, "CHAR" & I, "Pet") ' Saving space
                Call DelVar(FileName, "CHAR" & I, "PetLevel")
            End If

        Else
            Call PutVar(FileName, "CHAR" & I, "HasPet", 0)
            Call DelVar(FileName, "CHAR" & I, "Pet") ' Saving space
            Call DelVar(FileName, "CHAR" & I, "PetLevel")
        End If

        ' Friend list
        For N = 1 To MAX_FRIENDS
            Call PutVar(FileName, "CHAR" & I, "Friend" & N, Player(Index).Char(I).Friends(N))
        Next
    Next

End Sub

Public Sub DelVar(sFileName As String, _
   sSection As String, _
   sKey As String)

    If Len(Trim$(sKey)) <> 0 Then
        WritePrivateProfileString sSection, sKey, vbNullString, sFileName
    Else
        WritePrivateProfileString sSection, sKey, vbNullString, sFileName
    End If

End Sub

Sub LoadPlayer(ByVal Index As Long, _
   ByVal Name As String)
    Dim FileName As String
    Dim I As Long
    Dim N As Long

    'Call ClearPlayer(Index)
    FileName = App.Path & "\accounts\" & Trim$(Name) & ".ini"
    Player(Index).Login = GetVar(FileName, "GENERAL", "Login")
    Player(Index).Password = GetVar(FileName, "GENERAL", "Password")
    Player(Index).InGame = Val(GetVar(FileName, "GENERAL", "InGame"))
    Player(Index).Pet.Alive = NO

    For I = 1 To MAX_CHARS

        ' General
        Player(Index).Char(I).Name = GetVar(FileName, "CHAR" & I, "Name")
        Player(Index).Char(I).Sex = Val(GetVar(FileName, "CHAR" & I, "Sex"))
        Player(Index).Char(I).Class = Val(GetVar(FileName, "CHAR" & I, "Class"))

        If Player(Index).Char(I).Class = 0 Then Player(Index).Char(I).Class = 1
        Player(Index).Char(I).Sprite = Val(GetVar(FileName, "CHAR" & I, "Sprite"))
        Player(Index).Char(I).Level = Val(GetVar(FileName, "CHAR" & I, "Level"))
        Player(Index).Char(I).Exp = Val(GetVar(FileName, "CHAR" & I, "Exp"))
        Player(Index).Char(I).Access = Val(GetVar(FileName, "CHAR" & I, "Access"))
        Player(Index).Char(I).PK = Val(GetVar(FileName, "CHAR" & I, "PK"))
        Player(Index).Char(I).Guild = GetVar(FileName, "CHAR" & I, "Guild")
        Player(Index).Char(I).Guildaccess = Val(GetVar(FileName, "CHAR" & I, "Guildaccess"))

        ' Vitals
        Player(Index).Char(I).HP = Val(GetVar(FileName, "CHAR" & I, "HP"))
        Player(Index).Char(I).MP = Val(GetVar(FileName, "CHAR" & I, "MP"))
        Player(Index).Char(I).SP = Val(GetVar(FileName, "CHAR" & I, "SP"))

        ' Stats
        Player(Index).Char(I).STR = Val(GetVar(FileName, "CHAR" & I, "str"))
        Player(Index).Char(I).DEF = Val(GetVar(FileName, "CHAR" & I, "DEF"))
        Player(Index).Char(I).Speed = Val(GetVar(FileName, "CHAR" & I, "SPEED"))
        Player(Index).Char(I).Magi = Val(GetVar(FileName, "CHAR" & I, "MAGI"))
        Player(Index).Char(I).POINTS = Val(GetVar(FileName, "CHAR" & I, "POINTS"))

        ' Worn equipment
        Player(Index).Char(I).ArmorSlot = Val(GetVar(FileName, "CHAR" & I, "ArmorSlot"))
        Player(Index).Char(I).WeaponSlot = Val(GetVar(FileName, "CHAR" & I, "WeaponSlot"))
        Player(Index).Char(I).HelmetSlot = Val(GetVar(FileName, "CHAR" & I, "HelmetSlot"))
        Player(Index).Char(I).ShieldSlot = Val(GetVar(FileName, "CHAR" & I, "ShieldSlot"))

        ' Position
        Player(Index).Char(I).Map = Val(GetVar(FileName, "CHAR" & I, "Map"))
        Player(Index).Char(I).x = Val(GetVar(FileName, "CHAR" & I, "X"))
        Player(Index).Char(I).y = Val(GetVar(FileName, "CHAR" & I, "Y"))
        Player(Index).Char(I).Dir = Val(GetVar(FileName, "CHAR" & I, "Dir"))

        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(I).Map = 0 Then
            Player(Index).Char(I).Map = START_MAP
            Player(Index).Char(I).x = START_X
            Player(Index).Char(I).y = START_Y
        End If

        ' Inventory
        For N = 1 To MAX_INV
            Player(Index).Char(I).Inv(N).num = Val(GetVar(FileName, "CHAR" & I, "InvItemNum" & N))
            Player(Index).Char(I).Inv(N).Value = Val(GetVar(FileName, "CHAR" & I, "InvItemVal" & N))
            Player(Index).Char(I).Inv(N).Dur = Val(GetVar(FileName, "CHAR" & I, "InvItemDur" & N))
        Next

        ' Spells
        For N = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(I).Spell(N) = Val(GetVar(FileName, "CHAR" & I, "Spell" & N))
        Next

        If Val(GetVar(FileName, "CHAR" & I, "HasPet")) = 1 Then
            Player(Index).Pet.Sprite = Val(GetVar(FileName, "CHAR" & I, "Pet"))
            Player(Index).Pet.Alive = YES
            Player(Index).Pet.Dir = DIR_UP
            Player(Index).Pet.Map = Player(Index).Char(I).Map
            Player(Index).Pet.x = Player(Index).Char(I).x + Int((Rnd * 3) - 1)

            'If Player(Index).Pet.x < 0 Or Player(Index).Pet.x > MAX_MAPX Then Player(Index).Pet.x = GetPlayerX(Index)
            'Player(Index).Pet.y = Player(Index).Char(i).y + Int((Rnd * 3) - 1)

            'If Player(Index).Pet.y < 0 Or Player(Index).Pet.y > MAX_MAPY Then Player(Index).Pet.y = GetPlayerY(Index)
            Player(Index).Pet.MapToGo = 0
            Player(Index).Pet.XToGo = -1
            Player(Index).Pet.YToGo = -1
            Player(Index).Pet.Level = Val(GetVar(FileName, "CHAR" & I, "PetLevel"))
            Player(Index).Pet.HP = Player(Index).Pet.Level * 5 '???
        End If

        For N = 1 To MAX_FRIENDS
            Player(Index).Char(I).Friends(N) = GetVar(FileName, "CHAR" & I, "Friend" & N)
        Next
    Next

End Sub

Sub SpecialPutVar(File As String, _
   Header As String, _
   Var As String, _
   Value As String)

    ' Same as the one below except it keeps all 0 and blank values (used for config)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub
