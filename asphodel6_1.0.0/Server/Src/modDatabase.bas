Attribute VB_Name = "modDatabase"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

' Outputs string to text file
Public Sub AddLog(ByVal Text As String, ByVal FN As String)
Dim FileName As String
Dim F As Integer

    If ServerLog Then
        FileName = "\logs\" & FN
    
        If Not FileExist(FileName) Then
            F = FreeFile
            Open App.Path & FileName For Output As #F
            Close #F
        End If
    
        F = FreeFile
        Open App.Path & FileName For Append As #F
            Print #F, Time & ": " & Text
        Close #F
    End If
End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = vbNullString
  
    sSpaces = Space$(5000)
  
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal FileName As String) As Boolean
    FileExist = LenB(Dir$(App.Path & "\" & FileName)) > 0
End Function

Public Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
Dim FileName As String
Dim F As Long
Dim i As Long
Dim BIPIndex As Long
Dim BACIndex As Long

    FileName = App.Path & "\data\bans.ini"
    
    ' Make sure the file exists
    If Not FileExist("data\bans.ini") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
    
    BIPIndex = Val(GetVar(FileName, "IP", "Total"))
    
    If BIPIndex = 0 Then
        BIPIndex = 1
        PutVar FileName, "IP", "Total", CStr(BIPIndex)
    Else
        For i = 1 To BIPIndex
            If GetVar(FileName, "IP", "IP" & i) = vbNullString Then BIPIndex = i: Exit For
        Next
        If BIPIndex = Val(GetVar(FileName, "IP", "Total")) Then
            BIPIndex = BIPIndex + 1
            PutVar FileName, "IP", "Total", CStr(BIPIndex)
        End If
    End If
    
    PutVar FileName, "IP", "IP" & BIPIndex, GetPlayerIP(BanPlayerIndex)
    
    BACIndex = Val(GetVar(FileName, "ACCOUNT", "Total"))
    
    If BACIndex = 0 Then
        BACIndex = 1
        PutVar FileName, "ACCOUNT", "Total", CStr(BACIndex)
    Else
        For i = 1 To BACIndex
            If GetVar(FileName, "ACCOUNT", "Account" & i) = vbNullString Then BACIndex = i: Exit For
        Next
        If BACIndex = Val(GetVar(FileName, "ACCOUNT", "Total")) Then
            BACIndex = BACIndex + 1
            PutVar FileName, "ACCOUNT", "Total", CStr(BACIndex)
        End If
    End If

    
    PutVar FileName, "ACCOUNT", "Account" & BACIndex, GetPlayerLogin(BanPlayerIndex)
    
    If BannedByIndex = 0 Then
        Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by the server!", Color.White)
        Call AddLog("Server has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
        Call AlertMsg(BanPlayerIndex, "You have been banned by the server!")
    Else
        Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", Color.White)
        Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
        Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
    End If
    
    Load_BanTable
    
End Sub

Public Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
    BanIndex BanPlayerIndex, 0
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
Dim FileName As String

    FileName = "accounts\" & Trim$(Name) & ".bin"
    AccountExist = FileExist(FileName)

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim FileName As String
Dim RightPassword As String * NAME_LENGTH
Dim nFileNum As Integer

    If AccountExist(Name) Then
    
        FileName = App.Path & "\Accounts\" & Trim$(Name) & ".bin"
        
        nFileNum = FreeFile
        
        Open FileName For Binary As #nFileNum
            Get #nFileNum, NAME_LENGTH, RightPassword
        Close #nFileNum
        
        PasswordOK = (Trim$(Password) = Trim$(RightPassword))
        
    End If
    
End Function

Public Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
Dim i As Long

    Player(Index).Login = Name
    Player(Index).Password = Password
    
    For i = 1 To MAX_CHARS
        Call ClearChar(Index, i)
    Next
    
    Call SavePlayer(Index)
    
End Sub

Public Sub DeleteName(ByVal Name As String)
Dim f1 As Long
Dim f2 As Long
Dim s As String

    FileCopy App.Path & "\accounts\charlist.txt", App.Path & "\accounts\chartemp.txt"
    
    ' Destroy name from charlist
    f1 = FreeFile
    
    Open App.Path & "\accounts\chartemp.txt" For Input As #f1
    
        f2 = FreeFile
        
        Open App.Path & "\accounts\charlist.txt" For Output As #f2
            Do While Not EOF(f1)
                Input #f1, s
                If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
                    Print #f2, s
                End If
            Loop
        Close #f1
        
    Close #f2
    
    Call Kill(App.Path & "\accounts\chartemp.txt")
    
End Sub

' ****************
' ** Characters **
' ****************

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    CharExist = (LenB(Trim$(Player(Index).Char(CharNum).Name)) > 0)
End Function

Public Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
Dim F As Long
Dim n As Long

    If LenB(Trim$(Player(Index).Char(CharNum).Name)) = 0 Then
    
        TempPlayer(Index).CharNum = CharNum
        
        Player(Index).Char(CharNum).Name = Name
        Player(Index).Char(CharNum).Sex = Sex
        Player(Index).Char(CharNum).Class = ClassNum
        
        If Player(Index).Char(CharNum).Sex = GenderType.Male_ Then
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).Sprite
        Else
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).Sprite
        End If
        
        Player(Index).Char(CharNum).Level = 1
        
        For n = 1 To Stats.Stat_Count - 1
            Player(Index).Char(CharNum).Stat(n) = Class(ClassNum).Stat(n)
        Next
        
        Player(Index).Char(CharNum).Map = Class(ClassNum).StartLoc.MapNum
        Player(Index).Char(CharNum).X = Class(ClassNum).StartLoc.X
        Player(Index).Char(CharNum).Y = Class(ClassNum).StartLoc.Y
        
        Player(Index).Char(CharNum).Vital(Vitals.HP) = GetPlayerMaxVital(Index, Vitals.HP)
        Player(Index).Char(CharNum).Vital(Vitals.MP) = GetPlayerMaxVital(Index, Vitals.MP)
        Player(Index).Char(CharNum).Vital(Vitals.SP) = GetPlayerMaxVital(Index, Vitals.SP)
        
        ' Append name to file
        F = FreeFile
        
        Open App.Path & "\accounts\charlist.txt" For Append As #F
            Print #F, Name
        Close #F
        
        Call SavePlayer(Index)
        
    End If
End Sub

Public Sub DelChar(ByVal Index As Long, ByVal CharNum As Long)
    Call DeleteName(Player(Index).Char(CharNum).Name)
    Call ClearChar(Index, CharNum)
    Call SavePlayer(Index)
End Sub

Function FindChar(ByVal Name As String) As Boolean
Dim F As Long
Dim s As String
    
    F = FreeFile
    Open App.Path & "\Accounts\charlist.txt" For Input As #F
        Do While Not EOF(F)
            Input #F, s
            
            If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
                FindChar = True
                Close #F
                Exit Function
            End If
        Loop
    Close #F
    
End Function

' *************
' ** Players **
' *************

Public Sub SaveAllPlayersOnline()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next
    
End Sub

Public Sub SavePlayer(ByVal Index As Long)
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".bin"
    
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Player(Index)
    Close #F
    
End Sub

Public Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
Dim FileName As String
Dim F As Long

    Call ClearPlayer(Index)
    
    FileName = App.Path & "\accounts\" & Trim$(Name) & ".bin"
    
    F = FreeFile
    
    Open FileName For Binary As #F
        Get #F, , Player(Index)
    Close #F
    
End Sub

Public Sub ClearPlayer(ByVal Index As Long)
Dim i As Byte

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    
    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    TempPlayer(Index).Buffer = vbNullString
    
    For i = 0 To MAX_CHARS
        Player(Index).Char(i).Name = vbNullString
        Player(Index).Char(i).Class = 1
    Next
    
    frmServer.lstPlayers.List(Index - 1) = Index & ") None"
    
End Sub

Public Sub ClearChar(ByVal Index As Long, ByVal CharNum As Long)
    Call ZeroMemory(ByVal VarPtr(Player(Index).Char(CharNum)), LenB(Player(Index).Char(CharNum)))
    Player(Index).Char(CharNum).Name = vbNullString
    Player(Index).Char(CharNum).Class = 1
End Sub

' *************
' ** Classes **
' *************

Public Sub CreateClassesINI()
Dim FileName As String
Dim File As String

    FileName = "\data\classes.ini"
    
    Max_Classes = 2
    
    If Not FileExist(FileName) Then
        File = FreeFile
    
        Open App.Path & FileName For Output As File
            Print #File, "[INIT]"
            Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Public Sub LoadClasses()
Dim FileName As String
Dim i As Long

    If CheckClasses Then
        ReDim Preserve Class(1 To Max_Classes) As ClassRec
        Call SaveClasses
    Else
        FileName = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
        ReDim Preserve Class(1 To Max_Classes) As ClassRec
    End If
    
    ClearClasses
    
    For i = 1 To Max_Classes
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).Sprite = Val(GetVar(FileName, "CLASS" & i, "Sprite"))
        Class(i).Stat(Stats.Strength) = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class(i).Stat(Stats.Defense) = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class(i).Stat(Stats.SPEED) = Val(GetVar(FileName, "CLASS" & i, "Speed"))
        Class(i).Stat(Stats.Magic) = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
        
        If Val(GetVar(FileName, "CLASS" & i, "StartMap")) < 1 Or Val(GetVar(FileName, "CLASS" & i, "StartMap")) > MAX_MAPS Then
            PutVar FileName, "CLASS" & i, "StartMap", 1
        End If
        
        Class(i).StartLoc.MapNum = Val(GetVar(FileName, "CLASS" & i, "StartMap"))
        
        If Val(GetVar(FileName, "CLASS" & i, "StartX")) < 0 Or Val(GetVar(FileName, "CLASS" & i, "StartX")) > MAX_MAPX Then
            PutVar FileName, "CLASS" & i, "StartX", 1
        End If
        
        Class(i).StartLoc.X = Val(GetVar(FileName, "CLASS" & i, "StartX"))
        
        If Val(GetVar(FileName, "CLASS" & i, "StartY")) < 1 Or Val(GetVar(FileName, "CLASS" & i, "StartY")) > MAX_MAPY Then
            PutVar FileName, "CLASS" & i, "StartY", 1
        End If
        
        Class(i).StartLoc.Y = Val(GetVar(FileName, "CLASS" & i, "StartY"))
        
        If Val(GetVar(FileName, "CLASS" & i, "PointsPerLevel")) < 0 Then
            PutVar FileName, "CLASS" & i, "StartY", 0
        End If
        
        Class(i).PointsPerLevel = Val(GetVar(FileName, "CLASS" & i, "PointsPerLevel"))
        
        DoEvents
    Next
    
End Sub

Public Sub SaveClasses()
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\data\classes.ini"
    
    For i = 1 To Max_Classes
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "Sprite", CStr(Class(i).Sprite))
        Call PutVar(FileName, "CLASS" & i, "STR", CStr(Class(i).Stat(Stats.Strength)))
        Call PutVar(FileName, "CLASS" & i, "DEF", CStr(Class(i).Stat(Stats.Defense)))
        Call PutVar(FileName, "CLASS" & i, "Speed", CStr(Class(i).Stat(Stats.SPEED)))
        Call PutVar(FileName, "CLASS" & i, "MAGI", CStr(Class(i).Stat(Stats.Magic)))
        DoEvents
    Next
    
End Sub

Function CheckClasses() As Boolean
Dim FileName As String

    FileName = "\data\classes.ini"

    If Not FileExist(FileName) Then
        CreateClassesINI
        CheckClasses = True
    End If

End Function

Public Sub ClearClasses()
Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
        DoEvents
    Next
    
End Sub

' ***********
' ** Items **
' ***********

Public Sub SaveItems()
Dim i As Long
    
    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
        DoEvents
    Next
    
End Sub

Public Sub SaveItem(ByVal ItemNum As Long)
Dim FileName As String
Dim F  As Long
    
    FileName = App.Path & "\Data\items\item" & ItemNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Item(ItemNum)
    Close #F
    
End Sub

Public Sub LoadItems()
Dim FileName As String
Dim i As Long
Dim F As Long

    Call CheckItems
    
    For i = 1 To MAX_ITEMS
        FileName = App.Path & "\Data\Items\Item" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Item(i)
        Close #F
        DoEvents
    Next
    
End Sub

Public Sub CheckItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If
        DoEvents
    Next
    
End Sub

Public Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Durability = -1
End Sub

Public Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
        DoEvents
    Next
    
End Sub

' ***********
' ** Shops **
' ***********

Public Sub SaveShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
        DoEvents
    Next
    
End Sub

Public Sub SaveShop(ByVal ShopNum As Long)
    Dim FileName As String
    Dim F As Long

    FileName = App.Path & "\Data\shops\shop" & ShopNum & ".dat"

    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Shop(ShopNum)
    Close #F
    
End Sub

Public Sub LoadShops()
    Dim FileName As String
    Dim i As Long
    Dim F As Long

    Call CheckShops

    For i = 1 To MAX_SHOPS
        FileName = App.Path & "\Data\shops\shop" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Shop(i)
        Close #F
        DoEvents
    Next
    
End Sub

Public Sub CheckShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If
        DoEvents
    Next
    
End Sub

Public Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
    Shop(Index).JoinSay = vbNullString
    Shop(Index).LeaveSay = vbNullString
End Sub

Public Sub ClearShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
        DoEvents
    Next
    
End Sub

' ************
' ** Spells **
' ************

Public Sub SaveSpell(ByVal SpellNum As Long)
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\Data\spells\spells" & SpellNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
       Put #F, , Spell(SpellNum)
    Close #F
    
End Sub

Public Sub SaveSpells()
    Dim i As Long

    Call SetStatus("Saving spells... ")
    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
        DoEvents
    Next
    
End Sub

Public Sub LoadSpells()
    Dim FileName As String
    Dim i As Long
    Dim F As Long

    Call CheckSpells

    For i = 1 To MAX_SPELLS
        FileName = App.Path & "\Data\spells\spells" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Spell(i)
        Close #F
        DoEvents
    Next
    
End Sub

Public Sub CheckSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If
        DoEvents
    Next
    
End Sub

Public Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).Timer = 1000
End Sub

Public Sub ClearSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
        DoEvents
    Next
    
End Sub

' **********
' * GUILDS *
' **********

Public Sub SaveGuild(ByVal GuildNum As Long)
Dim FileName As String
Dim F As Long
Dim LoopI As Long

    FileName = App.Path & "\Data\guilds\guild" & GuildNum & ".dat"
    F = FreeFile
    
    With Guild(GuildNum)
        Open FileName For Binary As #F
            Put #F, , Len(.Name)
            Put #F, , .Name
            Put #F, , .TotalMembers
            Put #F, , UBound(.Member_Account)
            For LoopI = 0 To UBound(.Member_Account)
                Put #F, , .Member_Account(LoopI)
            Next
            For LoopI = 0 To UBound(.Member_Account)
                Put #F, , .Member_CharNum(LoopI)
            Next
        Close #F
    End With
    
End Sub

Public Sub SaveGuilds()
Dim LoopI As Long

    SetStatus "Saving guilds... "
    
    For LoopI = 1 To MAX_GUILDS
        SaveGuild LoopI
        DoEvents
    Next
    
End Sub

Public Sub LoadGuilds()
Dim FileName As String
Dim LoopI As Long
Dim LoopI2 As Long
Dim F As Long
Dim TempUBound As Long
Dim StringLen As Long

    CheckGuilds
    
    For LoopI = 1 To MAX_GUILDS
        FileName = App.Path & "\Data\guilds\guild" & LoopI & ".dat"
        F = FreeFile
        
        With Guild(LoopI)
            Open FileName For Binary As #F
                Get #F, , StringLen
                .Name = Space$(StringLen)
                Get #F, , .Name
                .Name = Trim$(.Name)
                Get #F, , .TotalMembers
                Get #F, , TempUBound
                ReDim .Member_Account(0 To TempUBound)
                ReDim .Member_CharNum(0 To TempUBound)
                For LoopI2 = 0 To TempUBound
                    Get #F, , .Member_Account(LoopI2)
                Next
                For LoopI2 = 0 To TempUBound
                    Get #F, , .Member_CharNum(LoopI2)
                Next
            Close #F
        End With
        DoEvents
    Next
    
End Sub

Public Sub CheckGuilds()
Dim LoopI As Long

    For LoopI = 1 To MAX_GUILDS
        If Not FileExist("\Data\guilds\guild" & LoopI & ".dat") Then SaveGuild LoopI
        DoEvents
    Next
    
End Sub

Public Sub ClearGuilds()
Dim LoopI As Long

    For LoopI = 1 To MAX_GUILDS
        ClearGuild LoopI
        DoEvents
    Next
    
End Sub

Public Sub ClearGuild(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Guild(Index)), LenB(Guild(Index))
    ReDim Guild(Index).Member_Account(0 To 0)
    ReDim Guild(Index).Member_CharNum(0 To 0)
End Sub

' **********
' * SIGNS  *
' **********

Public Sub SaveSign(ByVal SignNum As Long)
Dim FileName As String
Dim F As Long
Dim LoopI As Long

    FileName = App.Path & "\Data\signs\sign" & SignNum & ".dat"
    F = FreeFile
    
    With Sign(SignNum)
        Open FileName For Binary As #F
            Put #F, , .Name
            Put #F, , UBound(.Section)
            For LoopI = 0 To UBound(.Section)
                Put #F, , Len(Trim$(.Section(LoopI)))
                Put #F, , Trim$(.Section(LoopI))
            Next
        Close #F
    End With
    
End Sub

Public Sub SaveSigns()
Dim LoopI As Long

    SetStatus "Saving signs... "
    
    For LoopI = 1 To MAX_SIGNS
        SaveSign LoopI
        DoEvents
    Next
    
End Sub

Public Sub LoadSigns()
Dim FileName As String
Dim LoopI As Long
Dim LoopI2 As Long
Dim F As Long
Dim TempUBound As Long
Dim StringLen As Long

    CheckSigns
    
    For LoopI = 1 To MAX_SIGNS
        FileName = App.Path & "\Data\signs\sign" & LoopI & ".dat"
        F = FreeFile
        
        With Sign(LoopI)
            Open FileName For Binary As #F
                Get #F, , .Name
                Get #F, , TempUBound
                ReDim .Section(0 To TempUBound)
                For LoopI2 = 0 To TempUBound
                    Get #F, , StringLen
                    .Section(LoopI2) = Space$(StringLen)
                    Get #F, , .Section(LoopI2)
                    .Section(LoopI2) = Trim$(.Section(LoopI2))
                Next
            Close #F
        End With
        DoEvents
    Next
    
End Sub

Public Sub CheckSigns()
Dim LoopI As Long

    For LoopI = 1 To MAX_SIGNS
        If Not FileExist("\Data\signs\sign" & LoopI & ".dat") Then SaveSign LoopI
        DoEvents
    Next
    
End Sub

Public Sub ClearSigns()
Dim LoopI As Long

    For LoopI = 1 To MAX_SIGNS
        ClearSign LoopI
        DoEvents
    Next
    
End Sub

Public Sub ClearSign(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Sign(Index)), LenB(Sign(Index))
    ReDim Sign(Index).Section(0 To 0)
End Sub

' **********
' * ANIMS  *
' **********

Public Sub SaveAnim(ByVal AnimNum As Long)
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\Data\anims\anim" & AnimNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Animation(AnimNum)
    Close #F
    
End Sub

Public Sub SaveAnims()
Dim LoopI As Long

    SetStatus "Saving anims... "
    
    For LoopI = 1 To MAX_SIGNS
        SaveAnim LoopI
        DoEvents
    Next
    
End Sub

Public Sub LoadAnims()
Dim FileName As String
Dim LoopI As Long
Dim F As Long

    CheckAnims
    
    For LoopI = 1 To MAX_ANIMS
        FileName = App.Path & "\Data\anims\anim" & LoopI & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Animation(LoopI)
        Close #F
        DoEvents
    Next
    
End Sub

Public Sub CheckAnims()
Dim LoopI As Long

    For LoopI = 1 To MAX_ANIMS
        If Not FileExist("\Data\anims\anim" & LoopI & ".dat") Then SaveAnim LoopI
        DoEvents
    Next
    
End Sub

Public Sub ClearAnims()
Dim LoopI As Long

    For LoopI = 1 To MAX_ANIMS
        ClearAnim LoopI
        DoEvents
    Next
    
End Sub

Public Sub ClearAnim(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Animation(Index)), LenB(Animation(Index))
    Animation(Index).Delay = 100
    Animation(Index).Width = 32
    Animation(Index).Height = 32
End Sub

' **********
' ** NPCs **
' **********

Public Sub SaveNpcs()
Dim i As Long
    
    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
        DoEvents
    Next
    
End Sub

Public Sub SaveNpc(ByVal NpcNum As Long)
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\Data\npcs\npc" & NpcNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Npc(NpcNum)
    Close #F
    
End Sub

Public Sub LoadNpcs()
Dim FileName As String
Dim i As Integer
Dim F As Long

    Call CheckNpcs

    For i = 1 To MAX_NPCS
        FileName = App.Path & "\Data\npcs\npc" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Npc(i)
        Close #F
        DoEvents
    Next
    
End Sub

Public Sub CheckNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If
        DoEvents
    Next
    
End Sub

Public Sub ClearNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).Name = vbNullString
    Npc(Index).AttackSay = vbNullString
End Sub

Public Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
        DoEvents
    Next
    
End Sub

' **********
' ** Maps **
' **********

Public Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim F As Long
Dim LoopI As Long

    FileName = App.Path & "\Data\maps\map" & MapNum & ".dat"
    
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Map(MapNum)
        Put #F, , UBound(MapSpawn(MapNum).Npc)
        For LoopI = 1 To UBound(MapSpawn(MapNum).Npc)
            Put #F, , MapSpawn(MapNum).Npc(LoopI).Num
            Put #F, , MapSpawn(MapNum).Npc(LoopI).X
            Put #F, , MapSpawn(MapNum).Npc(LoopI).Y
        Next
    Close #F
    
End Sub

Public Sub SaveMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
        DoEvents
    Next
    
End Sub

Public Sub LoadMaps()
Dim FileName As String
Dim i As Long
Dim F As Long
Dim TempBound As Long
Dim LoopI As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        FileName = App.Path & "\Data\maps\map" & i & ".dat"
        
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Map(i)
            Get #F, , TempBound
            ReDim MapSpawn(i).Npc(1 To TempBound)
            ReDim MapNpc(i).MapNpc(1 To TempBound)
            For LoopI = 1 To TempBound
                Get #F, , MapSpawn(i).Npc(LoopI).Num
                Get #F, , MapSpawn(i).Npc(LoopI).X
                Get #F, , MapSpawn(i).Npc(LoopI).Y
            Next
        Close #F
        
        For F = 1 To UBound(MapSpawn(i).Npc)
            If MapSpawn(i).Npc(F).Num = 0 Then
                MapSpawn(i).Npc(F).X = -1
                MapSpawn(i).Npc(F).Y = -1
            End If
        Next
        
        DoEvents
    Next
    
End Sub

Public Sub CheckMaps()
Dim i As Long
        
    For i = 1 To MAX_MAPS
        
        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If
        DoEvents
    Next
    
End Sub

Public Sub ClearTempTile()
Dim i As Long
Dim Y As Long
Dim X As Long

    For i = 1 To MAX_MAPS
        TempTile(i).DoorTimer = 0
        
        For X = 0 To MAX_MAPX
            For Y = 0 To MAX_MAPY
                TempTile(i).DoorOpen(X, Y) = NO
                DoEvents
            Next
            DoEvents
        Next
        DoEvents
    Next
End Sub

Public Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, Index)), LenB(MapItem(MapNum, Index)))
End Sub

Public Sub ClearMapItems()
Dim X As Long
Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, Y)
            DoEvents
        Next
        DoEvents
    Next
    
End Sub

Public Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).MapNpc(Index)), LenB(MapNpc(MapNum).MapNpc(Index)))
End Sub

Public Sub ClearMapNpcs()
Dim X As Long
Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To UBound(MapSpawn(Y).Npc)
            Call ClearMapNpc(X, Y)
            DoEvents
        Next
        DoEvents
    Next
    
End Sub

Public Sub ClearMap(ByVal MapNum As Long)
Dim i As Long

    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Call ZeroMemory(ByVal VarPtr(MapSpawn(MapNum)), LenB(MapSpawn(MapNum)))
    
    ReDim MapSpawn(MapNum).Npc(1 To 10)
    ReDim MapNpc(MapNum).MapNpc(1 To 10)
    
    MapSpawn(MapNum).Npc(1).X = -1
    MapSpawn(MapNum).Npc(1).Y = -1
    
    Map(MapNum).Name = vbNullString
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = False
    
    ' Reset the map cache array for this map.
    MapCache(MapNum) = vbNullString
    
End Sub

Public Sub ClearMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
        DoEvents
    Next
    
End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            GetClassMaxVital = (1 + Int(Class(ClassNum).Stat(Stats.Strength) * 0.5) + Class(ClassNum).Stat(Stats.Strength)) * 2
        Case MP
            GetClassMaxVital = (1 + Int(Class(ClassNum).Stat(Stats.Magic) * 0.5) + Class(ClassNum).Stat(Stats.Magic)) * 2
        Case SP
            GetClassMaxVital = (1 + Int(Class(ClassNum).Stat(Stats.SPEED) * 0.5) + Class(ClassNum).Stat(Stats.SPEED)) * 2
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

Public Sub Load_BanTable()
Dim FileName As String
Dim LoopI As Long

    FileName = App.Path & "\data\bans.ini"
    
    frmServer.lstAccountBans.Clear
    frmServer.lstIPBans.Clear
    
    If Val(GetVar(FileName, "IP", "Total")) > 0 Then
        For LoopI = 1 To Val(GetVar(FileName, "IP", "Total"))
            frmServer.lstIPBans.AddItem LoopI & ") " & GetVar(FileName, "IP", "IP" & LoopI)
        Next
    End If
    
    If Val(GetVar(FileName, "ACCOUNT", "Total")) > 0 Then
        For LoopI = 1 To Val(GetVar(FileName, "ACCOUNT", "Total"))
            frmServer.lstAccountBans.AddItem LoopI & ") " & GetVar(FileName, "ACCOUNT", "Account" & LoopI)
        Next
    End If
    
    frmServer.lblIPBans.Caption = Val(GetVar(FileName, "IP", "Total")) & " IP Bans"
    frmServer.lblAccountBans.Caption = Val(GetVar(FileName, "ACCOUNT", "Total")) & " Account Bans"
    
End Sub

Public Sub Load_GameOptions()
Dim FileName As String
Dim i As Byte
Dim SplitString() As String

    FileName = App.Path & "\data\config.ini"
    
    ' set config password
    CONFIG_PASSWORD = GetVar(FileName, "SECURITY", "Password")
    
    ' get the IP address retrieval method
    IP_Source = GetVar(FileName, "OPTIONS", "IP_Method")
    
    ' force the method to either be com or org
    If IP_Source <> "com" Then
        If IP_Source <> "org" Then
            IP_Source = "com"
            PutVar FileName, "OPTIONS", "IP_Method", "com"
        End If
    End If
    
    ' for the client
    GAME_NAME = GetVar(FileName, "OPTIONS", "Game_Name")
    GAME_WEBSITE = GetVar(FileName, "OPTIONS", "Website")
    SPRITE_OFFSET = Val(GetVar(FileName, "OPTIONS", "Sprite_Offset"))
    TOTAL_ANIMFRAMES = Val(GetVar(FileName, "ANIMATION", "Total_AnimFrames"))
    CONFIG_STANDFRAME = Val(GetVar(FileName, "ANIMATION", "StandFrame"))
    
    ' max values
    MAX_PLAYERS = CLng(GetVar(FileName, "MAX_VALUES", "Players"))
    MAX_ITEMS = CLng(GetVar(FileName, "MAX_VALUES", "Items"))
    MAX_NPCS = CLng(GetVar(FileName, "MAX_VALUES", "Npcs"))
    MAX_SHOPS = CLng(GetVar(FileName, "MAX_VALUES", "Shops"))
    MAX_SPELLS = CLng(GetVar(FileName, "MAX_VALUES", "Spells"))
    MAX_MAPS = CLng(GetVar(FileName, "MAX_VALUES", "Maps"))
    MAX_SIGNS = CLng(GetVar(FileName, "MAX_VALUES", "Signs"))
    MAX_GUILDS = CLng(GetVar(FileName, "MAX_VALUES", "Guilds"))
    MAX_ANIMS = CLng(GetVar(FileName, "MAX_VALUES", "Anims"))
    
    If frmServer.scrlLevelLimit.Value > MAX_LEVELS Then frmServer.scrlLevelLimit.Value = MAX_LEVELS
    frmServer.scrlLevelLimit.Max = MAX_LEVELS
    
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
    ReDim Item(1 To MAX_ITEMS) As ItemRec
    ReDim Npc(1 To MAX_NPCS) As NpcRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim MapNpc(1 To MAX_MAPS)
    ReDim MapSpawn(1 To MAX_MAPS)
    ReDim MapCache(1 To MAX_MAPS) As String
    ReDim TempTile(1 To MAX_MAPS) As TempTileRec
    ReDim PlayersOnMap(1 To MAX_MAPS) As Boolean
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim Sign(1 To MAX_SIGNS) As SignRec
    ReDim Guild(1 To MAX_GUILDS) As GuildRec
    ReDim Animation(1 To MAX_ANIMS) As AnimationRec
    
    ReDim StatBonus(1 To MAX_PLAYERS, 1 To Stats.Stat_Count - 1) As Long
    ReDim VitalBonus(1 To MAX_PLAYERS, 1 To Vitals.Vital_Count - 1) As Long
    
    Guild_Creation_Item = CLng(GetVar(FileName, "GUILD_CONFIG", "ItemNum"))
    Guild_Creation_Cost = CLng(GetVar(FileName, "GUILD_CONFIG", "Cost"))
    
    If Guild_Creation_Item < 0 Then Guild_Creation_Item = 0
    If Guild_Creation_Item > MAX_ITEMS Then Guild_Creation_Item = 0
    If Guild_Creation_Cost < 1 Then Guild_Creation_Cost = 0: Guild_Creation_Item = 0
    
    UsersOnline_Start
    
    If GetVar(FileName, "ANIMATION", "WalkFrames") = vbNullString Then
        MsgBox "You need to make sure you have walk frames specified in the Data\config.ini!", vbOKOnly + vbCritical, "Error"
        End
    End If
    
    If Not InStr(1, GetVar(FileName, "ANIMATION", "WalkFrames"), ",", vbTextCompare) > 0 Then
        ReDim WalkFrame(1 To 1)
        TOTAL_WALKFRAMES = 1
        WalkFrame(1) = Val(GetVar(FileName, "ANIMATION", "WalkFrames"))
        GoTo Skip1
    End If
    
    SplitString = Split(GetVar(FileName, "ANIMATION", "WalkFrames"), ",", , vbTextCompare)
    
    ReDim WalkFrame(1 To UBound(SplitString) + 1)
    TOTAL_WALKFRAMES = UBound(WalkFrame)
    
    For i = 0 To UBound(SplitString)
        WalkFrame(i + 1) = Val(SplitString(i)) - 1
    Next
    
Skip1:
    
    If GetVar(FileName, "ANIMATION", "AttackFrames") = vbNullString Then GoTo Skip2
    
    If Not InStr(1, GetVar(FileName, "ANIMATION", "AttackFrames"), ",", vbTextCompare) > 0 Then
        ReDim AttackFrame(1 To 1)
        TOTAL_ATTACKFRAMES = 1
        AttackFrame(1) = Val(GetVar(FileName, "ANIMATION", "AttackFrames"))
        GoTo Skip2
    End If
    
    SplitString = Split(GetVar(FileName, "ANIMATION", "AttackFrames"), ",", , vbTextCompare)
    
    ReDim AttackFrame(1 To UBound(SplitString) + 1)
    TOTAL_ATTACKFRAMES = UBound(AttackFrame)
    
    For i = 0 To UBound(SplitString)
        AttackFrame(i + 1) = Val(SplitString(i)) - 1
    Next
    
Skip2:
    
    Direction_Anim(E_Direction.Up_) = Val(GetVar(FileName, "ANIMATION", "Anim_Up"))
    Direction_Anim(E_Direction.Down_) = Val(GetVar(FileName, "ANIMATION", "Anim_Down"))
    Direction_Anim(E_Direction.Left_) = Val(GetVar(FileName, "ANIMATION", "Anim_Left"))
    Direction_Anim(E_Direction.Right_) = Val(GetVar(FileName, "ANIMATION", "Anim_Right"))
    
    WALKANIM_SPEED = Val(GetVar(FileName, "ANIMATION", "WalkAnim_Speed"))
    
    ' for the server
    GAME_PORT = Val(GetVar(FileName, "SETUP", "Port"))
    
End Sub

Public Function Encryption(CodeKey As String, DataIn As String) As String
Dim lonDataPtr As Long
Dim strDataOut As String
Dim intXOrValue1 As Integer
Dim intXOrValue2 As Integer

    For lonDataPtr = 1 To Len(DataIn)
    
        intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        strDataOut = strDataOut + Chr$(intXOrValue1 Xor intXOrValue2)
    
    Next
    
    Encryption = strDataOut
   
End Function
