Attribute VB_Name = "modDatabase"
Option Explicit

Public Function ReadIniValue(INIpath As String, Key As String, Variable As String) As String
Dim NF As Integer
Dim Temp As String
Dim LcaseTemp As String
Dim ReadyToRead As Boolean
    
AssignVariables:
        NF = FreeFile
        ReadIniValue = vbNullString
        Key = "[" & LCase$(Key) & "]"
        Variable = LCase$(Variable)
    
EnsureFileExists:
    Open INIpath For Binary As NF
    Close NF
    SetAttr INIpath, vbArchive
    
LoadFile:
    Open INIpath For Input As NF
    While Not EOF(NF)
    Line Input #NF, Temp
    LcaseTemp = LCase$(Temp)
    If InStr(LcaseTemp, "[") <> 0 Then ReadyToRead = False
    If LcaseTemp = Key Then ReadyToRead = True
    If InStr(LcaseTemp, "[") = 0 And ReadyToRead = True Then
        If InStr(LcaseTemp, Variable & "=") = 1 Then
            ReadIniValue = Mid$(Temp, 1 + Len(Variable & "="))
            Close NF: Exit Function
            End If
        End If
    Wend
    Close NF
End Function

Public Function WriteIniValue(INIpath As String, PutKey As String, PutVariable As String, PutValue As String)
Dim Temp As String
Dim LcaseTemp As String
Dim ReadKey As String
Dim ReadVariable As String
Dim LOKEY As Integer
Dim HIKEY As Integer
Dim KeyLen As Integer
Dim VAR As Integer
Dim VARENDOFLINE As Integer
Dim NF As Integer
Dim X As Integer

AssignVariables:
    NF = FreeFile
    ReadKey = vbCrLf & "[" & LCase$(PutKey) & "]" & Chr$(13)
    KeyLen = Len(ReadKey)
    ReadVariable = Chr$(10) & LCase$(PutVariable) & "="
        
EnsureFileExists:
    Open INIpath For Binary As NF
    Close NF
    SetAttr INIpath, vbArchive
    
LoadFile:
    Open INIpath For Input As NF
    Temp = Input$(LOF(NF), NF)
    Temp = vbCrLf & Temp & "[]"
    Close NF
    LcaseTemp = LCase$(Temp)
    
LogicMenu:
    LOKEY = InStr(LcaseTemp, ReadKey)
    If LOKEY = 0 Then GoTo AddKey:
    HIKEY = InStr(LOKEY + KeyLen, LcaseTemp, "[")
    VAR = InStr(LOKEY, LcaseTemp, ReadVariable)
    If VAR > HIKEY Or VAR < LOKEY Then GoTo AddVariable:
    GoTo RenewVariable:
    
AddKey:
        Temp = Left$(Temp, Len(Temp) - 2)
        Temp = Temp & vbCrLf & vbCrLf & "[" & PutKey & "]" & vbCrLf & PutVariable & "=" & PutValue
        GoTo TrimFinalString:
        
AddVariable:
        Temp = Left$(Temp, Len(Temp) - 2)
        Temp = Left$(Temp, LOKEY + KeyLen) & PutVariable & "=" & PutValue & vbCrLf & Mid$(Temp, LOKEY + KeyLen + 1)
        GoTo TrimFinalString:
        
RenewVariable:
        Temp = Left$(Temp, Len(Temp) - 2)
        VARENDOFLINE = InStr(VAR, Temp, Chr$(13))
        Temp = Left$(Temp, VAR) & PutVariable & "=" & PutValue & Mid$(Temp, VARENDOFLINE)
        GoTo TrimFinalString:

TrimFinalString:
        Temp = Mid$(Temp, 2)
        Do Until InStr(Temp, vbCrLf & vbCrLf & vbCrLf) = 0
        Temp = Replace(Temp, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
        Loop
    
        Do Until Right$(Temp, 1) > Chr$(13)
        Temp = Left$(Temp, Len(Temp) - 1)
        Loop
    
        Do Until Left$(Temp, 1) > Chr$(13)
        Temp = Mid$(Temp, 2)
        Loop
    
OutputAmendedINIFile:
        Open INIpath For Output As NF
        Print #NF, Temp
        Close NF
    
End Function

Function GetVar(File As String, Header As String, VAR As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = vbNullString
  
    sSpaces = Space$(5000)
  
    GetPrivateProfileString Header, VAR, szReturn, sSpaces, Len(sSpaces), File
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, Header As String, VAR As String, Value As String)
    WritePrivateProfileString Header, VAR, Value, File
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean

   If RAW = False Then
       If Dir$(App.Path & "\" & FileName) = vbNullString Then
           FileExist = False
           Exit Function
       Else
           FileExist = True
           Exit Function
       End If
   Else
       If Dir$(FileName) = vbNullString Then
           FileExist = False
           Exit Function
       Else
           FileExist = True
       End If
   End If
End Function

'**************************************
'** Player                           **
'**************************************
Sub SavePlayer(ByVal Index As Long)
Dim FileName As String
Dim nFileNum As Integer

    FileName = AccountPath & "\" & Trim$(Player(Index).Login) & ".acc"
    
    nFileNum = FreeFile
    Open FileName For Binary As #nFileNum
        Put #nFileNum, (NAME_LENGTH * 2) + (LenB(Player(Index).Char) * (Current_CharNum(Index) - 1)), Player(Index).Char
    Close #nFileNum
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
Dim FileName As String
Dim nFileNum As Integer

    ClearChar Index
    
    FileName = AccountPath & "\" & Trim$(Name) & ".acc" 'Cool file extention
    
    nFileNum = FreeFile
    Open FileName For Binary As #nFileNum
        Get #nFileNum, (NAME_LENGTH * 2) + (LenB(Player(Index).Char) * (Current_CharNum(Index) - 1)), Player(Index).Char
    Close #nFileNum
End Sub

Function AccountExist(ByVal Name As String) As Boolean
Dim FileName As String

    FileName = AccountPath & "\" & Name & ".acc"
    
    If FileExist(FileName, True) Then
        AccountExist = True
    Else
        AccountExist = False
    End If
End Function

Function CharExist(ByVal Index As Long) As Boolean
    CharExist = False
    If LenB(Current_Name(Index)) > 0 Then CharExist = True
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim FileName As String
Dim RightPassword As String * NAME_LENGTH
Dim nFileNum As Integer

   PasswordOK = False
   
   If AccountExist(Name) Then
       FileName = AccountPath & "\" & Trim$(Name) & ".acc"
       nFileNum = FreeFile
       Open FileName For Binary As #nFileNum
       
       Get #nFileNum, 20, RightPassword
       
       Close #nFileNum
       
       If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
           PasswordOK = True
       End If
   End If
   
End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
Dim i As Long
Dim FileName As String
Dim nFileNum As Integer

    Player(Index).Login = Name
    Player(Index).Password = Password
    
    ClearChar Index

    FileName = AccountPath & "\" & Trim$(Player(Index).Login) & ".acc"
    
    nFileNum = FreeFile
    Open FileName For Binary As #nFileNum
        Put #nFileNum, , Player(Index).Login
        Put #nFileNum, , Player(Index).Password
        For i = 1 To MAX_CHARS
            Put #nFileNum, (NAME_LENGTH * 2) + (LenB(Player(Index).Char) * (i - 1)), Player(Index).Char
        Next
    Close #nFileNum
End Sub

Sub PlayerAddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long, ByVal SpriteNum As Long)
Dim f As Long, n As Long
        
    With Player(Index).Char
        .Name = Name
        .Sex = Sex
        .Class = ClassNum
        
        .Sprite = SpriteNum
        
        .Level = 1
        
        .Exp = 0
        .Access = 0
        .PK = 0
        .Points = 0
        .Guild = 0
        .GuildRank = 0
        .GuildName = vbNullString
    
        For n = 1 To Stats.Stat_Count
            .Stat(n) = Class(ClassNum).Stat(n)
        Next
        
        .Position.Map = START_MAP
        .Position.X = START_X
        .Position.Y = START_Y
            
        ' Set their initial bound spot
        .Bound.Map = START_MAP
        .Bound.X = START_X
        .Bound.Y = START_Y
        
        For n = 1 To Vitals.Vital_Count
            .Vital(n) = Current_MaxVital(Index, n)
        Next
    End With
    
    ' Append name to file
    f = FreeFile
    Open App.Path & "\Data\accounts\charlist.txt" For Append As #f
        Print #f, Name
    Close #f
    
    SavePlayer Index
End Sub

Function FindChar(ByVal Name As String) As Boolean
Dim f As Long
Dim s As String

    FindChar = False
    
    f = FreeFile
    Open AccountPath & "\charlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            
            If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
                FindChar = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
End Function

Function FindGuildName(ByVal GuildName As String) As Boolean
Dim f As Long
Dim s As String

    FindGuildName = False
    
    f = FreeFile
    Open GuildPath & "\GuildName.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            
            If Trim$(LCase$(s)) = Trim$(LCase$(GuildName)) Then
                FindGuildName = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
End Function

Sub DeleteGuildName(ByVal GuildName As String)
Dim f1 As Long, f2 As Long
Dim s As String

    FileCopy GuildPath & "\GuildName.txt", GuildPath & "\GuildNameTemp.txt"
    
    ' Destroy name from charlist
    f1 = FreeFile
    Open GuildPath & "\GuildNameTemp.txt" For Input As #f1
    
    f2 = FreeFile
    Open GuildPath & "\GuildName.txt" For Output As #f2
        
    Do While Not EOF(f1)
        Input #f1, s
        If Trim$(LCase$(s)) <> Trim$(LCase$(GuildName)) Then
            Print #f2, s
        End If
    Loop
    
    Close #f1
    Close #f2
    
    Kill (GuildPath & "\GuildNameTemp.txt")
End Sub

Function FindGuildAbbreviation(ByVal GuildAbbreviation As String) As Boolean
Dim f As Long
Dim s As String

    FindGuildAbbreviation = False
    
    f = FreeFile
    Open GuildPath & "\GuildAbbreviation.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            
            If Trim$(LCase$(s)) = Trim$(LCase$(GuildAbbreviation)) Then
                FindGuildAbbreviation = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
End Function

Sub DeleteGuildAbbreviation(ByVal GuildAbbreviation As String)
Dim f1 As Long, f2 As Long
Dim s As String

    FileCopy GuildPath & "\GuildAbbreviation.txt", GuildPath & "\GuildAbbreviationTemp.txt"
    
    ' Destroy name from charlist
    f1 = FreeFile
    Open GuildPath & "\GuildAbbreviationTemp.txt" For Input As #f1
    f2 = FreeFile
    Open GuildPath & "\GuildAbbreviation.txt" For Output As #f2
        
    Do While Not EOF(f1)
        Input #f1, s
        If Trim$(LCase$(s)) <> Trim$(LCase$(GuildAbbreviation)) Then
            Print #f2, s
        End If
    Loop
    
    Close #f1
    Close #f2
    
    Kill (GuildPath & "\GuildAbbreviationTemp.txt")
End Sub

Sub SaveAllPlayersOnline()
Dim i As Long

    For i = 1 To OnlinePlayersCount
        SavePlayer OnlinePlayers(i)
    Next
End Sub

Sub LoadClasses()
Dim FileName As String
Dim i As Long
Dim n As Long

    CheckClasses
    
    FileName = App.Path & "\Data\Core Files\classes.ini"
    
    ' Loads the class data
    ' Makes sure to load only the first 32
    MAX_CLASSES = Clamp(Val(GetVar(FileName, "MAX", "ClassCount")) - 1, 0, 31)
    ReDim Class(0 To MAX_CLASSES)
    
    For i = 0 To MAX_CLASSES
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).MaleSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        Class(i).FemaleSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        ' Vitals
        For n = 1 To Vitals.Vital_Count
            Class(i).Vital(n) = Val(GetVar(FileName, "CLASS" & i, VitalName(n)))
        Next
        ' Stats
        For n = 1 To Stats.Stat_Count
            Class(i).Stat(n) = Val(GetVar(FileName, "CLASS" & i, StatAbbreviation(n)))
        Next
        ' Other
        Class(i).BaseDodge = Val(GetVar(FileName, "CLASS" & i, "DODGE"))
        Class(i).BaseCrit = Val(GetVar(FileName, "CLASS" & i, "CRIT"))
        Class(i).BaseBlock = Val(GetVar(FileName, "CLASS" & i, "BLOCK"))
        Class(i).Threat = Val(GetVar(FileName, "CLASS" & i, "THREAT"))
    Next
End Sub

Sub SaveClasses()
Dim FileName As String
Dim i As Long
Dim n As Long

    FileName = App.Path & "\Data\Core Files\classes.ini"
    
    For i = 0 To MAX_CLASSES
        PutVar FileName, "CLASS" & i, "Name", Trim$(Class(i).Name)
        PutVar FileName, "CLASS" & i, "MaleSprite", CStr(Class(i).MaleSprite)
        PutVar FileName, "CLASS" & i, "FemaleSprite", CStr(Class(i).FemaleSprite)
        ' Vitals
        For n = 1 To Vitals.Vital_Count
            PutVar FileName, "CLASS" & i, VitalName(n), CStr(Class(i).Vital(n))
        Next
        ' Stats
        For n = 1 To Stats.Stat_Count
            PutVar FileName, "CLASS" & i, StatAbbreviation(n), CStr(Class(i).Stat(n))
        Next
        ' Other
        PutVar FileName, "CLASS" & i, "DODGE", CStr(Class(i).BaseDodge)
        PutVar FileName, "CLASS" & i, "CRIT", CStr(Class(i).BaseCrit)
        PutVar FileName, "CLASS" & i, "BLOCK", CStr(Class(i).BaseBlock)
        PutVar FileName, "CLASS" & i, "THREAT", CStr(Class(i).Threat)
    Next
End Sub

Sub CheckClasses()
    If Not FileExist("\Data\Core Files\classes.ini") Then
        SaveClasses
    End If
End Sub

Sub SaveItems()
Dim i As Long
    
    For i = 1 To MAX_ITEMS
        If Not FileExist(ItemPath & "\item" & i & ".kde", True) Then
            SaveItem (i)
        End If
    Next
End Sub

Sub SaveItem(ByVal ItemNum As Long)
Dim FileName As String
Dim f  As Long

    FileName = ItemPath & "\item" & ItemNum & ".kde"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Item(ItemNum)
    Close #f
End Sub

Sub LoadItems()
Dim FileName As String
Dim i As Long
Dim f As Long

    CheckItems
    
    For i = 1 To MAX_ITEMS
        FileName = ItemPath & "\Item" & i & ".kde"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Item(i)
        Close #f
    Next
End Sub

Sub CheckItems()
    SaveItems
End Sub

Sub SaveShops()
Dim i As Long
    
    For i = 1 To MAX_SHOPS
        If Not FileExist(ShopPath & "\shop" & i & ".kde", True) Then
            SaveShop (i)
        End If
    Next
End Sub

Sub SaveShop(ByVal ShopNum As Long)
Dim FileName As String
Dim f As Long

    FileName = ShopPath & "\shop" & ShopNum & ".kde"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Shop(ShopNum)
    Close #f
End Sub

Sub LoadShops()
Dim FileName As String
Dim i As Long, f As Long

    CheckShops
    
    For i = 1 To MAX_SHOPS
        FileName = ShopPath & "\shop" & i & ".kde"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Shop(i)
        Close #f
    Next
End Sub

Sub CheckShops()

    SaveShops
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
Dim FileName As String
Dim f As Long

    FileName = SpellPath & "\spells" & SpellNum & ".kde"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Spell(SpellNum)
    Close #f
End Sub

Sub SaveSpells()
Dim i As Long
    
    For i = 1 To MAX_SPELLS
        If Not FileExist(SpellPath & "\spells" & i & ".kde", True) Then
            SaveSpell (i)
        End If
    Next
End Sub

Sub LoadSpells()
Dim FileName As String
Dim i As Long
Dim f As Long

    CheckSpells
    
    For i = 1 To MAX_SPELLS
        FileName = SpellPath & "\spells" & i & ".kde"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Spell(i)
        Close #f
    Next
End Sub

Sub CheckSpells()
    SaveSpells
End Sub

Sub SaveEmoticon(ByVal emoNum As Long)
Dim FileName As String
Dim f As Long

    FileName = EmoticonPath & "\emoticon" & emoNum & ".kde"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Emoticons(emoNum)
    Close #f
End Sub

Sub SaveEmoticons()
Dim i As Long

    For i = 1 To MAX_EMOTICONS
        If Not FileExist(EmoticonPath & "\emoticon" & i & ".kde", True) Then
            SaveEmoticon i
        End If
    Next
End Sub

Sub LoadEmos()
Dim FileName As String
Dim i As Long
Dim f As Long

    CheckEmos
    
    For i = 1 To MAX_EMOTICONS
        FileName = EmoticonPath & "\emoticon" & i & ".kde"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Emoticons(i)
        Close #f
    Next
End Sub

Sub CheckEmos()
  SaveEmoticons
End Sub

Sub SaveNpcs()
Dim i As Long
    
    For i = 1 To MAX_NPCS
        If Not FileExist(NpcPath & "\npc" & i & ".kde", True) Then
            SaveNpc (i)
        End If
    Next
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
Dim FileName As String
Dim f As Long

    FileName = NpcPath & "\npc" & NpcNum & ".kde"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Npc(NpcNum)
    Close #f
End Sub

Sub LoadNpcs()
Dim FileName As String
Dim i As Long
Dim f As Long

    CheckNpcs
      
    For i = 1 To MAX_NPCS
        FileName = NpcPath & "\npc" & i & ".kde"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Npc(i)
        Close #f
    Next
End Sub

Sub CheckNpcs()
    SaveNpcs
End Sub

Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long
Dim i As Long
Dim X As Long
Dim Y As Long

    FileName = MapPath & "\map" & MapNum & ".dat"
    
    If FileExist(FileName, True) Then Kill FileName
        
    f = FreeFile
    Open FileName For Binary As #f
        With Map(MapNum)
            Put #f, , .Name
            Put #f, , .Revision
            Put #f, , .Moral
            Put #f, , .Up
            Put #f, , .Down
            Put #f, , .Left
            Put #f, , .Right
            Put #f, , .Music
            Put #f, , .BootMap
            Put #f, , .BootX
            Put #f, , .BootY
            Put #f, , .TileSet
            Put #f, , .MaxX
            Put #f, , .MaxY
            
            For X = 0 To .MaxX
                For Y = 0 To .MaxY
                    Put #f, , .Tile(X, Y)
                Next
            Next
            
            For i = 1 To MAX_MOBS
                Put #f, , .Mobs(i).NpcCount
                If .Mobs(i).NpcCount > 0 Then
                    For Y = 1 To .Mobs(i).NpcCount
                        Put #f, , .Mobs(i).Npc(Y)
                    Next
                End If
            Next
        End With
    Close #f
End Sub

Sub SaveMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        SaveMap (i)
    Next
End Sub

Sub LoadMaps()
Dim FileName As String
Dim f As Long
Dim i As Long
Dim X As Long
Dim Y As Long

    CheckMaps
    
    For i = 1 To MAX_MAPS
        FileName = MapPath & "\map" & i & ".dat"
        
        f = FreeFile
        Open FileName For Binary As #f
            With Map(i)
                Get #f, , .Name
                Get #f, , .Revision
                Get #f, , .Moral
                Get #f, , .Up
                Get #f, , .Down
                Get #f, , .Left
                Get #f, , .Right
                Get #f, , .Music
                Get #f, , .BootMap
                Get #f, , .BootX
                Get #f, , .BootY
                Get #f, , .TileSet
                Get #f, , .MaxX
                Get #f, , .MaxY
                
                ReDim .Tile(0 To .MaxX, 0 To .MaxY) As TileRec
                For X = 0 To .MaxX
                    For Y = 0 To .MaxY
                        Get #f, , .Tile(X, Y)
                    Next
                Next
                
                For X = 1 To MAX_MOBS
                    Get #f, , .Mobs(X).NpcCount
                    ReDim .Mobs(X).Npc(.Mobs(X).NpcCount)
                    
                    If .Mobs(X).NpcCount > 0 Then
                        For Y = 1 To .Mobs(X).NpcCount
                            Get #f, , .Mobs(X).Npc(Y)
                        Next
                    End If
                Next
            End With
        Close #f
        
        ClearTempTile i
        UpdateMapNpc i
        
        DoEvents
    Next
    
    ClearMapNpcs
End Sub

Sub CheckMaps()
Dim FileName As String
Dim i As Long
        
    For i = 1 To MAX_MAPS
        FileName = MapPath & "\map" & i & ".dat"
        
        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(FileName, True) Then
            SaveMap (i)
        End If
    Next
End Sub

Sub AddLog(ByVal Text As String, ByVal FN As String)
Dim FileName As String
Dim f As Long

    If ServerLog Then
        FileName = LogPath & "\" & FN
   
        If Not FileExist(FileName, True) Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If
   
        f = FreeFile
        Open FileName For Append As #f
            Print #f, Time & ": " & Text
        Close #f
    End If
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
Dim FileName As String, IP As String
Dim f As Long, i As Long

    FileName = App.Path & "\Data\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist(FileName, True) Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    ' Cut off last portion of ip
    IP = Current_IP(BanPlayerIndex)
            
    For i = Len(IP) To 1 Step -1
        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If
    Next
    IP = Mid$(IP, 1, i)
            
    f = FreeFile
    Open FileName For Append As #f
        Print #f, IP & "," & Current_Name(BannedByIndex)
    Close #f
    
    SendGlobalMsg "[Realm Event] " & Current_Name(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & Current_Name(BannedByIndex) & "!", ActionColor
    AddLog Current_Name(BannedByIndex) & " has banned " & Current_Name(BanPlayerIndex) & ".", ADMIN_LOG
    SendAlertMsg BanPlayerIndex, "You have been banished by " & Current_Name(BannedByIndex) & "."
End Sub

Sub DeleteName(ByVal Name As String)
Dim f1 As Long, f2 As Long
Dim s As String

    FileCopy AccountPath & "\charlist.txt", AccountPath & "\chartemp.txt"
    
    ' Destroy name from charlist
    f1 = FreeFile
    Open AccountPath & "\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open AccountPath & "\charlist.txt" For Output As #f2
        
    Do While Not EOF(f1)
        Input #f1, s
        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If
    Loop
    
    Close #f1
    Close #f2
    
    Kill (AccountPath & "\chartemp.txt")
End Sub

Sub ZeroBanList()
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\Data\banlist.txt"
    Kill (FileName)
    
    ' Make sure the file exists
    If Not FileExist(FileName) Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    f = FreeFile
    Open FileName For Append As #f
        Print #f, "0.0.0.,abc"
    Close #f
End Sub

Sub SaveGuilds()
Dim i As Long
    
    For i = 1 To MAX_GUILDS
        SaveGuild (i)
    Next
End Sub

Sub SaveGuild(ByVal GuildNum As Long)
Dim FileName As String
Dim f As Long

    FileName = GuildPath & "\guild" & GuildNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Guild(GuildNum)
    Close #f
End Sub

Sub LoadGuilds()
Dim FileName As String
Dim i As Long
Dim f As Long
    
    For i = 1 To MAX_GUILDS
        FileName = GuildPath & "\guild" & i & ".dat"
        
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Guild(i)
        Close #f
    Next
End Sub

Sub CheckGuilds()
    SaveGuilds
End Sub

'////////////////////////
'// Animation Database //
'////////////////////////
Sub CheckAnimations()
    SaveAnimations
End Sub
Sub SaveAnimations()
Dim i As Long
    
    For i = 1 To MAX_ANIMATIONS
        If Not FileExist(AnimationPath & "\Animation" & i & ".dat", True) Then
            SaveAnimation i
        End If
    Next
End Sub
Sub SaveAnimation(ByVal AnimationNum As Long)
Dim FileName As String
Dim f As Long

    FileName = AnimationPath & "\animation" & AnimationNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Animation(AnimationNum)
    Close #f
End Sub
Sub LoadAnimations()
Dim FileName As String
Dim i As Long
Dim f As Long
    
    CheckAnimations
    
    For i = 1 To MAX_ANIMATIONS
        FileName = AnimationPath & "\animation" & i & ".dat"
        
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Animation(i)
        Close #f
    Next
End Sub
Sub ClearAnimations()
Dim i As Long
    
    For i = 1 To MAX_ANIMATIONS
        ClearAnimation i
    Next
End Sub
Sub ClearAnimation(ByVal AnimationNum As Long)
    ZeroMemory ByVal VarPtr(Animation(AnimationNum)), LenB(Animation(AnimationNum))
    Animation(AnimationNum).Name = vbNullString
End Sub

Sub ClearTempTile(ByVal MapNum As Long)
Dim Y As Long
Dim X As Long

    ' set the DoorOpen()
    ReDim MapData(MapNum).TempTile.DoorOpen(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    ReDim MapData(MapNum).TempTile.DoorTimer(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            MapData(MapNum).TempTile.DoorOpen(X, Y) = False
            MapData(MapNum).TempTile.DoorTimer(X, Y) = 0
        Next
    Next
End Sub

Sub ClearTempTiles()
Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next
End Sub

Sub ClearClasses()
Dim i As Long
    ReDim Class(0 To MAX_CLASSES)
    For i = 0 To MAX_CLASSES
        ZeroMemory ByVal VarPtr(Class(i)), LenB(Class(i))
        Class(i).Name = vbNullString
    Next
End Sub
    
Sub ClearItem(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Item(Index)), LenB(Item(Index))
    Item(Index).Name = vbNullString
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        ClearItem (i)
    Next
End Sub

Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        ClearNpc (i)
    Next
End Sub
Sub ClearNpc(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Npc(Index)), LenB(Npc(Index))
    Npc(Index).Name = vbNullString
    Npc(Index).AttackSay = vbNullString
End Sub

Sub ClearMapItems()
Dim MapNum As Long, MapItemNum As Long

    For MapNum = 1 To MAX_MAPS
        For MapItemNum = 1 To MAX_MAP_ITEMS
            ClearMapItem MapNum, MapItemNum
        Next
    Next
End Sub
Sub ClearMapItem(ByVal MapNum As Long, ByVal MapItemNum As Long)
    ZeroMemory ByVal VarPtr(MapData(MapNum).MapItem(MapItemNum)), LenB(MapData(MapNum).MapItem(MapItemNum))
End Sub

Sub ClearMapNpcs()
Dim MapNum As Long, MapNpcNum As Long

    For MapNum = 1 To MAX_MAPS
        For MapNpcNum = 1 To MapData(MapNum).NpcCount
            ClearMapNpc MapNum, MapNpcNum
        Next
    Next
End Sub
Sub ClearMapNpc(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    ZeroMemory ByVal VarPtr(MapData(MapNum).MapNpc(MapNpcNum)), LenB(MapData(MapNum).MapNpc(MapNpcNum))
    Set MapData(MapNum).MapNpc(MapNpcNum).Damage = New Dictionary
End Sub

Sub ClearMap(ByVal MapNum As Long)
Dim i As Long

    ZeroMemory ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum))
    Map(MapNum).Name = vbNullString
    
    ' set the min value for maxx and maxy
    Map(MapNum).MaxX = MAX_MAPX
    Map(MapNum).MaxY = MAX_MAPY
    
    ' set the Tile()
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    
    'set the min for npcs
    For i = 1 To MAX_MOBS
        Map(MapNum).Mobs(i).NpcCount = 0
        ReDim Map(MapNum).Mobs(i).Npc(0)
    Next
    
    UpdateMapNpc MapNum
    
    ' Reset the values for if a player is on the map or not
    MapData(MapNum).MapPlayersCount = 0
End Sub

Sub ClearMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        ClearMap i
    Next
End Sub

Sub ClearShop(ByVal Index As Long)

    ZeroMemory ByVal VarPtr(Shop(Index)), LenB(Shop(Index))
    Shop(Index).Name = vbNullString
    Shop(Index).JoinSay = vbNullString
End Sub

Sub ClearShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        ClearShop i
    Next
End Sub

Sub ClearSpell(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Spell(Index)), LenB(Spell(Index))
    Spell(Index).Name = vbNullString
End Sub

Sub ClearSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        ClearSpell (i)
    Next
End Sub

Sub ClearEmos()
Dim i As Long

    For i = 1 To MAX_EMOTICONS
        Emoticons(i).Pic = 0
        Emoticons(i).Command = vbNullString
    Next
End Sub

'Dugor - 4/05/08
Sub ClearGuilds()
Dim i As Long

    For i = 1 To MAX_GUILDS
        ClearGuild (i)
    Next
End Sub
'Dugor - 4/05/08
Sub ClearGuild(ByVal Index As Long)
Dim i As Long
    
    'ZeroMemory Guild(Index), Len(Guild(Index))
    
    Guild(Index).GuildName = vbNullString
    Guild(Index).GuildAbbreviation = vbNullString
    Guild(Index).GMOTD = vbNullString
    Guild(Index).Owner = vbNullString
    For i = 1 To MAX_GUILD_RANKS
        Guild(Index).Rank(i) = vbNullString
    Next
End Sub
