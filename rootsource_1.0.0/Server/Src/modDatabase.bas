Attribute VB_Name = "modDatabase"
Option Explicit

' ********************************************
' **               rootSource               **
' ********************************************

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public ServerLog As Boolean ' Used for logging

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    If Not RAW Then
        If LenB(Dir(App.Path & "\" & FileName)) > 0 Then
            FileExist = True
         End If
    Else
        If LenB(Dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If
End Function

Public Sub AddLog(ByVal Text As String, ByVal FN As String)
Dim FileName As String
Dim F As Integer

    If ServerLog Then
        FileName = App.Path & "\logs\" & FN
    
        If Not FileExist(FileName, True) Then
            F = FreeFile
            Open FileName For Output As #F
            Close #F
        End If
    
        F = FreeFile
        Open FileName For Append As #F
            Print #F, time & ": " & Text
        Close #F
    End If
End Sub

Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = vbNullString
  
    sSpaces = Space$(5000)
  
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
Dim FileName As String
Dim IP As String
Dim F As Long
Dim i As Long

    FileName = App.Path & "\data\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
    
    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)
            
    For i = Len(IP) To 1 Step -1
        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If
    Next
    IP = Mid$(IP, 1, i)
            
    F = FreeFile
    Open FileName For Append As #F
        Print #F, IP & "," & GetPlayerName(BannedByIndex)
    Close #F
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Public Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
    Dim FileName As String
    Dim IP As String
    Dim F As Long
    Dim i As Long

    FileName = App.Path & "\data\banlist.txt"
   
    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
   
    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)
           
    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
           
    F = FreeFile
    Open FileName For Append As #F
    Print #F, IP & ", Server"
    Close #F
   
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by the Server!", White)
    Call AddLog("The Server has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by The Server!")
End Sub

' **************
' ** Accounts **
' **************
Public Function AccountExist(ByVal Name As String) As Boolean
Dim FileName As String

    FileName = "accounts\" & Trim(Name) & ".bin"
   
    If FileExist(FileName) Then
        AccountExist = True
    End If
End Function

Public Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim FileName As String
Dim RightPassword As String * NAME_LENGTH
Dim nFileNum As Integer

    If AccountExist(Name) Then
        FileName = App.Path & "\Accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open FileName For Binary As #nFileNum
       
        Get #nFileNum, NAME_LENGTH, RightPassword
       
        Close #nFileNum
       
        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
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

    Call FileCopy(App.Path & "\accounts\charlist.txt", App.Path & "\accounts\chartemp.txt")
    
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

Public Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If LenB(Trim$(Player(Index).Char(CharNum).Name)) > 0 Then
        CharExist = True
    End If
End Function

Public Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
Dim F As Long
Dim n As Long

    If LenB(Trim$(Player(Index).Char(CharNum).Name)) = 0 Then
        TempPlayer(Index).CharNum = CharNum
        
        Player(Index).Char(CharNum).Name = Name
        Player(Index).Char(CharNum).Sex = Sex
        Player(Index).Char(CharNum).Class = ClassNum
        
        If Player(Index).Char(CharNum).Sex = SEX_MALE Then
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).Sprite
        Else
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).Sprite
        End If
        
        Player(Index).Char(CharNum).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(Index).Char(CharNum).Stat(n) = Class(ClassNum).Stat(n)
        Next n
        
        Player(Index).Char(CharNum).Map = START_MAP
        Player(Index).Char(CharNum).X = START_X
        Player(Index).Char(CharNum).y = START_Y
            
        Player(Index).Char(CharNum).Vital(Vitals.HP) = GetPlayerMaxVital(Index, Vitals.HP)
        Player(Index).Char(CharNum).Vital(Vitals.MP) = GetPlayerMaxVital(Index, Vitals.MP)
        Player(Index).Char(CharNum).Vital(Vitals.SP) = GetPlayerMaxVital(Index, Vitals.SP)
                
        ' Append name to file
        F = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Append As #F
            Print #F, Name
        Close #F
        
        Call SavePlayer(Index)
            
        Exit Sub
    End If
End Sub

Public Sub DelChar(ByVal Index As Long, ByVal CharNum As Long)
    Call DeleteName(Player(Index).Char(CharNum).Name)
    Call ClearChar(Index, CharNum)
    Call SavePlayer(Index)
End Sub

Public Function FindChar(ByVal Name As String) As Boolean
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

    For i = 1 To TotalPlayersOnline
        Call SavePlayer(PlayersOnline(i))
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
   
    FileName = App.Path & "\accounts\" & Trim(Name) & ".bin"

    F = FreeFile
    Open FileName For Binary As #F
        Get #F, , Player(Index)
    Close #F
End Sub

Public Sub ClearPlayer(ByVal Index As Long)
Dim i As Long

    ' Clear the tempPlayer also
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Set TempPlayer(Index).Buffer = New clsBuffer
    
    ' Clear the player UDT
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))

    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString

    For i = 0 To MAX_CHARS
        Call ClearChar(Index, i)
    Next

    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
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

    FileName = App.Path & "\data\classes.ini"
    
    Max_Classes = 2
    
    If Not FileExist(FileName, True) Then
        File = FreeFile
    
        Open FileName For Output As File
            Print #File, "[INIT]"
            Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Public Sub LoadClasses()
Dim FileName As String
Dim i As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        FileName = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
        
    End If

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).Sprite = Val(GetVar(FileName, "CLASS" & i, "Sprite"))
        Class(i).Stat(Stats.Strength) = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class(i).Stat(Stats.Defense) = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class(i).Stat(Stats.Speed) = Val(GetVar(FileName, "CLASS" & i, "Speed"))
        Class(i).Stat(Stats.Magic) = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
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
        Call PutVar(FileName, "CLASS" & i, "Speed", CStr(Class(i).Stat(Stats.Speed)))
        Call PutVar(FileName, "CLASS" & i, "MAGI", CStr(Class(i).Stat(Stats.Magic)))
    Next
End Sub

Public Function CheckClasses() As Boolean
Dim FileName As String

    FileName = App.Path & "\data\classes.ini"

    If Not FileExist(FileName, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Public Sub ClearClasses()
Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next
End Sub

' ***********
' ** Items **
' ***********

Public Sub SaveItems()
Dim i As Long
    
    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
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

    Next
End Sub

Public Sub CheckItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If
    Next
    
End Sub

Public Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
End Sub

Public Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
End Sub

' ***********
' ** Shops **
' ***********

Public Sub SaveShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
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

    Next
End Sub

Public Sub CheckShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If
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

    Next
End Sub

Public Sub CheckSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If
    Next
End Sub

Public Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).LevelReq = 1 'Needs to be 1 for the spell editor
End Sub

Public Sub ClearSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next
End Sub

' **********
' ** Npcs **
' **********

Public Sub SaveNpcs()
Dim i As Long
    
    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next
End Sub

Public Sub SaveNpc(ByVal NpcNum As Long)
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\Data\Npcs\Npc" & NpcNum & ".dat"
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
        FileName = App.Path & "\Data\Npcs\Npc" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Npc(i)
        Close #F

    Next
End Sub

Public Sub CheckNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        If Not FileExist("\Data\Npcs\Npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If
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
    Next
End Sub

' **********
' ** Maps **
' **********

Public Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\Data\maps\map" & MapNum & ".dat"
        
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Map(MapNum)
    Close #F
End Sub

Public Sub SaveMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next
End Sub

Public Sub LoadMaps()
Dim FileName As String
Dim i As Long
Dim F As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        FileName = App.Path & "\Data\maps\map" & i & ".dat"
        
        F = FreeFile
        Open FileName For Binary Access Read As #F
            Get #F, , Map(i)
        Close #F
    Next
End Sub

Public Sub LoadMapsX()
Dim FileName As String
Dim i As Long
Dim F As Long

    ''This is an experimental optimization for map loading.
    ''Cuts load times for 5000 maps from 21 seconds to 5.
    ''Speed benefits don't outweigh the hassle.
    ''Use at own risk. (Please note that corresponding
    ''save method will need to be written)

    If Not FileExist("\Data\maps\maps.map") Then
        FileName = App.Path & "\Data\maps\maps.map"
        
        F = FreeFile
        Open FileName For Binary As #F
            For i = 1 To MAX_MAPS
                Put #F, , Map(i)
            Next
        Close #F
    End If

    FileName = App.Path & "\Data\maps\maps.map"
        
    F = FreeFile
    Open FileName For Binary Access Read As #F
        For i = 1 To MAX_MAPS
            Get #F, , Map(i)
        Next
    Close #F
End Sub

Public Sub CheckMaps()
Dim i As Long
        
    For i = 1 To MAX_MAPS
        
        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If
    Next
End Sub

Public Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, Index)), LenB(MapItem(MapNum, Index)))
End Sub

Public Sub ClearMapItems()
Dim X As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, y)
        Next
    Next
End Sub

Public Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum, Index)), LenB(MapNpc(MapNum, Index)))
End Sub

Public Sub ClearMapNpcs()
Dim X As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, y)
        Next
    Next
End Sub

Public Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    
    Map(MapNum).TileSet = 1
    
    Map(MapNum).Name = vbNullString
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    
    ' Reset the map cache array for this map.
    MapCache(MapNum).Cache = vbNullString
    
End Sub

Public Sub ClearMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next
End Sub


Public Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Public Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            GetClassMaxVital = (1 + (Class(ClassNum).Stat(Stats.Strength) \ 2) + Class(ClassNum).Stat(Stats.Strength)) * 2
        Case MP
            GetClassMaxVital = (1 + (Class(ClassNum).Stat(Stats.Magic) \ 2) + Class(ClassNum).Stat(Stats.Magic)) * 2
        Case SP
            GetClassMaxVital = (1 + (Class(ClassNum).Stat(Stats.Speed) \ 2) + Class(ClassNum).Stat(Stats.Speed)) * 2
    End Select
End Function

Public Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////

Public Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function
Public Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Public Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function
Public Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Public Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(TempPlayer(Index).CharNum).Name)
End Function
Public Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(TempPlayer(Index).CharNum).Name = Name
End Sub

Public Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(TempPlayer(Index).CharNum).Guild)
End Function
Public Sub SetPlayerGuild(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(TempPlayer(Index).CharNum).Guild = Name
End Sub

Public Function GetPlayerGAccess(ByVal Index As Long) As Long
    GetPlayerGAccess = Player(Index).Char(TempPlayer(Index).CharNum).GuildAccess
End Function
Public Sub SetPlayerGAccess(ByVal Index As Long, ByVal GAccess As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).GuildAccess = GAccess
End Sub

Public Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(TempPlayer(Index).CharNum).Class
End Function
Public Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Class = ClassNum
End Sub

Public Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(TempPlayer(Index).CharNum).Sprite
End Function
Public Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Sprite = Sprite
End Sub

Public Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(TempPlayer(Index).CharNum).Level
End Function
Public Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    If Level > MAX_LEVELS Then Exit Sub
    Player(Index).Char(TempPlayer(Index).CharNum).Level = Level
End Sub

Public Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = (GetPlayerLevel(Index) + 1) * (GetPlayerStat(Index, Stats.Strength) + GetPlayerStat(Index, Stats.Defense) + GetPlayerStat(Index, Stats.Magic) + GetPlayerStat(Index, Stats.Speed) + GetPlayerPOINTS(Index)) * 25
End Function

Public Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(TempPlayer(Index).CharNum).Exp
End Function
Public Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Exp = Exp
End Sub

Public Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(TempPlayer(Index).CharNum).Access
End Function
Public Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Access = Access
End Sub

Public Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(TempPlayer(Index).CharNum).PK
End Function
Public Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).PK = PK
End Sub

Public Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    GetPlayerVital = Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital)
End Function
Public Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = Value
    
    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If
    If GetPlayerVital(Index, Vital) < 0 Then
        Player(Index).Char(TempPlayer(Index).CharNum).Vital(Vital) = 0
    End If
End Sub

Public Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
Dim CharNum As Long

    Select Case Vital
        Case HP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + (GetPlayerStat(Index, Stats.Strength) \ 2) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.Strength)) * 2
        Case MP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + (GetPlayerStat(Index, Stats.Magic) \ 2) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.Magic)) * 2
        Case SP
            CharNum = TempPlayer(Index).CharNum
            GetPlayerMaxVital = (Player(Index).Char(CharNum).Level + (GetPlayerStat(Index, Stats.Speed) \ 2) + Class(Player(Index).Char(CharNum).Class).Stat(Stats.Speed)) * 2
    End Select
End Function

Public Function GetPlayerStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    GetPlayerStat = Player(Index).Char(TempPlayer(Index).CharNum).Stat(Stat)
End Function
Public Sub SetPlayerStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Stat(Stat) = Value
End Sub

Public Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(TempPlayer(Index).CharNum).POINTS
End Function
Public Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).POINTS = POINTS
End Sub

Public Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(TempPlayer(Index).CharNum).Map
End Function
Public Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(TempPlayer(Index).CharNum).Map = MapNum
    End If
End Sub

Public Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(TempPlayer(Index).CharNum).X
End Function
Public Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).X = X
End Sub

Public Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(TempPlayer(Index).CharNum).y
End Function
Public Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).y = y
End Sub

Public Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(TempPlayer(Index).CharNum).Dir
End Function
Public Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Dir = Dir
End Sub

Public Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Public Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Num
End Function
Public Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Num = ItemNum
End Sub

Public Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Value
End Function
Public Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Public Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Dur
End Function
Public Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Public Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(TempPlayer(Index).CharNum).Spell(SpellSlot)
End Function
Public Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(TempPlayer(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Public Function GetPlayerEquipmentSlot(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Byte
    GetPlayerEquipmentSlot = Player(Index).Char(TempPlayer(Index).CharNum).Equipment(EquipmentSlot)
End Function
Public Sub SetPlayerEquipmentSlot(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Char(TempPlayer(Index).CharNum).Equipment(EquipmentSlot) = InvNum
End Sub

Public Function GetPlayerInvItemName(ByVal Index As Long, ByVal InvSlot As Long) As String
Dim ItemNum As Long
    ItemNum = GetPlayerInvItemNum(Index, InvSlot)
    If ItemNum > 0 Then
        GetPlayerInvItemName = Trim$(Item(ItemNum).Name)
    End If
End Function


