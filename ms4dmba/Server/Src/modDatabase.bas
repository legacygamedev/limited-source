Attribute VB_Name = "modDatabase"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)


' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
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

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
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

Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
Dim FileName As String
Dim IP As String
Dim F As Long
Dim i As Long

    FileName = App.Path & "data\banlist.txt"
    
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
        Print #F, IP & "," & "Server"
    Close #F
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & "the Server" & "!", White)
    Call AddLog("The Server" & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & "The Server" & "!")
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
Dim FileName As String

    FileName = "accounts\" & Trim(Name) & ".bin"
   
    If FileExist(FileName) Then
        AccountExist = True
    End If
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
       
        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If
    
End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
Dim i As Long

    Player(Index).Login = Name
    Player(Index).Password = Password
    
    For i = 1 To MAX_CHARS
        Call ClearChar(Index, i)
    Next
    
    Call SavePlayer(Index)
End Sub

Sub DeleteName(ByVal Name As String)
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

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If LenB(Trim$(Player(Index).Char(CharNum).Name)) > 0 Then
        CharExist = True
    End If
End Function

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
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
        Player(Index).Char(CharNum).x = START_X
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

Sub DelChar(ByVal Index As Long, ByVal CharNum As Long)
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

Sub SaveAllPlayersOnline()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next
End Sub

Sub SavePlayer(ByVal Index As Long)
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".bin"
       
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Player(Index)
    Close #F
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
Dim FileName As String
Dim F As Long

    Call ClearPlayer(Index)
   
    FileName = App.Path & "\accounts\" & Trim(Name) & ".bin"

    F = FreeFile
    Open FileName For Binary As #F
        Get #F, , Player(Index)
    Close #F
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim i As Long

    Set TempPlayer(Index).Buffer = New clsBuffer

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

Sub ClearChar(ByVal Index As Long, ByVal CharNum As Long)
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

Sub LoadClasses()
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
    
    Call ClearClasses
    
    For i = 1 To Max_Classes
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).Sprite = Val(GetVar(FileName, "CLASS" & i, "Sprite"))
        Class(i).Stat(Stats.Strength) = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class(i).Stat(Stats.Defense) = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class(i).Stat(Stats.Speed) = Val(GetVar(FileName, "CLASS" & i, "Speed"))
        Class(i).Stat(Stats.Magic) = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
    Next
End Sub

Sub SaveClasses()
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

Function CheckClasses() As Boolean
Dim FileName As String

    FileName = App.Path & "\data\classes.ini"

    If Not FileExist(FileName, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next
End Sub

' ***********
' ** Items **
' ***********

Sub SaveItems()
Dim i As Long
    
    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next
End Sub

Sub SaveItem(ByVal ItemNum As Long)
Dim FileName As String
Dim F  As Long
    
    FileName = App.Path & "\Data\items\item" & ItemNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Item(ItemNum)
    Close #F
End Sub

Sub LoadItems()
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

Sub CheckItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If
    Next
    
End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
End Sub

' ***********
' ** Shops **
' ***********

Sub SaveShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next
End Sub

Sub SaveShop(ByVal ShopNum As Long)
    Dim FileName As String
    Dim F As Long

    FileName = App.Path & "\Data\shops\shop" & ShopNum & ".dat"

    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Shop(ShopNum)
    Close #F
End Sub

Sub LoadShops()
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

Sub CheckShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If
    Next
End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
    Shop(Index).JoinSay = vbNullString
    Shop(Index).LeaveSay = vbNullString
End Sub

Sub ClearShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next
End Sub

' ************
' ** Spells **
' ************

Sub SaveSpell(ByVal SpellNum As Long)
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\Data\spells\spells" & SpellNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
       Put #F, , Spell(SpellNum)
    Close #F
End Sub

Sub SaveSpells()
    Dim i As Long

    Call SetStatus("Saving spells... ")
    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next
End Sub

Sub LoadSpells()
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

Sub CheckSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If
    Next
End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).LevelReq = 1 'Needs to be 1 for the spell editor
End Sub

Sub ClearSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next
End Sub

' **********
' ** NPCs **
' **********

Sub SaveNpcs()
Dim i As Long
    
    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\Data\npcs\npc" & NpcNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Npc(NpcNum)
    Close #F
End Sub

Sub LoadNpcs()
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

    Next
End Sub

Sub CheckNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If
    Next
End Sub

Sub ClearNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).Name = vbNullString
    Npc(Index).AttackSay = vbNullString
End Sub

Sub ClearNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next
End Sub

' **********
' ** Maps **
' **********

Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim F As Long
Dim x As Long
Dim y As Long

    FileName = App.Path & "\Data\maps\map" & MapNum & ".dat"
        
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Map(MapNum).Name
        Put #F, , Map(MapNum).Revision
        Put #F, , Map(MapNum).Moral
        Put #F, , Map(MapNum).Up
        Put #F, , Map(MapNum).Down
        Put #F, , Map(MapNum).Left
        Put #F, , Map(MapNum).Right
        Put #F, , Map(MapNum).Music
        Put #F, , Map(MapNum).BootMap
        Put #F, , Map(MapNum).BootX
        Put #F, , Map(MapNum).BootY
        Put #F, , Map(MapNum).Shop
        Put #F, , Map(MapNum).MaxX
        Put #F, , Map(MapNum).MaxY
        
        For x = 0 To Map(MapNum).MaxX
            For y = 0 To Map(MapNum).MaxY
                Put #F, , Map(MapNum).Tile(x, y)
            Next
        Next
        
        For x = 1 To MAX_MAP_NPCS
            Put #F, , Map(MapNum).Npc(x)
        Next
    Close #F
    
    DoEvents
End Sub

Sub SaveMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next
End Sub

Sub LoadMaps()
Dim FileName As String
Dim i As Long
Dim F As Long
Dim x As Long
Dim y As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        FileName = App.Path & "\Data\maps\map" & i & ".dat"
        
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Map(i).Name
            Get #F, , Map(i).Revision
            Get #F, , Map(i).Moral
            Get #F, , Map(i).Up
            Get #F, , Map(i).Down
            Get #F, , Map(i).Left
            Get #F, , Map(i).Right
            Get #F, , Map(i).Music
            Get #F, , Map(i).BootMap
            Get #F, , Map(i).BootX
            Get #F, , Map(i).BootY
            Get #F, , Map(i).Shop
            Get #F, , Map(i).MaxX
            Get #F, , Map(i).MaxY
           
            ' have to set the tile()
            ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)
            
            For x = 0 To Map(i).MaxX
                For y = 0 To Map(i).MaxY
                    Get #F, , Map(i).Tile(x, y)
                Next
            Next
            
            For x = 1 To MAX_MAP_NPCS
                Get #F, , Map(i).Npc(x)
            Next
        Close #F
        
        ClearTempTile i
        
        DoEvents
    Next
End Sub

Sub CheckMaps()
Dim i As Long
        
    For i = 1 To MAX_MAPS
        
        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If
    Next
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, Index)), LenB(MapItem(MapNum, Index)))
End Sub

Sub ClearMapItems()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next
End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    ReDim MapNpc(MapNum).Npc(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).Npc(Index)), LenB(MapNpc(MapNum).Npc(Index)))
End Sub

Sub ClearMapNpcs()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next
End Sub

Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    
    Map(MapNum).TileSet = 1
    
    Map(MapNum).Name = vbNullString
    
    Map(MapNum).MaxX = MAX_MAPX
    Map(MapNum).MaxY = MAX_MAPY
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString
    
End Sub

Sub ClearMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next
End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            GetClassMaxVital = (1 + (Class(ClassNum).Stat(Stats.Strength) \ 2) + Class(ClassNum).Stat(Stats.Strength)) * 2
        Case MP
            GetClassMaxVital = (1 + (Class(ClassNum).Stat(Stats.Magic) \ 2) + Class(ClassNum).Stat(Stats.Magic)) * 2
        Case SP
            GetClassMaxVital = (1 + (Class(ClassNum).Stat(Stats.Speed) \ 2) + Class(ClassNum).Stat(Stats.Speed)) * 2
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

