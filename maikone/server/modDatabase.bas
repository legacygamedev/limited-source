Attribute VB_Name = "modDatabase"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Const START_MAP = 1
Public Const START_X = MAX_MAPX / 2
Public Const START_Y = MAX_MAPY / 2

Public Const ADMIN_LOG = "System\Admin.txt"
Public Const PLAYER_LOG = "System\Player.txt"

Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Function FileExist(ByVal filename As String) As Boolean
    If Dir(App.Path & "\" & filename) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Sub SavePlayer(ByVal Index As Long)
Dim filename As String
Dim i As Long
Dim n As Long

    filename = App.Path & "\accounts\" & Trim(Player(Index).Login) & ".ini"
    
    Call PutVar(filename, "GENERAL", "Login", Trim(Player(Index).Login))
    Call PutVar(filename, "GENERAL", "Password", Trim(Player(Index).Password))

    For i = 1 To MAX_CHARS
        ' General
        Call PutVar(filename, "CHAR" & i, "Name", Trim$(Player(Index).Char(i).Name))
        Call PutVar(filename, "CHAR" & i, "Class", STR(Player(Index).Char(i).Class))
        Call PutVar(filename, "CHAR" & i, "Sex", STR(Player(Index).Char(i).Sex))
        Call PutVar(filename, "CHAR" & i, "Sprite", STR(Player(Index).Char(i).Sprite))
        Call PutVar(filename, "CHAR" & i, "Level", STR(Player(Index).Char(i).Level))
        Call PutVar(filename, "CHAR" & i, "Exp", STR(Player(Index).Char(i).Exp))
        Call PutVar(filename, "CHAR" & i, "Access", STR(Player(Index).Char(i).Access))
        Call PutVar(filename, "CHAR" & i, "PK", STR(Player(Index).Char(i).PK))
        Call PutVar(filename, "CHAR" & i, "Rebirth", STR(Player(Index).Char(i).Rebirth))
        Call PutVar(filename, "CHAR" & i, "Guild", STR(Player(Index).Char(i).Guild))
        
        ' Vitals
        Call PutVar(filename, "CHAR" & i, "HP", STR(Player(Index).Char(i).HP))
        Call PutVar(filename, "CHAR" & i, "MP", STR(Player(Index).Char(i).MP))
        Call PutVar(filename, "CHAR" & i, "SP", STR(Player(Index).Char(i).SP))
        
        ' Stats
        Call PutVar(filename, "CHAR" & i, "STR", STR(Player(Index).Char(i).STR))
        Call PutVar(filename, "CHAR" & i, "DEF", STR(Player(Index).Char(i).DEF))
        Call PutVar(filename, "CHAR" & i, "SPEED", STR(Player(Index).Char(i).Speed))
        Call PutVar(filename, "CHAR" & i, "MAGI", STR(Player(Index).Char(i).Magi))
        Call PutVar(filename, "CHAR" & i, "POINTS", STR(Player(Index).Char(i).POINTS))
        
        ' Worn equipment
        Call PutVar(filename, "CHAR" & i, "ArmorSlot", STR(Player(Index).Char(i).ArmorSlot))
        Call PutVar(filename, "CHAR" & i, "WeaponSlot", STR(Player(Index).Char(i).WeaponSlot))
        Call PutVar(filename, "CHAR" & i, "HelmetSlot", STR(Player(Index).Char(i).HelmetSlot))
        Call PutVar(filename, "CHAR" & i, "ShieldSlot", STR(Player(Index).Char(i).ShieldSlot))
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(i).Map = 0 Then
            Player(Index).Char(i).Map = START_MAP
            Player(Index).Char(i).X = START_X
            Player(Index).Char(i).y = START_Y
        End If
            
        ' Position
        Call PutVar(filename, "CHAR" & i, "Map", STR(Player(Index).Char(i).Map))
        Call PutVar(filename, "CHAR" & i, "X", STR(Player(Index).Char(i).X))
        Call PutVar(filename, "CHAR" & i, "Y", STR(Player(Index).Char(i).y))
        Call PutVar(filename, "CHAR" & i, "Dir", STR(Player(Index).Char(i).Dir))
        
        ' Inventory
        For n = 1 To MAX_INV
            Call PutVar(filename, "CHAR" & i, "InvItemNum" & n, STR(Player(Index).Char(i).Inv(n).num))
            Call PutVar(filename, "CHAR" & i, "InvItemVal" & n, STR(Player(Index).Char(i).Inv(n).Value))
            Call PutVar(filename, "CHAR" & i, "InvItemDur" & n, STR(Player(Index).Char(i).Inv(n).Dur))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Call PutVar(filename, "CHAR" & i, "Spell" & n, STR(Player(Index).Char(i).Spell(n)))
        Next n
    Next i
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
Dim filename As String
Dim i As Long
Dim n As Long

    Call ClearPlayer(Index)
    
    filename = App.Path & "\accounts\" & Trim(Name) & ".ini"

    Player(Index).Login = GetVar(filename, "GENERAL", "Login")
    Player(Index).Password = GetVar(filename, "GENERAL", "Password")

        For i = 1 To MAX_CHARS
            ' General
            Player(Index).Char(i).Name = GetVar(filename, "CHAR" & i, "Name")
            Player(Index).Char(i).Sex = Val(GetVar(filename, "CHAR" & i, "Sex"))
            Player(Index).Char(i).Class = Val(GetVar(filename, "CHAR" & i, "Class"))
            Player(Index).Char(i).Sprite = Val(GetVar(filename, "CHAR" & i, "Sprite"))
            Player(Index).Char(i).Level = Val(GetVar(filename, "CHAR" & i, "Level"))
            Player(Index).Char(i).Exp = Val(GetVar(filename, "CHAR" & i, "Exp"))
            Player(Index).Char(i).Access = Val(GetVar(filename, "CHAR" & i, "Access"))
            Player(Index).Char(i).PK = Val(GetVar(filename, "CHAR" & i, "PK"))
            Player(Index).Char(i).Rebirth = Val(GetVar(filename, "CHAR" & i, "Rebirth"))
            Player(Index).Char(i).Guild = Val(GetVar(filename, "CHAR" & i, "Guild"))
        
            ' Vitals
            Player(Index).Char(i).HP = Val(GetVar(filename, "CHAR" & i, "HP"))
            Player(Index).Char(i).MP = Val(GetVar(filename, "CHAR" & i, "MP"))
            Player(Index).Char(i).SP = Val(GetVar(filename, "CHAR" & i, "SP"))
        
            ' Stats
            Player(Index).Char(i).STR = Val(GetVar(filename, "CHAR" & i, "STR"))
            Player(Index).Char(i).DEF = Val(GetVar(filename, "CHAR" & i, "DEF"))
            Player(Index).Char(i).Speed = Val(GetVar(filename, "CHAR" & i, "SPEED"))
            Player(Index).Char(i).Magi = Val(GetVar(filename, "CHAR" & i, "MAGI"))
            Player(Index).Char(i).POINTS = Val(GetVar(filename, "CHAR" & i, "POINTS"))
        
            ' Worn equipment
            Player(Index).Char(i).ArmorSlot = Val(GetVar(filename, "CHAR" & i, "ArmorSlot"))
            Player(Index).Char(i).WeaponSlot = Val(GetVar(filename, "CHAR" & i, "WeaponSlot"))
            Player(Index).Char(i).HelmetSlot = Val(GetVar(filename, "CHAR" & i, "HelmetSlot"))
            Player(Index).Char(i).ShieldSlot = Val(GetVar(filename, "CHAR" & i, "ShieldSlot"))
        
            ' Position
            Player(Index).Char(i).Map = Val(GetVar(filename, "CHAR" & i, "Map"))
            Player(Index).Char(i).X = Val(GetVar(filename, "CHAR" & i, "X"))
            Player(Index).Char(i).y = Val(GetVar(filename, "CHAR" & i, "Y"))
            Player(Index).Char(i).Dir = Val(GetVar(filename, "CHAR" & i, "Dir"))
        
            ' Check to make sure that they aren't on map 0, if so reset'm
            If Player(Index).Char(i).Map = 0 Then
                Player(Index).Char(i).Map = START_MAP
                Player(Index).Char(i).X = START_X
                Player(Index).Char(i).y = START_Y
            End If
        
            ' Inventory
            For n = 1 To MAX_INV
                Player(Index).Char(i).Inv(n).num = Val(GetVar(filename, "CHAR" & i, "InvItemNum" & n))
                Player(Index).Char(i).Inv(n).Value = Val(GetVar(filename, "CHAR" & i, "InvItemVal" & n))
                Player(Index).Char(i).Inv(n).Dur = Val(GetVar(filename, "CHAR" & i, "InvItemDur" & n))
            Next n
        
            ' Spells
            For n = 1 To MAX_PLAYER_SPELLS
                Player(Index).Char(i).Spell(n) = Val(GetVar(filename, "CHAR" & i, "Spell" & n))
            Next n
        Next i
End Sub

Function AccountExist(ByVal Name As String) As Boolean
Dim filename As String

    filename = "accounts\" & Trim(Name) & ".ini"
    
    If FileExist(filename) Then
        AccountExist = True
    Else
        AccountExist = False
    End If
End Function

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If Trim(Player(Index).Char(CharNum).Name) <> "" Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim filename As String
Dim RightPassword As String

    PasswordOK = False
    
    If AccountExist(Name) Then
        filename = App.Path & "\accounts\" & Trim(Name) & ".ini"
        RightPassword = GetVar(filename, "GENERAL", "Password")
        
        If UCase(Trim(Password)) = UCase(Trim(RightPassword)) Then
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
    Next i
    
    Call SavePlayer(Index)
End Sub

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
Dim f As Long

    If Trim(Player(Index).Char(CharNum).Name) = "" Then
        Player(Index).CharNum = CharNum
        
        Player(Index).Char(CharNum).Name = Name
        Player(Index).Char(CharNum).Sex = Sex
        Player(Index).Char(CharNum).Class = ClassNum
        
        If Player(Index).Char(CharNum).Sex = SEX_MALE Then
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).Sprite
        Else
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).Sprite
        End If
        
        Player(Index).Char(CharNum).Level = 1
                    
        Player(Index).Char(CharNum).STR = Class(ClassNum).STR
        Player(Index).Char(CharNum).DEF = Class(ClassNum).DEF
        Player(Index).Char(CharNum).Speed = Class(ClassNum).Speed
        Player(Index).Char(CharNum).Magi = Class(ClassNum).Magi
        
        Player(Index).Char(CharNum).Map = START_MAP
        Player(Index).Char(CharNum).X = START_X
        Player(Index).Char(CharNum).y = START_Y
            
        Player(Index).Char(CharNum).HP = GetPlayerMaxHP(Index)
        Player(Index).Char(CharNum).MP = GetPlayerMaxMP(Index)
        Player(Index).Char(CharNum).SP = GetPlayerMaxSP(Index)
                
        ' Append name to file
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Append As #f
            Print #f, Name
        Close #f
        
        Call SavePlayer(Index)
            
        Exit Sub
    End If
End Sub

Sub DelChar(ByVal Index As Long, ByVal CharNum As Long)
Dim f1 As Long, f2 As Long
Dim s As String

    Call DeleteName(Player(Index).Char(CharNum).Name)
    Call ClearChar(Index, CharNum)
    Call SavePlayer(Index)
End Sub

Function FindChar(ByVal Name As String) As Boolean
Dim f As Long
Dim s As String

    FindChar = False
    
    f = FreeFile
    Open App.Path & "\accounts\charlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            
            If Trim(LCase(s)) = Trim(LCase(Name)) Then
                FindChar = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
End Function

Sub SaveAllPlayersOnline()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next i
End Sub

Sub LoadClasses()
Dim filename As String
Dim i As Long

    Call CheckClasses
    filename = App.Path & "\Classes\Info.ini"
    Max_Classes = Val(GetVar(filename, "INFO", "MaxClasses"))
    ReDim Class(1 To Max_Classes) As ClassRec
    Call ClearClasses
    
    For i = 1 To Max_Classes
        Call SetStatus("Loading Classes.. Please wait.", Int((i / Max_Classes) * 100))
        filename = App.Path & "\Classes\Class" & i & ".ini"
        Class(i).Name = GetVar(filename, "CLASS", "Name")
        Class(i).Sprite = GetVar(filename, "CLASS", "Sprite")
        Class(i).STR = Val(GetVar(filename, "CLASS", "STR"))
        Class(i).DEF = Val(GetVar(filename, "CLASS", "DEF"))
        Class(i).Speed = Val(GetVar(filename, "CLASS", "SPEED"))
        Class(i).Magi = Val(GetVar(filename, "CLASS", "MAGI"))
        DoEvents
    Next
End Sub

Sub SaveClasses()
Dim filename As String
Dim i As Long

    filename = App.Path & "\Classes\Info.ini"
    
    If Not FileExist("Classes\Info.ini") Then
        Call SpecialPutVar(filename, "INFO", "MaxClasses", 3)
        Max_Classes = 3
    End If
    
    For i = 1 To Max_Classes
        Call SetStatus("Saving Classes.. Please wait.", Int((i / Max_Classes) * 100))
        DoEvents
        
        filename = App.Path & "\Classes\Class" & i & ".ini"
    
    If Not FileExist("Classes\Class" & i & ".ini") Then
        Call PutVar(filename, "CLASS", "Name", Trim$(Class(i).Name))
        Call PutVar(filename, "CLASS", "Sprite", STR(Class(i).Sprite))
        Call PutVar(filename, "CLASS", "STR", STR(Class(i).STR))
        Call PutVar(filename, "CLASS", "DEF", STR(Class(i).DEF))
        Call PutVar(filename, "CLASS", "SPEED", STR(Class(i).Speed))
        Call PutVar(filename, "CLASS", "MAGI", STR(Class(i).Magi))
    End If
        Next
End Sub

Sub CheckClasses()
    If Not FileExist("Classes\Info.ini") Then
        Call SaveClasses
    End If
End Sub

Sub SaveItems()
Dim i As Long
    
    For i = 1 To MAX_ITEMS

        If Not FileExist("System\items\item" & i & ".dat") Then
            Call SetStatus("Saving Items.. Please wait.", Int((i / MAX_ITEMS) * 100))
            DoEvents

            Call SaveItem(i)
        End If
    Next
End Sub

Sub SaveItem(ByVal ItemNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\System\items\item" & ItemNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Item(ItemNum)
    Close #f
End Sub

Sub LoadItems()
Dim filename As String
Dim i As Long
Dim f As Long

    Call CheckItems
    For i = 1 To MAX_ITEMS
        Call SetStatus("Loading items.. Please wait.", Int((i / MAX_ITEMS) * 100))
        
        filename = App.Path & "\System\Items\Item" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Item(i)
        Close #f
        DoEvents
        
    Next
End Sub

Sub CheckItems()
    Call SaveItems
End Sub

Sub SaveShops()
Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("System\shops\shop" & i & ".dat") Then
            Call SetStatus("Saving Shops.. Please wait.", Int((i / MAX_SHOPS) * 100))
            DoEvents

            Call SaveShop(i)
        End If
    Next
End Sub

Sub SaveShop(ByVal ShopNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\System\shops\shop" & ShopNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Shop(ShopNum)
    Close #f
End Sub

Sub LoadShops()
Dim filename As String
Dim i As Long, f As Long

    Call CheckShops
    For i = 1 To MAX_SHOPS
        Call SetStatus("Loading Shops.. Please wait.", Int((i / MAX_SHOPS) * 100))
        filename = App.Path & "\System\shops\shop" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Shop(i)
        Close #f
        DoEvents

    Next
End Sub

Sub CheckShops()
    Call SaveShops
End Sub

Sub SaveSpell(ByVal spellnum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\System\spells\spells" & spellnum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Spell(spellnum)
    Close #f
End Sub

Sub SaveSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("System\spells\spells" & i & ".dat") Then
            Call SetStatus("Saving Spells.. Please wait.", Int((i / MAX_SPELLS) * 100))
            DoEvents

            Call SaveSpell(i)
        End If
    Next
End Sub

Sub LoadSpells()
Dim filename As String
Dim i As Long
Dim f As Long

    Call CheckSpells
    For i = 1 To MAX_SPELLS
        Call SetStatus("Loading Spells.. Please wait.", Int((i / MAX_SPELLS) * 100))
        filename = App.Path & "\System\spells\spells" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Spell(i)
        Close #f
        DoEvents

    Next
End Sub

Sub CheckSpells()
    Call SaveSpells
End Sub

Sub SaveNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("System\npcs\npc" & i & ".dat") Then
            Call SetStatus("Saving Npcs.. Please wait.", Int((i / MAX_NPCS) * 100))
            DoEvents

            Call SaveNpc(i)
        End If
    Next
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\System\npcs\npc" & NpcNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Npc(NpcNum)
    Close #f
End Sub

Sub LoadNpcs()
Dim filename As String
Dim i As Long
Dim f As Long

    Call CheckNpcs
    For i = 1 To MAX_NPCS
        Call SetStatus("Loading Npcs.. Please wait.", Int((i / MAX_NPCS) * 100))
        filename = App.Path & "\System\npcs\npc" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Npc(i)
        Close #f
        DoEvents

    Next
End Sub

Sub CheckNpcs()
    Call SaveNpcs
End Sub

Sub SaveMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\System\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub SaveMaps()
Dim filename As String
Dim i As Long
Dim f As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next i
End Sub

Sub LoadMaps()
Dim filename As String
Dim i As Long
Dim f As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        filename = App.Path & "\System\maps\map" & i & ".dat"
        Call SetStatus("Loading Maps.. Please wait.", Int((i / MAX_MAPS) * 100))
        
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , Map(i)
        Close #f
    
        DoEvents
    Next i
End Sub

Sub ConvertOldMapsToNew()
Dim filename As String
Dim i As Long
Dim f As Long
Dim X As Long, y As Long
Dim OldMap As OldMapRec
Dim NewMap As MapRec

    For i = 1 To MAX_MAPS
        filename = App.Path & "\System\maps\map" & i & ".dat"
        
        ' Get the old file
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , OldMap
        Close #f
        
        ' Delete the old file
        Call Kill(filename)
        
        ' Convert
        NewMap.Name = OldMap.Name
        NewMap.Revision = OldMap.Revision + 1
        NewMap.Moral = OldMap.Moral
        NewMap.Up = OldMap.Up
        NewMap.Down = OldMap.Down
        NewMap.Left = OldMap.Left
        NewMap.Right = OldMap.Right
        NewMap.Music = OldMap.Music
        NewMap.BootMap = OldMap.BootMap
        NewMap.BootX = OldMap.BootX
        NewMap.BootY = OldMap.BootY
        NewMap.Shop = OldMap.Shop
        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                NewMap.Tile(X, y).Ground = OldMap.Tile(X, y).Ground
                NewMap.Tile(X, y).Mask = OldMap.Tile(X, y).Mask
                NewMap.Tile(X, y).Anim = OldMap.Tile(X, y).Anim
                NewMap.Tile(X, y).Fringe = OldMap.Tile(X, y).Fringe
                NewMap.Tile(X, y).Type = OldMap.Tile(X, y).Type
                NewMap.Tile(X, y).Data1 = OldMap.Tile(X, y).Data1
                NewMap.Tile(X, y).Data2 = OldMap.Tile(X, y).Data2
                NewMap.Tile(X, y).Data3 = OldMap.Tile(X, y).Data3
            Next X
        Next y
        
        For X = 1 To MAX_MAP_NPCS
            NewMap.Npc(X) = OldMap.Npc(X)
        Next X
        
        ' Set new values to 0 or null
        NewMap.Indoors = NO
        
        ' Save the new map
        f = FreeFile
        Open filename For Binary As #f
            Put #f, , NewMap
        Close #f
    Next i
End Sub

Sub CheckMaps()
Dim filename As String
Dim X As Long
Dim y As Long
Dim i As Long
Dim n As Long

    Call ClearMaps
        
    For i = 1 To MAX_MAPS
        filename = "System\maps\map" & i & ".dat"
        
        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(filename) Then
            Call SaveMap(i)
        End If
    Next i
End Sub

Sub AddLog(ByVal text As String, ByVal FN As String)
Dim filename As String
Dim f As Long

    If ServerLog = True Then
        filename = App.Path & "\" & FN
    
        If Not FileExist(FN) Then
            f = FreeFile
            Open filename For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open filename For Append As #f
            Print #f, Time & ": " & text
        Close #f
    End If
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
Dim filename, IP As String
Dim f As Long, i As Long

    filename = App.Path & "\System\Banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("System\Banlist.txt") Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If
    
    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)
            
    For i = Len(IP) To 1 Step -1
        If Mid(IP, i, 1) = "." Then
            Exit For
        End If
    Next i
    IP = Mid(IP, 1, i)
            
    f = FreeFile
    Open filename For Append As #f
        Print #f, IP & "," & GetPlayerName(BannedByIndex)
    Close #f
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Sub DeleteName(ByVal Name As String)
Dim f1 As Long, f2 As Long
Dim s As String

    Call FileCopy(App.Path & "\accounts\charlist.txt", App.Path & "\accounts\chartemp.txt")
    
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\accounts\charlist.txt" For Output As #f2
        
    Do While Not EOF(f1)
        Input #f1, s
        If Trim(LCase(s)) <> Trim(LCase(Name)) Then
            Print #f2, s
        End If
    Loop
    
    Close #f1
    Close #f2
    
    Call Kill(App.Path & "\accounts\chartemp.txt")
End Sub

Sub SpecialPutVar(File As String, _
   Header As String, _
   Var As String, _
   Value As String)

    ' Same as the one below except it keeps all 0 and blank values (used for config)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub
