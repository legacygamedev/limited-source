Attribute VB_Name = "modDatabase"
Option Explicit

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

Function FileExist(ByVal FileName As String) As Boolean
    If Dir(App.Path & "\" & FileName) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Sub SavePlayer(ByVal index As Long)
Dim FileName As String
Dim i As Long
Dim n As Long

    FileName = App.Path & "\accounts\" & Trim(Player(index).Login) & ".ini"
    
    Call PutVar(FileName, "GENERAL", "Login", Trim(Player(index).Login))
    Call PutVar(FileName, "GENERAL", "Password", Trim(Player(index).Password))

    For i = 1 To MAX_CHARS
        ' General
        Call PutVar(FileName, "CHAR" & i, "Name", Trim(Player(index).Char(i).Name))
        Call PutVar(FileName, "CHAR" & i, "Class", STR(Player(index).Char(i).Class))
        Call PutVar(FileName, "CHAR" & i, "Sex", STR(Player(index).Char(i).Sex))
        Call PutVar(FileName, "CHAR" & i, "Sprite", STR(Player(index).Char(i).Sprite))
        Call PutVar(FileName, "CHAR" & i, "Level", STR(Player(index).Char(i).Level))
        Call PutVar(FileName, "CHAR" & i, "Exp", STR(Player(index).Char(i).Exp))
        Call PutVar(FileName, "CHAR" & i, "Access", STR(Player(index).Char(i).Access))
        Call PutVar(FileName, "CHAR" & i, "PK", STR(Player(index).Char(i).PK))
        Call PutVar(FileName, "CHAR" & i, "Guild", STR(Player(index).Char(i).Guild))
        
        ' Vitals
        Call PutVar(FileName, "CHAR" & i, "HP", STR(Player(index).Char(i).HP))
        Call PutVar(FileName, "CHAR" & i, "MP", STR(Player(index).Char(i).MP))
        Call PutVar(FileName, "CHAR" & i, "SP", STR(Player(index).Char(i).SP))
        
        ' Stats
        Call PutVar(FileName, "CHAR" & i, "STR", STR(Player(index).Char(i).STR))
        Call PutVar(FileName, "CHAR" & i, "DEF", STR(Player(index).Char(i).DEF))
        Call PutVar(FileName, "CHAR" & i, "SPEED", STR(Player(index).Char(i).SPEED))
        Call PutVar(FileName, "CHAR" & i, "MAGI", STR(Player(index).Char(i).MAGI))
        Call PutVar(FileName, "CHAR" & i, "POINTS", STR(Player(index).Char(i).POINTS))
        
        ' Worn equipment
        Call PutVar(FileName, "CHAR" & i, "ArmorSlot", STR(Player(index).Char(i).ArmorSlot))
        Call PutVar(FileName, "CHAR" & i, "WeaponSlot", STR(Player(index).Char(i).WeaponSlot))
        Call PutVar(FileName, "CHAR" & i, "HelmetSlot", STR(Player(index).Char(i).HelmetSlot))
        Call PutVar(FileName, "CHAR" & i, "ShieldSlot", STR(Player(index).Char(i).ShieldSlot))
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(index).Char(i).Map = 0 Then
            Player(index).Char(i).Map = START_MAP
            Player(index).Char(i).x = START_X
            Player(index).Char(i).y = START_Y
        End If
            
        ' Position
        Call PutVar(FileName, "CHAR" & i, "Map", STR(Player(index).Char(i).Map))
        Call PutVar(FileName, "CHAR" & i, "X", STR(Player(index).Char(i).x))
        Call PutVar(FileName, "CHAR" & i, "Y", STR(Player(index).Char(i).y))
        Call PutVar(FileName, "CHAR" & i, "Dir", STR(Player(index).Char(i).Dir))
        
        ' Inventory
        For n = 1 To MAX_INV
            Call PutVar(FileName, "CHAR" & i, "InvItemNum" & n, STR(Player(index).Char(i).Inv(n).Num))
            Call PutVar(FileName, "CHAR" & i, "InvItemVal" & n, STR(Player(index).Char(i).Inv(n).Value))
            Call PutVar(FileName, "CHAR" & i, "InvItemDur" & n, STR(Player(index).Char(i).Inv(n).Dur))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Call PutVar(FileName, "CHAR" & i, "Spell" & n, STR(Player(index).Char(i).Spell(n)))
        Next n
    Next i
End Sub

Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
Dim FileName As String
Dim i As Long
Dim n As Long

    Call ClearPlayer(index)
    
    FileName = App.Path & "\accounts\" & Trim(Name) & ".ini"

    Player(index).Login = GetVar(FileName, "GENERAL", "Login")
    Player(index).Password = GetVar(FileName, "GENERAL", "Password")

    For i = 1 To MAX_CHARS
        ' General
        Player(index).Char(i).Name = GetVar(FileName, "CHAR" & i, "Name")
        Player(index).Char(i).Sex = Val(GetVar(FileName, "CHAR" & i, "Sex"))
        Player(index).Char(i).Class = Val(GetVar(FileName, "CHAR" & i, "Class"))
        Player(index).Char(i).Sprite = Val(GetVar(FileName, "CHAR" & i, "Sprite"))
        Player(index).Char(i).Level = Val(GetVar(FileName, "CHAR" & i, "Level"))
        Player(index).Char(i).Exp = Val(GetVar(FileName, "CHAR" & i, "Exp"))
        Player(index).Char(i).Access = Val(GetVar(FileName, "CHAR" & i, "Access"))
        Player(index).Char(i).PK = Val(GetVar(FileName, "CHAR" & i, "PK"))
        Player(index).Char(i).Guild = Val(GetVar(FileName, "CHAR" & i, "Guild"))
        
        ' Vitals
        Player(index).Char(i).HP = Val(GetVar(FileName, "CHAR" & i, "HP"))
        Player(index).Char(i).MP = Val(GetVar(FileName, "CHAR" & i, "MP"))
        Player(index).Char(i).SP = Val(GetVar(FileName, "CHAR" & i, "SP"))
        
        ' Stats
        Player(index).Char(i).STR = Val(GetVar(FileName, "CHAR" & i, "STR"))
        Player(index).Char(i).DEF = Val(GetVar(FileName, "CHAR" & i, "DEF"))
        Player(index).Char(i).SPEED = Val(GetVar(FileName, "CHAR" & i, "SPEED"))
        Player(index).Char(i).MAGI = Val(GetVar(FileName, "CHAR" & i, "MAGI"))
        Player(index).Char(i).POINTS = Val(GetVar(FileName, "CHAR" & i, "POINTS"))
        
        ' Worn equipment
        Player(index).Char(i).ArmorSlot = Val(GetVar(FileName, "CHAR" & i, "ArmorSlot"))
        Player(index).Char(i).WeaponSlot = Val(GetVar(FileName, "CHAR" & i, "WeaponSlot"))
        Player(index).Char(i).HelmetSlot = Val(GetVar(FileName, "CHAR" & i, "HelmetSlot"))
        Player(index).Char(i).ShieldSlot = Val(GetVar(FileName, "CHAR" & i, "ShieldSlot"))
        
        ' Position
        Player(index).Char(i).Map = Val(GetVar(FileName, "CHAR" & i, "Map"))
        Player(index).Char(i).x = Val(GetVar(FileName, "CHAR" & i, "X"))
        Player(index).Char(i).y = Val(GetVar(FileName, "CHAR" & i, "Y"))
        Player(index).Char(i).Dir = Val(GetVar(FileName, "CHAR" & i, "Dir"))
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(index).Char(i).Map = 0 Then
            Player(index).Char(i).Map = START_MAP
            Player(index).Char(i).x = START_X
            Player(index).Char(i).y = START_Y
        End If
        
        ' Inventory
        For n = 1 To MAX_INV
            Player(index).Char(i).Inv(n).Num = Val(GetVar(FileName, "CHAR" & i, "InvItemNum" & n))
            Player(index).Char(i).Inv(n).Value = Val(GetVar(FileName, "CHAR" & i, "InvItemVal" & n))
            Player(index).Char(i).Inv(n).Dur = Val(GetVar(FileName, "CHAR" & i, "InvItemDur" & n))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Player(index).Char(i).Spell(n) = Val(GetVar(FileName, "CHAR" & i, "Spell" & n))
        Next n
    Next i
End Sub

Function AccountExist(ByVal Name As String) As Boolean
Dim FileName As String

    FileName = "accounts\" & Trim(Name) & ".ini"
    
    If FileExist(FileName) Then
        AccountExist = True
    Else
        AccountExist = False
    End If
End Function

Function CharExist(ByVal index As Long, ByVal CharNum As Long) As Boolean
    If Trim(Player(index).Char(CharNum).Name) <> "" Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim FileName As String
Dim RightPassword As String

    PasswordOK = False
    
    If AccountExist(Name) Then
        FileName = App.Path & "\accounts\" & Trim(Name) & ".ini"
        RightPassword = GetVar(FileName, "GENERAL", "Password")
        
        If UCase(Trim(Password)) = UCase(Trim(RightPassword)) Then
            PasswordOK = True
        End If
    End If
End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String)
Dim i As Long

    Player(index).Login = Name
    Player(index).Password = Password
    
    For i = 1 To MAX_CHARS
        Call ClearChar(index, i)
    Next i
    
    Call SavePlayer(index)
End Sub

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
Dim f As Long


    If Trim(Player(index).Char(CharNum).Name) = "" Then
        Player(index).CharNum = CharNum
        
        Player(index).Char(CharNum).Name = Name
        Player(index).Char(CharNum).Sex = Sex
        Player(index).Char(CharNum).Class = ClassNum
        
        If Player(index).Char(CharNum).Sex = SEX_MALE Then
            Player(index).Char(CharNum).Sprite = Class(ClassNum).Sprite
        Else
            Player(index).Char(CharNum).Sprite = Class(ClassNum).Sprite
        End If
        
        Player(index).Char(CharNum).Level = 1
                    
        Player(index).Char(CharNum).STR = Class(ClassNum).STR
        Player(index).Char(CharNum).DEF = Class(ClassNum).DEF
        Player(index).Char(CharNum).SPEED = Class(ClassNum).SPEED
        Player(index).Char(CharNum).MAGI = Class(ClassNum).MAGI
       
        Player(index).Char(CharNum).Map = START_MAP
        Player(index).Char(CharNum).x = START_X
        Player(index).Char(CharNum).y = START_Y
            
        Player(index).Char(CharNum).HP = GetPlayerMaxHP(index)
        Player(index).Char(CharNum).MP = GetPlayerMaxMP(index)
        Player(index).Char(CharNum).SP = GetPlayerMaxSP(index)
                
        ' Append name to file
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Append As #f
            Print #f, Name
        Close #f
        
        Call SavePlayer(index)
            
        Exit Sub
    End If
End Sub

Sub DelChar(ByVal index As Long, ByVal CharNum As Long)
Dim f1 As Long, f2 As Long
Dim s As String

    Call DeleteName(Player(index).Char(CharNum).Name)
    Call ClearChar(index, CharNum)
    Call SavePlayer(index)
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
Dim FileName As String
Dim i As Long

    Call CheckClasses
    
    FileName = App.Path & "\classes.ini"
    
    Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class(0 To Max_Classes) As ClassRec
    
    Call ClearClasses
    
    For i = 0 To Max_Classes
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).Sprite = GetVar(FileName, "CLASS" & i, "Sprite")
        Class(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class(i).SPEED = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class(i).MAGI = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
        
        DoEvents
    Next i
End Sub

Sub SaveClasses()
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\classes.ini"
    
    For i = 0 To Max_Classes
        Call PutVar(FileName, "CLASS" & i, "Name", Trim(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "Sprite", STR(Class(i).Sprite))
        Call PutVar(FileName, "CLASS" & i, "STR", STR(Class(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR(Class(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR(Class(i).SPEED))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR(Class(i).MAGI))
    Next i
End Sub

Sub CheckClasses()
    If Not FileExist("classes.ini") Then
        Call SaveClasses
    End If
End Sub

Sub SaveItems()
Dim i As Long
    
    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next i
End Sub

Sub SaveItem(ByVal ItemNum As Long)
    Dim FileName As String
    Dim f  As Long

    FileName = App.Path & "\items\item" & ItemNum & ".GXL"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Item(ItemNum)
    Close #f
End Sub

Sub LoadItems()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckItems

    For i = 1 To MAX_ITEMS
        Call SetStatus("Loading items... " & Int((i / MAX_ITEMS) * 100) & "%")
        FileName = App.Path & "\Items\Item" & i & ".GXL"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Item(i)
        Close #f

        DoEvents
    Next
End Sub

Sub CheckItems()
Dim i As Long
For i = 1 To MAX_ITEMS
    If Not FileExist("items\item" & i & ".GXL") Then
        Call SaveItems
    End If
Next
End Sub

Sub SaveShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next i
End Sub

Sub SaveShop(ByVal ShopNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\shops\shop" & ShopNum & ".GXL"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Shop(ShopNum)
    Close #f
End Sub

Sub LoadShops()
    Dim FileName As String
    Dim i As Long, f As Long

    Call CheckShops

    For i = 1 To MAX_SHOPS
        Call SetStatus("Loading shops... " & Int((i / MAX_SHOPS) * 100) & "%")
        FileName = App.Path & "\shops\shop" & i & ".GXL"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Shop(i)
        Close #f

        DoEvents
    Next
End Sub

Sub CheckShops()
Dim i As Long
For i = 1 To MAX_SHOPS
    If Not FileExist("shops\shop" & i & ".GXL") Then
        Call SaveShops
    End If
Next
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\spells\spells" & SpellNum & ".GXL"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Spell(SpellNum)
    Close #f
End Sub

Sub SaveSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next i
End Sub

Sub LoadSpells()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckSpells

    For i = 1 To MAX_SPELLS
        Call SetStatus("Loading spells... " & Int((i / MAX_SPELLS) * 100) & "%")
        FileName = App.Path & "\spells\spells" & i & ".GXL"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Spell(i)
        Close #f

        DoEvents
    Next

End Sub

Sub CheckSpells()
Dim i As Long
For i = 1 To MAX_SPELLS
    If Not FileExist("Spells\spells" & i & ".GXL") Then
        Call SaveSpells
    End If
Next
End Sub

Sub SaveNpcs()
Dim i As Long
    
    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next i
End Sub
Sub SaveNpc(ByVal npcnum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\npcs\npc" & npcnum & ".GXL"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Npc(npcnum)
    Close #f
End Sub

Sub LoadNpcs()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckNpcs

    For i = 1 To MAX_NPCS
        Call SetStatus("Loading npcs... " & Int((i / MAX_NPCS) * 100) & "%")
        FileName = App.Path & "\npcs\npc" & i & ".GXL"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Npc(i)
        Close #f

        DoEvents
    Next
End Sub

Sub CheckNpcs()
Dim i As Long
For i = 1 To MAX_NPCS
    If Not FileExist("NPCs\npc" & i & ".GXL") Then
        Call SaveNpcs
    End If
Next
End Sub

Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".MAP"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub SaveMaps()
Dim FileName As String
Dim i As Long
Dim f As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next i
End Sub

Sub LoadMaps()
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        FileName = App.Path & "\maps\map" & i & ".MAP"
        
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Map(i)
        Close #f
    
        DoEvents
    Next i
End Sub

Sub ConvertOldMapsToNew()
Dim FileName As String
Dim i As Long
Dim f As Long
Dim x As Long, y As Long
Dim OldMap As OldMapRec
Dim NewMap As MapRec

    For i = 1 To MAX_MAPS
        FileName = App.Path & "\maps\map" & i & ".GXL"
        
        ' Get the old file
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , OldMap
        Close #f
        
        ' Delete the old file
        Call Kill(FileName)
        
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
            For x = 0 To MAX_MAPX
                NewMap.Tile(x, y).Ground = OldMap.Tile(x, y).Ground
                NewMap.Tile(x, y).Mask = OldMap.Tile(x, y).Mask
                NewMap.Tile(x, y).Anim = OldMap.Tile(x, y).Anim
                NewMap.Tile(x, y).Fringe = OldMap.Tile(x, y).Fringe
                NewMap.Tile(x, y).Type = OldMap.Tile(x, y).Type
                NewMap.Tile(x, y).Data1 = OldMap.Tile(x, y).Data1
                NewMap.Tile(x, y).Data2 = OldMap.Tile(x, y).Data2
                NewMap.Tile(x, y).Data3 = OldMap.Tile(x, y).Data3
            Next x
        Next y
        
        For x = 1 To MAX_MAP_NPCS
            NewMap.Npc(x) = OldMap.Npc(x)
        Next x
        
        ' Set new values to 0 or null
        NewMap.Indoors = NO
        
        ' Save the new map
        f = FreeFile
        Open FileName For Binary As #f
            Put #f, , NewMap
        Close #f
    Next i
End Sub

Sub CheckMaps()
Dim FileName As String
Dim x As Long
Dim y As Long
Dim i As Long
Dim n As Long

    Call ClearMaps
        
    For i = 1 To MAX_MAPS
        FileName = "maps\map" & i & ".MAP"
        
        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(FileName) Then
            Call SaveMap(i)
        End If
    Next i
End Sub

Sub AddLog(ByVal Text As String, ByVal FN As String)
Dim FileName As String
Dim f As Long

    If ServerLog = True Then
        FileName = App.Path & "\" & FN
    
        If Not FileExist(FN) Then
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
Dim FileName, IP As String
Dim f As Long, i As Long

    FileName = App.Path & "\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
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
    Open FileName For Append As #f
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
