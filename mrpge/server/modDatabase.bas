Attribute VB_Name = "modDatabase"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Const START_MAP = 1
Public Const START_X = MAX_MAPX / 2
Public Const START_Y = MAX_MAPY / 2

Public Const ADMIN_LOG = "admin.txt"
Public Const PLAYER_LOG = "player.txt"

Function GetVar(file As String, header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(header, Var, szReturn, sSpaces, Len(sSpaces), file)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(file As String, header As String, Var As String, value As String)
    Call WritePrivateProfileString(header, Var, value, file)
End Sub

Function FileExist(ByVal filename As String) As Boolean
    If Dir(App.Path & "\" & filename) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Sub SavePlayer(ByVal index As Long, Optional ByVal blnAll As Boolean = True, Optional ByVal logoutCharNum As Long = -1)
Dim filename As String
Dim bioFilename As String
Dim i As Long
Dim n As Long
Dim loopMax As Long
Dim loopMin As Long
If blnAll Then
    loopMax = MAX_CHARS
    loopMin = 1
Else
    loopMin = player(index).CharNum
    loopMax = player(index).CharNum
End If
    filename = App.Path & "\accounts\" & Trim(player(index).Login) & ".ini"
    bioFilename = App.Path & "\accounts\bios\" & Trim(player(index).Login) & ".ini"
    Call PutVar(filename, "GENERAL", "Login", Trim(player(index).Login))
    Call PutVar(filename, "GENERAL", "Password", Trim(player(index).Password))
    Call PutVar(bioFilename, "GENERAL", "Name", Trim(player(index).RealName))
    Call PutVar(bioFilename, "GENERAL", "Email", Trim(player(index).Email))
    Call PutVar(bioFilename, "GENERAL", "Bio", Trim(player(index).Bio))
'no more loop tho all chars if not needed
    For i = loopMin To loopMax
        If loopMax = 1 Then
            i = player(index).CharNum
            If logoutCharNum > 0 Then
                i = logoutCharNum
            End If
        End If
        ' General
        Call PutVar(filename, "CHAR" & i, "Name", Trim(player(index).Char(i).Name))
        Call PutVar(filename, "CHAR" & i, "Class", str(player(index).Char(i).Class))
        Call PutVar(filename, "CHAR" & i, "Sex", str(player(index).Char(i).Sex))
        Call PutVar(filename, "CHAR" & i, "Sprite", str(player(index).Char(i).sprite))
        Call PutVar(filename, "CHAR" & i, "Level", str(player(index).Char(i).level))
        Call PutVar(filename, "CHAR" & i, "Exp", str(player(index).Char(i).Exp))
        Call PutVar(filename, "CHAR" & i, "Access", str(player(index).Char(i).Access))
        Call PutVar(filename, "CHAR" & i, "PK", str(player(index).Char(i).PK))
        Call PutVar(filename, "CHAR" & i, "Guild", str(player(index).Char(i).Guild))
        Call PutVar(filename, "CHAR" & i, "GuildAccess", str(player(index).Char(i).GuildAccess))
        Call PutVar(filename, "CHAR" & i, "txtColour", str(player(index).Char(i).txtColour))
        Call PutVar(filename, "CHAR" & i, "ingameColour", str(player(index).Char(i).ingameColour))
        Call PutVar(filename, "CHAR" & i, "CurrentQuest", str(player(index).Char(i).CurrentQuest))
        Call PutVar(filename, "CHAR" & i, "QuestStatus", str(player(index).Char(i).QuestStatus))
        
        
        ' Vitals
        Call PutVar(filename, "CHAR" & i, "HP", str(player(index).Char(i).HP))
        Call PutVar(filename, "CHAR" & i, "MP", str(player(index).Char(i).MP))
        Call PutVar(filename, "CHAR" & i, "SP", str(player(index).Char(i).SP))
        Call PutVar(filename, "CHAR" & i, "PP", str(player(index).Char(i).PP))
        
        ' Stats
        Call PutVar(filename, "CHAR" & i, "STR", str(player(index).Char(i).str))
        Call PutVar(filename, "CHAR" & i, "INTEL", str(player(index).Char(i).intel))
        Call PutVar(filename, "CHAR" & i, "DEX", str(player(index).Char(i).dex))
        Call PutVar(filename, "CHAR" & i, "CON", str(player(index).Char(i).con))
        Call PutVar(filename, "CHAR" & i, "WIZ", str(player(index).Char(i).wiz))
        Call PutVar(filename, "CHAR" & i, "CHA", str(player(index).Char(i).cha))
        'old stuff
        'Call PutVar(filename, "CHAR" & i, "DEF", STR(Player(index).Char(i).DEF))
        'Call PutVar(filename, "CHAR" & i, "SPEED", STR(Player(index).Char(i).speed))
        'Call PutVar(filename, "CHAR" & i, "MAGI", STR(Player(index).Char(i).MAGI))
        Call PutVar(filename, "CHAR" & i, "POINTS", str(player(index).Char(i).POINTS))
        
        ' Worn equipment
        Call PutVar(filename, "CHAR" & i, "ArmorSlot", str(player(index).Char(i).ArmorSlot))
        Call PutVar(filename, "CHAR" & i, "WeaponSlot", str(player(index).Char(i).WeaponSlot))
        Call PutVar(filename, "CHAR" & i, "HelmetSlot", str(player(index).Char(i).HelmetSlot))
        Call PutVar(filename, "CHAR" & i, "ShieldSlot", str(player(index).Char(i).ShieldSlot))
        
        'If player(index).Char(i).PetId <> 0 Then
        '    Call PutVar(filename, "CHAR" & i, "HasPet", "true")
        'Else
        '    Call PutVar(filename, "CHAR" & i, "HasPet", "false")
        'End If
        
        'Call PutVar(filename, "CHAR" & i, "PetName", Pets(player(index).Char(i).PetId).Name)
        'Call PutVar(filename, "CHAR" & i, "PetSprite", Pets(player(index).Char(i).PetId).sprite)
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If player(index).Char(i).map = 0 Then
            player(index).Char(i).map = START_MAP
            player(index).Char(i).x = START_X
            player(index).Char(i).y = START_Y
        End If
            
        ' Position
        Call PutVar(filename, "CHAR" & i, "Map", str(player(index).Char(i).map))
        Call PutVar(filename, "CHAR" & i, "X", str(player(index).Char(i).x))
        Call PutVar(filename, "CHAR" & i, "Y", str(player(index).Char(i).y))
        Call PutVar(filename, "CHAR" & i, "Dir", str(player(index).Char(i).Dir))
        
        ' Inventory
        For n = 1 To MAX_INV
            Call PutVar(filename, "CHAR" & i, "InvItemNum" & n, str(player(index).Char(i).Inv(n).num))
            Call PutVar(filename, "CHAR" & i, "InvItemVal" & n, str(player(index).Char(i).Inv(n).value))
            Call PutVar(filename, "CHAR" & i, "InvItemDur" & n, str(player(index).Char(i).Inv(n).Dur))
        Next n
        For n = 1 To MAX_BANK
            Call PutVar(filename, "CHAR" & i, "BankItemNum" & n, str(player(index).Char(i).Bank(n).num))
            Call PutVar(filename, "CHAR" & i, "BankItemVal" & n, str(player(index).Char(i).Bank(n).value))
            Call PutVar(filename, "CHAR" & i, "BankItemDur" & n, str(player(index).Char(i).Bank(n).Dur))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Call PutVar(filename, "CHAR" & i, "Spell" & n, str(player(index).Char(i).Spell(n)))
            Call PutVar(filename, "CHAR" & i, "Prayer" & n, str(player(index).Char(i).Prayer(n)))
        Next n
    Next i
End Sub

Sub SavePlayerBank(ByVal index As Long)
    Dim filename As String
    Dim i As Long
    Dim n As Long
    filename = App.Path & "\accounts\" & Trim(player(index).Login) & ".ini"
    i = player(index).CharNum
    For n = 1 To MAX_INV
        Call PutVar(filename, "CHAR" & i, "InvItemNum" & n, str(player(index).Char(i).Inv(n).num))
        Call PutVar(filename, "CHAR" & i, "InvItemVal" & n, str(player(index).Char(i).Inv(n).value))
        Call PutVar(filename, "CHAR" & i, "InvItemDur" & n, str(player(index).Char(i).Inv(n).Dur))
        'DoEvents
    Next n
    DoEvents
    For n = 1 To MAX_BANK
        Call PutVar(filename, "CHAR" & i, "BankItemNum" & n, str(player(index).Char(i).Bank(n).num))
        Call PutVar(filename, "CHAR" & i, "BankItemVal" & n, str(player(index).Char(i).Bank(n).value))
        Call PutVar(filename, "CHAR" & i, "BankItemDur" & n, str(player(index).Char(i).Bank(n).Dur))
        'DoEvents
    Next n
    DoEvents
End Sub


Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
Dim filename As String
Dim bioFilename As String

Dim i As Long
Dim n As Long

    Call ClearPlayer(index)
    
    filename = App.Path & "\accounts\" & Trim(Name) & ".ini"
    bioFilename = App.Path & "\accounts\bios\" & Trim(Name) & ".ini"
    
    player(index).Login = GetVar(filename, "GENERAL", "Login")
    player(index).Password = GetVar(filename, "GENERAL", "Password")
    player(index).RealName = GetVar(bioFilename, "GENERAL", "Name")
    player(index).Email = GetVar(bioFilename, "GENERAL", "Email")
    player(index).Bio = GetVar(bioFilename, "GENERAL", "Bio")

    For i = 1 To MAX_CHARS
        ' General
        player(index).Char(i).Name = GetVar(filename, "CHAR" & i, "Name")
        player(index).Char(i).Sex = Val(GetVar(filename, "CHAR" & i, "Sex"))
        player(index).Char(i).Class = Val(GetVar(filename, "CHAR" & i, "Class"))
        player(index).Char(i).sprite = Val(GetVar(filename, "CHAR" & i, "Sprite"))
        player(index).Char(i).level = Val(GetVar(filename, "CHAR" & i, "Level"))
        player(index).Char(i).Exp = Val(GetVar(filename, "CHAR" & i, "Exp"))
        player(index).Char(i).Access = Val(GetVar(filename, "CHAR" & i, "Access"))
        player(index).Char(i).PK = Val(GetVar(filename, "CHAR" & i, "PK"))
        player(index).Char(i).Guild = Val(GetVar(filename, "CHAR" & i, "Guild"))
        player(index).Char(i).GuildAccess = Val(GetVar(filename, "CHAR" & i, "GuildAccess"))
        player(index).Char(i).txtColour = Val(GetVar(filename, "CHAR" & i, "txtColour"))
        player(index).Char(i).ingameColour = Val(GetVar(filename, "CHAR" & i, "ingameColour"))
        player(index).Char(i).QuestStatus = Val(GetVar(filename, "CHAR" & i, "QuestStatus"))
        player(index).Char(i).CurrentQuest = Val(GetVar(filename, "CHAR" & i, "CurrentQuest"))
        
        ' Vitals
        player(index).Char(i).HP = Val(GetVar(filename, "CHAR" & i, "HP"))
        player(index).Char(i).MP = Val(GetVar(filename, "CHAR" & i, "MP"))
        player(index).Char(i).SP = Val(GetVar(filename, "CHAR" & i, "SP"))
        player(index).Char(i).PP = Val(GetVar(filename, "CHAR" & i, "PP"))
        
        ' Stats
        player(index).Char(i).str = Val(GetVar(filename, "CHAR" & i, "STR"))
        player(index).Char(i).intel = Val(GetVar(filename, "CHAR" & i, "INTEL"))
        player(index).Char(i).dex = Val(GetVar(filename, "CHAR" & i, "DEX"))
        player(index).Char(i).con = Val(GetVar(filename, "CHAR" & i, "CON"))
        player(index).Char(i).wiz = Val(GetVar(filename, "CHAR" & i, "WIZ"))
        player(index).Char(i).cha = Val(GetVar(filename, "CHAR" & i, "CHA"))
        'old stuff
        'Player(index).Char(i).DEF = Val(GetVar(filename, "CHAR" & i, "DEF"))
        'Player(index).Char(i).speed = Val(GetVar(filename, "CHAR" & i, "SPEED"))
        'Player(index).Char(i).MAGI = Val(GetVar(filename, "CHAR" & i, "MAGI"))
        player(index).Char(i).POINTS = Val(GetVar(filename, "CHAR" & i, "POINTS"))
        
        ' Worn equipment
        player(index).Char(i).ArmorSlot = Val(GetVar(filename, "CHAR" & i, "ArmorSlot"))
        player(index).Char(i).WeaponSlot = Val(GetVar(filename, "CHAR" & i, "WeaponSlot"))
        player(index).Char(i).HelmetSlot = Val(GetVar(filename, "CHAR" & i, "HelmetSlot"))
        player(index).Char(i).ShieldSlot = Val(GetVar(filename, "CHAR" & i, "ShieldSlot"))
        
        ' Position
        player(index).Char(i).map = Val(GetVar(filename, "CHAR" & i, "Map"))
        player(index).Char(i).x = Val(GetVar(filename, "CHAR" & i, "X"))
        player(index).Char(i).y = Val(GetVar(filename, "CHAR" & i, "Y"))
        player(index).Char(i).Dir = Val(GetVar(filename, "CHAR" & i, "Dir"))
        'player(index).Char(i).HasPet = GetVar(filename, "CHAR" & i, "HasPet")
        'player(index).Char(i).PetName = GetVar(filename, "CHAR" & i, "PetName")
        'player(index).Char(i).PetSprite = GetVar(filename, "CHAR" & i, "PetSprite")
        
        
'        'load its pet, if it has one.
'        If GetVar(filename, "CHAR" & i, "HasPet") = "true" Then
'            player(index).Char(i).PetId = GetFreePetID
'            Pets(player(index).Char(i).PetId).owner = index
'            Call loadPet(player(index).Char(i).PetId, GetVar(filename, "CHAR" & i, "PetName"), GetVar(filename, "CHAR" & i, "PetSprite"))
'        Else
'            player(index).Char(i).PetId = 0
'        End If
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If player(index).Char(i).map = 0 Then
            player(index).Char(i).map = START_MAP
            player(index).Char(i).x = START_X
            player(index).Char(i).y = START_Y
        End If
        
        ' Inventory
        For n = 1 To MAX_INV
            player(index).Char(i).Inv(n).num = Val(GetVar(filename, "CHAR" & i, "InvItemNum" & n))
            player(index).Char(i).Inv(n).value = Val(GetVar(filename, "CHAR" & i, "InvItemVal" & n))
            player(index).Char(i).Inv(n).Dur = Val(GetVar(filename, "CHAR" & i, "InvItemDur" & n))
        Next n
        
        For n = 1 To MAX_BANK
        'Debug.Print "itemNum: " & GetVar(filename, "CHAR" & i, "BankItemNum" & n)
            player(index).Char(i).Bank(n).num = Val(GetVar(filename, "CHAR" & i, "BankItemNum" & n))
            player(index).Char(i).Bank(n).value = Val(GetVar(filename, "CHAR" & i, "BankItemVal" & n))
            player(index).Char(i).Bank(n).Dur = Val(GetVar(filename, "CHAR" & i, "BankItemDur" & n))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            player(index).Char(i).Spell(n) = Val(GetVar(filename, "CHAR" & i, "Spell" & n))
            player(index).Char(i).Prayer(n) = Val(GetVar(filename, "CHAR" & i, "Prayer" & n))
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

Function CharExist(ByVal index As Long, ByVal CharNum As Long) As Boolean
    If Trim(player(index).Char(CharNum).Name) <> "" Then
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

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String)
Dim i As Long

    player(index).Login = Name
    player(index).Password = Password
    
    For i = 1 To MAX_CHARS
        Call ClearChar(index, i)
    Next i
    
    Call SavePlayer(index, True)
End Sub

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long, ByVal sprite As Long)
Dim f As Long

    If Trim(player(index).Char(CharNum).Name) = "" Then
        player(index).CharNum = CharNum
        
        player(index).Char(CharNum).Name = Name
        player(index).Char(CharNum).Sex = Sex
        player(index).Char(CharNum).Class = ClassNum
        
        If player(index).Char(CharNum).Sex = SEX_MALE Then
            player(index).Char(CharNum).sprite = sprite 'Class(ClassNum).sprite
        Else
            player(index).Char(CharNum).sprite = sprite 'Class(ClassNum).sprite
        End If
        
        player(index).Char(CharNum).level = 1
                    
        player(index).Char(CharNum).str = Class(ClassNum).str
        'Player(index).Char(CharNum).DEF = Class(ClassNum).DEF
        'Player(index).Char(CharNum).speed = Class(ClassNum).speed
        'Player(index).Char(CharNum).MAGI = Class(ClassNum).MAGI
        player(index).Char(CharNum).intel = Class(ClassNum).intel
        player(index).Char(CharNum).dex = Class(ClassNum).dex
        player(index).Char(CharNum).con = Class(ClassNum).con
        player(index).Char(CharNum).wiz = Class(ClassNum).wiz
        player(index).Char(CharNum).cha = Class(ClassNum).cha
        
        player(index).Char(CharNum).map = START_MAP
        player(index).Char(CharNum).x = START_X
        player(index).Char(CharNum).y = START_Y
            
        player(index).Char(CharNum).HP = GetPlayerMaxHP(index)
        player(index).Char(CharNum).MP = GetPlayerMaxMP(index)
        player(index).Char(CharNum).SP = GetPlayerMaxSP(index)
        
        'txt colours
        player(index).Char(CharNum).txtColour = 15
        player(index).Char(CharNum).ingameColour = 15
                
        ' Append name to file
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Append As #f
            Print #f, Name
        Close #f
        'MARKER
        Call SavePlayer(index, False)
            
        Exit Sub
    End If
End Sub

Sub DelChar(ByVal index As Long, ByVal CharNum As Long)
Dim f1 As Long, f2 As Long
Dim s As String

    Call DeleteName(player(index).Char(CharNum).Name)
    Call ClearChar(index, CharNum)
    Call SavePlayer(index, False, CharNum)
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
            Call SavePlayer(i, False)
        End If
    Next i
End Sub

Sub LoadClasses()
Dim filename As String
Dim i As Long

    Call CheckClasses
    
    filename = App.Path & "\classes.ini"
    
    Max_Classes = Val(GetVar(filename, "INIT", "MaxClasses"))
    
    ReDim Class(0 To Max_Classes) As ClassRec
    
    Call ClearClasses
    
    For i = 0 To Max_Classes
        Class(i).Name = GetVar(filename, "CLASS" & i, "Name")
        Class(i).sprite = GetVar(filename, "CLASS" & i, "Sprite")
        Class(i).str = Val(GetVar(filename, "CLASS" & i, "STR"))
        'Class(i).DEF = Val(GetVar(filename, "CLASS" & i, "DEF"))
        'Class(i).speed = Val(GetVar(filename, "CLASS" & i, "SPEED"))
        'Class(i).MAGI = Val(GetVar(filename, "CLASS" & i, "MAGI"))
        Class(i).intel = Val(GetVar(filename, "CLASS" & i, "INTEL"))
        Class(i).dex = Val(GetVar(filename, "CLASS" & i, "DEX"))
        Class(i).con = Val(GetVar(filename, "CLASS" & i, "CON"))
        Class(i).wiz = Val(GetVar(filename, "CLASS" & i, "WIZ"))
        Class(i).cha = Val(GetVar(filename, "CLASS" & i, "CHA"))
        DoEvents
    Next i
End Sub

Sub SaveClasses()
Dim filename As String
Dim i As Long

    filename = App.Path & "\classes.ini"
    
    For i = 0 To Max_Classes
        Call PutVar(filename, "CLASS" & i, "Name", Trim(Class(i).Name))
        Call PutVar(filename, "CLASS" & i, "Sprite", str(Class(i).sprite))
        Call PutVar(filename, "CLASS" & i, "STR", str(Class(i).str))
        'Call PutVar(filename, "CLASS" & i, "DEF", STR(Class(i).DEF))
        'Call PutVar(filename, "CLASS" & i, "SPEED", STR(Class(i).speed))
        'Call PutVar(filename, "CLASS" & i, "MAGI", STR(Class(i).MAGI))
        Call PutVar(filename, "CLASS" & i, "INTEL", str(Class(i).intel))
        Call PutVar(filename, "CLASS" & i, "DEX", str(Class(i).dex))
        Call PutVar(filename, "CLASS" & i, "CON", str(Class(i).con))
        Call PutVar(filename, "CLASS" & i, "WIZ", str(Class(i).wiz))
        Call PutVar(filename, "CLASS" & i, "CHA", str(Class(i).cha))
        
    Next i
End Sub

Sub CheckClasses()
    If Not FileExist("classes.ini") Then
        Call SaveClasses
    End If
End Sub

Sub saveItems()
Dim i As Long
    
    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next i
End Sub

Sub SaveItem(ByVal itemnum As Long)
Dim filename As String

    filename = App.Path & "\items.ini"
    
    Call PutVar(filename, "ITEM" & itemnum, "Name", Trim(Item(itemnum).Name))
    Call PutVar(filename, "ITEM" & itemnum, "Pic", Trim(Item(itemnum).Pic))
    Call PutVar(filename, "ITEM" & itemnum, "Type", Trim(Item(itemnum).type))
    Call PutVar(filename, "ITEM" & itemnum, "Data1", Trim(Item(itemnum).Data1))
    Call PutVar(filename, "ITEM" & itemnum, "Data2", Trim(Item(itemnum).Data2))
    Call PutVar(filename, "ITEM" & itemnum, "Data3", Trim(Item(itemnum).Data3))
    
    Call PutVar(filename, "ITEM" & itemnum, "BaseDamage", Trim(Item(itemnum).BaseDamage))
    Call PutVar(filename, "ITEM" & itemnum, "str", Trim(Item(itemnum).str))
    Call PutVar(filename, "ITEM" & itemnum, "intel", Trim(Item(itemnum).intel))
    Call PutVar(filename, "ITEM" & itemnum, "dex", Trim(Item(itemnum).dex))
    Call PutVar(filename, "ITEM" & itemnum, "con", Trim(Item(itemnum).con))
    Call PutVar(filename, "ITEM" & itemnum, "wiz", Trim(Item(itemnum).wiz))
    Call PutVar(filename, "ITEM" & itemnum, "cha", Trim(Item(itemnum).cha))
    Call PutVar(filename, "ITEM" & itemnum, "description", Trim(Item(itemnum).Description))
    Call PutVar(filename, "ITEM" & itemnum, "Poison_length", Trim(Item(itemnum).Poison_length))
    Call PutVar(filename, "ITEM" & itemnum, "Poisons", Trim(Item(itemnum).Poisons))
    Call PutVar(filename, "ITEM" & itemnum, "Poison_vital", Trim(Item(itemnum).Poison_vital))
    Call PutVar(filename, "ITEM" & itemnum, "WeaponType", Trim(Item(itemnum).weaponType))
End Sub

Sub LoadItems()
On Error Resume Next
Dim filename As String
Dim i As Long

    Call CheckItems
    
    filename = App.Path & "\items.ini"
    
    For i = 1 To MAX_ITEMS
        Item(i).Name = GetVar(filename, "ITEM" & i, "Name")
        Item(i).Pic = Val(GetVar(filename, "ITEM" & i, "Pic"))
        Item(i).type = Val(GetVar(filename, "ITEM" & i, "Type"))
        Item(i).Data1 = Val(GetVar(filename, "ITEM" & i, "Data1"))
        Item(i).Data2 = Val(GetVar(filename, "ITEM" & i, "Data2"))
        Item(i).Data3 = Val(GetVar(filename, "ITEM" & i, "Data3"))
        
        Item(i).BaseDamage = Val(GetVar(filename, "ITEM" & i, "BaseDamage"))
        Item(i).str = Val(GetVar(filename, "ITEM" & i, "str"))
        Item(i).intel = Val(GetVar(filename, "ITEM" & i, "intel"))
        Item(i).dex = Val(GetVar(filename, "ITEM" & i, "dex"))
        Item(i).con = Val(GetVar(filename, "ITEM" & i, "con"))
        Item(i).wiz = Val(GetVar(filename, "ITEM" & i, "wiz"))
        Item(i).cha = Val(GetVar(filename, "ITEM" & i, "cha"))
        Item(i).Description = GetVar(filename, "ITEM" & i, "description")
        'Debug.Print GetVar(filename, "ITEM" & i, "Poison_length")
        Item(i).Poison_length = GetVar(filename, "ITEM" & i, "Poison_length")
        Item(i).Poisons = GetVar(filename, "ITEM" & i, "Poisons")
        Item(i).Poison_vital = GetVar(filename, "ITEM" & i, "Poison_vital")
        Item(i).weaponType = Val(GetVar(filename, "ITEM" & i, "weaponType"))
        
        DoEvents
    Next i
End Sub

Sub LoadGuilds()
Dim filename As String
Dim i As Long
Dim y As Long

    Call CheckGuilds
    
    filename = App.Path & "\guilds.ini"
    
    For i = 1 To MAX_GUILDS
        Guild(i).Founder = GetVar(filename, "GUILD" & i, "Founder")
        Guild(i).Member(1) = GetVar(filename, "GUILD" & i, "Member")
        Guild(i).Description = GetVar(filename, "GUILD" & i, "Description")
        Guild(i).Name = GetVar(filename, "GUILD" & i, "Name")
        Guild(i).InviteList = GetVar(filename, "GUILD" & i, "InviteList")
        For y = 1 To MAX_GUILD_MEMBERS
            Guild(i).Member(y) = GetVar(filename, "GUILD" & i, "Member" & y)
            Guild(i).Leaders(y) = GetVar(filename, "GUILD" & i, "Leaders" & y)
        Next y
        DoEvents
    Next i
End Sub

Sub SaveGuild(ByVal GuildNum As Long)
Dim filename As String
Dim i As Long
    
    filename = App.Path & "\guilds.ini"
    If Guild(GuildNum).Founder = "" Then Guild(GuildNum).Founder = Chr(0)
    PutVar filename, "GUILD" & GuildNum, "Founder", Guild(GuildNum).Founder
    If Guild(GuildNum).Name = "" Then Guild(GuildNum).Name = Chr(0)
    PutVar filename, "GUILD" & GuildNum, "Name", Guild(GuildNum).Name
    If Guild(GuildNum).Description = "" Then Guild(GuildNum).Description = Chr(0)
    PutVar filename, "GUILD" & GuildNum, "Description", Guild(GuildNum).Description
    If Guild(GuildNum).InviteList = "" Then Guild(GuildNum).InviteList = Chr(0)
    PutVar filename, "GUILD" & GuildNum, "InviteList", Guild(GuildNum).InviteList
    For i = 1 To MAX_GUILD_MEMBERS
        If Guild(GuildNum).Member(i) = "" Then Guild(GuildNum).Member(i) = Chr(0)
        PutVar filename, "GUILD" & GuildNum, "Member" & i, Guild(GuildNum).Member(i)
        If Guild(GuildNum).Leaders(i) = "" Then Guild(GuildNum).Leaders(i) = Chr(0)
        PutVar filename, "GUILD" & GuildNum, "Leader" & i, Guild(GuildNum).Leaders(i)
    Next i
    DoEvents

End Sub

Sub SaveGuilds()
Dim i As Long

    For i = 1 To MAX_GUILDS
        Call SaveGuild(i)
    Next i
End Sub

Sub CheckGuilds()
    If Not FileExist("guilds.ini") Then
        Call SaveGuilds
    End If
End Sub

Sub CheckItems()
    If Not FileExist("items.ini") Then
        Call saveItems
    End If
End Sub


Sub SaveShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next i
End Sub

Sub SaveShop(ByVal ShopNum As Long)
Dim filename As String
Dim i As Long

    filename = App.Path & "\shops.ini"
    
    Call PutVar(filename, "SHOP" & ShopNum, "Name", Trim(Shop(ShopNum).Name))
    Call PutVar(filename, "SHOP" & ShopNum, "JoinSay", Trim(Shop(ShopNum).JoinSay))
    Call PutVar(filename, "SHOP" & ShopNum, "LeaveSay", Trim(Shop(ShopNum).LeaveSay))
    Call PutVar(filename, "SHOP" & ShopNum, "FixesItems", Trim(Shop(ShopNum).FixesItems))
    
    For i = 1 To MAX_TRADES
        Call PutVar(filename, "SHOP" & ShopNum, "GiveItem" & i, Trim(Shop(ShopNum).TradeItem(i).GiveItem))
        Call PutVar(filename, "SHOP" & ShopNum, "GiveValue" & i, Trim(Shop(ShopNum).TradeItem(i).GiveValue))
        Call PutVar(filename, "SHOP" & ShopNum, "GetItem" & i, Trim(Shop(ShopNum).TradeItem(i).GetItem))
        Call PutVar(filename, "SHOP" & ShopNum, "GetValue" & i, Trim(Shop(ShopNum).TradeItem(i).GetValue))
    Next i
End Sub

Sub LoadShops()
On Error Resume Next

Dim filename As String
Dim x As Long, y As Long

    Call CheckShops
    
    filename = App.Path & "\shops.ini"
    
    For y = 1 To MAX_SHOPS
        Shop(y).Name = GetVar(filename, "SHOP" & y, "Name")
        Shop(y).JoinSay = GetVar(filename, "SHOP" & y, "JoinSay")
        Shop(y).LeaveSay = GetVar(filename, "SHOP" & y, "LeaveSay")
        Shop(y).FixesItems = GetVar(filename, "SHOP" & y, "FixesItems")
        
        For x = 1 To MAX_TRADES
            Shop(y).TradeItem(x).GiveItem = GetVar(filename, "SHOP" & y, "GiveItem" & x)
            Shop(y).TradeItem(x).GiveValue = GetVar(filename, "SHOP" & y, "GiveValue" & x)
            Shop(y).TradeItem(x).GetItem = GetVar(filename, "SHOP" & y, "GetItem" & x)
            Shop(y).TradeItem(x).GetValue = GetVar(filename, "SHOP" & y, "GetValue" & x)
        Next x
    
        DoEvents
    Next y
End Sub

Sub CheckShops()
    If Not FileExist("shops.ini") Then
        Call SaveShops
    End If
End Sub

Sub CheckPrayers()
    If Not FileExist("prayer.ini") Then
        Call SavePrayers
    End If
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
Dim filename As String
Dim i As Long

    filename = App.Path & "\spells.ini"
    
    Call PutVar(filename, "SPELL" & SpellNum, "Name", Trim(Spell(SpellNum).Name))
    Call PutVar(filename, "SPELL" & SpellNum, "ClassReq", Trim(Spell(SpellNum).ClassReq))
    Call PutVar(filename, "SPELL" & SpellNum, "Type", Trim(Spell(SpellNum).type))
    Call PutVar(filename, "SPELL" & SpellNum, "Data1", Trim(Spell(SpellNum).Data1))
    Call PutVar(filename, "SPELL" & SpellNum, "Data2", Trim(Spell(SpellNum).Data2))
    Call PutVar(filename, "SPELL" & SpellNum, "Data3", Trim(Spell(SpellNum).Data3))
    Call PutVar(filename, "SPELL" & SpellNum, "Sound", Trim(Spell(SpellNum).sound))
    Call PutVar(filename, "SPELL" & SpellNum, "ManaUse", Trim(Spell(SpellNum).manaUse))
End Sub

Sub SaveQuest(ByVal QuestID As Long)
Dim filename As String
Dim startmsg As String
Dim midmsg As String
Dim endmsg As String

Dim i As Long
'i = QuestID
    filename = App.Path & "\quests.ini"
    
    For i = 1 To Len(Quests(QuestID).GetItemQuestMsg) Step 1
        Debug.Print i & ": #" & Asc(Mid(Quests(QuestID).GetItemQuestMsg, i)) & "#"
        'MsgBox Asc(Mid(Quests(i).GetItemQuestMsg, i, 1))
    Next i
    
    'startmsg = Replace(Quests(QuestID).StartQuestMsg, vbCrLf, "/n")
    'midmsg = Replace(Quests(QuestID).GetItemQuestMsg, vbCrLf, "/n")
    'endmsg = Replace(Quests(QuestID).FinishQuestMessage, vbCrLf, "/n")
    'startmsg = Replace(Quests(QuestID).StartQuestMsg, vbCr, "/n")
    'midmsg = Replace(Quests(QuestID).GetItemQuestMsg, vbCr, "/n")
    'endmsg = Replace(Quests(QuestID).FinishQuestMessage, vbCr, "/n")
    'startmsg = Replace(Quests(QuestID).StartQuestMsg, vbLf, "/n")
    'midmsg = Replace(Quests(QuestID).GetItemQuestMsg, vbLf, "/n")
    'endmsg = Replace(Quests(QuestID).FinishQuestMessage, vbLf, "/n")
    startmsg = Replace(Quests(QuestID).StartQuestMsg, Chr(13), "/n", , , vbTextCompare)
    midmsg = Replace(Quests(QuestID).GetItemQuestMsg, Chr(13), "/n", , , vbTextCompare)
    endmsg = Replace(Quests(QuestID).FinishQuestMessage, Chr(13), "/n", , , vbTextCompare)
    
    startmsg = Replace(startmsg, Chr(10), "/n", , , vbTextCompare)
    midmsg = Replace(midmsg, Chr(10), "/n", , , vbTextCompare)
    endmsg = Replace(endmsg, Chr(10), "/n", , , vbTextCompare)
    
    'startmsg = Replace(startmsg, Chr(97), "/n", , , vbTextCompare)
    'midmsg = Replace(midmsg, Chr(97), "/n", , , vbTextCompare)
    'endmsg = Replace(endmsg, Chr(97), "/n", , , vbTextCompare)
'    For i = 0 To 50 Step 1
'        startmsg = Replace(Quests(QuestID).StartQuestMsg, Chr(i), "/n")
'        midmsg = Replace(Quests(QuestID).GetItemQuestMsg, Chr(i), "/n")
'        endmsg = Replace(Quests(QuestID).FinishQuestMessage, Chr(i), "/n")
'    Next i
'    MsgBox Asc(Mid(startmsg, 1, 1)) & " - " & Mid(startmsg, 2, 1)
'    MsgBox Asc(Mid(startmsg, 2, 1)) & " - " & Mid(startmsg, 3, 1)
    Call PutVar(filename, "QUEST" & QuestID, "EXPGIVEN", Trim(Quests(QuestID).ExpGiven))
    Call PutVar(filename, "QUEST" & QuestID, "FINISHMSG", Trim(endmsg))
    Call PutVar(filename, "QUEST" & QuestID, "MIDDLEMSG", Trim(midmsg))
    Call PutVar(filename, "QUEST" & QuestID, "GETITEM", Trim(Quests(QuestID).ItemGiven))
    Call PutVar(filename, "QUEST" & QuestID, "FINDITEM", Trim(Quests(QuestID).ItemToObtain))
    Call PutVar(filename, "QUEST" & QuestID, "GETITEMVAL", Trim(Quests(QuestID).ItemValGiven))
    Call PutVar(filename, "QUEST" & QuestID, "REQUIREDLEVEL", Trim(Quests(QuestID).requiredLevel))
    Call PutVar(filename, "QUEST" & QuestID, "STARTMSG", Trim(startmsg))
    Call PutVar(filename, "QUEST" & QuestID, "GOLDGIVEN", Trim(Quests(QuestID).goldGiven))
    
End Sub
Sub SaveQuests()
Dim i As Long

    For i = 1 To MAX_QUESTS
        Call SaveQuest(i)
    Next i
End Sub



Sub SaveSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next i
End Sub

Sub SavePrayer(ByVal PrayerNum As Long)
Dim filename As String
Dim i As Long

    filename = App.Path & "\Prayer.ini"
    
    Call PutVar(filename, "PRAYER" & PrayerNum, "Name", Trim(Prayer(PrayerNum).Name))
    Call PutVar(filename, "PRAYER" & PrayerNum, "ClassReq", Trim(Prayer(PrayerNum).ClassReq))
    Call PutVar(filename, "PRAYER" & PrayerNum, "Type", Trim(Prayer(PrayerNum).type))
    Call PutVar(filename, "PRAYER" & PrayerNum, "Data1", Trim(Prayer(PrayerNum).Data1))
    Call PutVar(filename, "PRAYER" & PrayerNum, "Data2", Trim(Prayer(PrayerNum).Data2))
    Call PutVar(filename, "PRAYER" & PrayerNum, "Data3", Trim(Prayer(PrayerNum).Data3))
    Call PutVar(filename, "PRAYER" & PrayerNum, "Sound", Trim(Prayer(PrayerNum).sound))
    Call PutVar(filename, "PRAYER" & PrayerNum, "ManaUse", Trim(Prayer(PrayerNum).manaUse))
End Sub

Sub SavePrayers()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call SavePrayer(i)
    Next i
End Sub

Sub SaveSign(ByVal SignNum As Long)
Dim filename As String
Dim i As Long

    filename = App.Path & "\signs.ini"
    Signs(SignNum).msg = Replace(Trim(Signs(SignNum).msg), Chr(13), "/n")
    Signs(SignNum).msg = Replace(Trim(Signs(SignNum).msg), Chr(10), "/n")
    
    
    Call PutVar(filename, "SIGN" & SignNum, "Header", Trim(Signs(SignNum).header))
    Call PutVar(filename, "SIGN" & SignNum, "Msg", Trim(Signs(SignNum).msg))
    Signs(SignNum).msg = Replace(Trim(Signs(SignNum).msg), "/n", vbCrLf)
End Sub

Sub SaveBio(ByVal index As Long)
Dim bioFilename As String

    bioFilename = App.Path & "\accounts\bios\" & Trim(player(index).Login) & ".ini"
    
    Call PutVar(bioFilename, "GENERAL", "Name", Trim(player(index).RealName))
    Call PutVar(bioFilename, "GENERAL", "Email", Trim(player(index).Email))
    Call PutVar(bioFilename, "GENERAL", "Bio", Trim(player(index).Bio))
End Sub

Sub LoadBio(ByVal index As Long)
Dim bioFilename As String

    bioFilename = App.Path & "\accounts\bios\" & Trim(player(index).Login) & ".ini"
    
    player(index).RealName = GetVar(bioFilename, "GENERAL", "Name")
    player(index).Email = GetVar(bioFilename, "GENERAL", "Email")
    player(index).Bio = GetVar(bioFilename, "GENERAL", "Bio")
End Sub

Sub SaveSigns()
Dim i As Long

    For i = 1 To MAX_SIGNS
        Call SaveSign(i)
    Next i
End Sub

Sub LoadSpells()
Dim filename As String
Dim i As Long

    Call CheckSpells
    
    filename = App.Path & "\spells.ini"
    
    For i = 1 To MAX_SPELLS
        Spell(i).Name = GetVar(filename, "SPELL" & i, "Name")
        Spell(i).ClassReq = Val(GetVar(filename, "SPELL" & i, "ClassReq"))
        Spell(i).type = Val(GetVar(filename, "SPELL" & i, "Type"))
        Spell(i).Data1 = Val(GetVar(filename, "SPELL" & i, "Data1"))
        Spell(i).Data2 = Val(GetVar(filename, "SPELL" & i, "Data2"))
        Spell(i).Data3 = Val(GetVar(filename, "SPELL" & i, "Data3"))
        Spell(i).sound = Val(GetVar(filename, "SPELL" & i, "Sound"))
        Spell(i).manaUse = Val(GetVar(filename, "SPELL" & i, "ManaUse"))
        DoEvents
    Next i
End Sub

Sub LoadQuests()
Dim filename As String
Dim i As Long

    Call CheckQuests
    
    filename = App.Path & "\quests.ini"
    
    For i = 1 To MAX_QUESTS
        Quests(i).ID = i
        Quests(i).ExpGiven = Val(GetVar(filename, "QUEST" & i, "EXPGIVEN"))
        Quests(i).FinishQuestMessage = (GetVar(filename, "QUEST" & i, "FINISHMSG"))
        Quests(i).GetItemQuestMsg = (GetVar(filename, "QUEST" & i, "MIDDLEMSG"))
        Quests(i).ItemGiven = Val(GetVar(filename, "QUEST" & i, "GETITEM"))
        Quests(i).ItemToObtain = Val(GetVar(filename, "QUEST" & i, "FINDITEM"))
        Quests(i).ItemValGiven = Val(GetVar(filename, "QUEST" & i, "GETITEMVAL"))
        Quests(i).requiredLevel = Val(GetVar(filename, "QUEST" & i, "REQUIREDLEVEL"))
        Quests(i).StartQuestMsg = (GetVar(filename, "QUEST" & i, "STARTMSG"))
        Quests(i).goldGiven = Val(GetVar(filename, "QUEST" & i, "GOLDGIVEN"))
        DoEvents
        Quests(i).StartQuestMsg = Replace(Quests(i).StartQuestMsg, "/n", vbCrLf)
        Quests(i).GetItemQuestMsg = Replace(Quests(i).GetItemQuestMsg, "/n", vbCrLf)
        Quests(i).FinishQuestMessage = Replace(Quests(i).FinishQuestMessage, "/n", vbCrLf)
    Next i
End Sub

Sub LoadPrayers()
Dim filename As String
Dim i As Long

    Call CheckPrayers
    
    filename = App.Path & "\prayer.ini"
    
    For i = 1 To MAX_SPELLS
        Prayer(i).Name = GetVar(filename, "PRAYER" & i, "Name")
        Prayer(i).ClassReq = Val(GetVar(filename, "PRAYER" & i, "ClassReq"))
        Prayer(i).type = Val(GetVar(filename, "PRAYER" & i, "Type"))
        Prayer(i).Data1 = Val(GetVar(filename, "PRAYER" & i, "Data1"))
        Prayer(i).Data2 = Val(GetVar(filename, "PRAYER" & i, "Data2"))
        Prayer(i).Data3 = Val(GetVar(filename, "PRAYER" & i, "Data3"))
        Prayer(i).sound = Val(GetVar(filename, "PRAYER" & i, "Sound"))
        Prayer(i).manaUse = Val(GetVar(filename, "PRAYER" & i, "ManaUse"))
        DoEvents
    Next i
End Sub

Sub LoadSigns()
Dim filename As String
Dim i As Long

    Call CheckSigns
    
    filename = App.Path & "\signs.ini"
    
    For i = 1 To MAX_SIGNS
        Signs(i).header = GetVar(filename, "SIGN" & i, "Header")
        Signs(i).msg = GetVar(filename, "SIGN" & i, "Msg")
        Signs(i).msg = Replace(Trim(Signs(i).msg), "/n", vbCrLf)
        DoEvents
    Next i
End Sub


Sub CheckSpells()
    If Not FileExist("spells.ini") Then
        Call SaveSpells
    End If
End Sub

Sub CheckQuests()
    If Not FileExist("quests.ini") Then
        Call SaveQuests
    End If
End Sub

Sub CheckSigns()
    If Not FileExist("signs.ini") Then
        Call SaveSigns
    End If
End Sub

Sub SaveNpcs()
Dim i As Long
    
    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next i
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
Dim filename As String

    filename = App.Path & "\npcs.ini"
    
    Call PutVar(filename, "NPC" & NpcNum, "Name", Trim(Npc(NpcNum).Name))
    Call PutVar(filename, "NPC" & NpcNum, "AttackSay", Trim(Npc(NpcNum).AttackSay))
    Call PutVar(filename, "NPC" & NpcNum, "Sprite", Trim(Npc(NpcNum).sprite))
    Call PutVar(filename, "NPC" & NpcNum, "SpawnSecs", Trim(Npc(NpcNum).SpawnSecs))
    Call PutVar(filename, "NPC" & NpcNum, "Behavior", Trim(Npc(NpcNum).Behavior))
    Call PutVar(filename, "NPC" & NpcNum, "Range", Trim(Npc(NpcNum).Range))
    Call PutVar(filename, "NPC" & NpcNum, "DropChance", Trim(Npc(NpcNum).DropChance))
    Call PutVar(filename, "NPC" & NpcNum, "DropItem", Trim(Npc(NpcNum).DropItem))
    Call PutVar(filename, "NPC" & NpcNum, "DropItemValue", Trim(Npc(NpcNum).DropItemValue))
    Call PutVar(filename, "NPC" & NpcNum, "STR", Trim(Npc(NpcNum).str))
    Call PutVar(filename, "NPC" & NpcNum, "DEF", Trim(Npc(NpcNum).def))
    Call PutVar(filename, "NPC" & NpcNum, "SPEED", Trim(Npc(NpcNum).speed))
    Call PutVar(filename, "NPC" & NpcNum, "MAGI", Trim(Npc(NpcNum).MAGI))
    Call PutVar(filename, "NPC" & NpcNum, "Gold", Trim(Npc(NpcNum).Gold))
    Call PutVar(filename, "NPC" & NpcNum, "ExpGiven", Trim(Npc(NpcNum).ExpGiven))
    Call PutVar(filename, "NPC" & NpcNum, "HP", Trim(Npc(NpcNum).HP))
    Call PutVar(filename, "NPC" & NpcNum, "Respawn", Trim(Npc(NpcNum).Respawn))
    Call PutVar(filename, "NPC" & NpcNum, "Attack_with_Poison", Trim(Npc(NpcNum).Attack_with_Poison))
    Call PutVar(filename, "NPC" & NpcNum, "QuestID", Trim(Npc(NpcNum).QuestID))
    Call PutVar(filename, "NPC" & NpcNum, "OpensShop", Trim(Npc(NpcNum).opensShop))
    Call PutVar(filename, "NPC" & NpcNum, "OpensBank", Trim(Npc(NpcNum).opensBank))
    Call PutVar(filename, "NPC" & NpcNum, "Type", Trim(Npc(NpcNum).type))
End Sub

Sub LoadNpcs()
On Error Resume Next

Dim filename As String
Dim i As Long

    Call CheckNpcs
    
    filename = App.Path & "\npcs.ini"
    
    For i = 1 To MAX_NPCS
        Npc(i).Name = GetVar(filename, "NPC" & i, "Name")
        Npc(i).AttackSay = GetVar(filename, "NPC" & i, "AttackSay")
        Npc(i).sprite = GetVar(filename, "NPC" & i, "Sprite")
        Npc(i).SpawnSecs = GetVar(filename, "NPC" & i, "SpawnSecs")
        Npc(i).Behavior = GetVar(filename, "NPC" & i, "Behavior")
        Npc(i).Range = GetVar(filename, "NPC" & i, "Range")
        Npc(i).DropChance = GetVar(filename, "NPC" & i, "DropChance")
        Npc(i).DropItem = GetVar(filename, "NPC" & i, "DropItem")
        Npc(i).DropItemValue = GetVar(filename, "NPC" & i, "DropItemValue")
        Npc(i).str = GetVar(filename, "NPC" & i, "STR")
        Npc(i).def = GetVar(filename, "NPC" & i, "DEF")
        Npc(i).speed = GetVar(filename, "NPC" & i, "SPEED")
        Npc(i).MAGI = GetVar(filename, "NPC" & i, "MAGI")
        Npc(i).Gold = GetVar(filename, "NPC" & i, "Gold")
        Npc(i).ExpGiven = GetVar(filename, "NPC" & i, "ExpGiven")
        Npc(i).HP = GetVar(filename, "NPC" & i, "HP")
        Npc(i).Respawn = GetVar(filename, "NPC" & i, "Respawn")
        Npc(i).Attack_with_Poison = GetVar(filename, "NPC" & i, "Attack_with_Poison")
        Npc(i).QuestID = GetVar(filename, "NPC" & i, "QuestID")
        Npc(i).opensShop = GetVar(filename, "NPC" & i, "OpensShop")
        Npc(i).opensBank = GetVar(filename, "NPC" & i, "OpensBank")
        Npc(i).type = GetVar(filename, "NPC" & i, "Type")
    
        DoEvents
    Next i
End Sub

Sub CheckNpcs()
    If Not FileExist("npcs.ini") Then
        Call SaveNpcs
    End If
End Sub

Sub SaveMap(ByVal mapNum As Long)
Dim filename As String
Dim f As Long

    filename = App.Path & "\maps\map" & mapNum & ".dat"
        
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , map(mapNum)
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
        filename = App.Path & "\maps\map" & i & ".dat"
        
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , map(i)
        Close #f
    
        DoEvents
    Next i
End Sub

Sub ConvertOldMapsToNew()
Dim filename As String
Dim i As Long
Dim f As Long
Dim x As Long, y As Long
Dim OldMap As OldMapRec
Dim NewMap As MapRec

    For i = 1 To MAX_MAPS
        filename = App.Path & "\maps\map" & i & ".dat"
        
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
        NewMap.Night = OldMap.Night
        NewMap.Respawn = OldMap.Respawn
        NewMap.Bank = OldMap.Bank
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                NewMap.Tile(x, y).Ground = OldMap.Tile(x, y).Ground
                NewMap.Tile(x, y).Mask = OldMap.Tile(x, y).Mask
                NewMap.Tile(x, y).Anim = OldMap.Tile(x, y).Anim
                NewMap.Tile(x, y).Fringe = OldMap.Tile(x, y).Fringe
                NewMap.Tile(x, y).type = OldMap.Tile(x, y).type
                NewMap.Tile(x, y).Data1 = OldMap.Tile(x, y).Data1
                NewMap.Tile(x, y).Data2 = OldMap.Tile(x, y).Data2
                NewMap.Tile(x, y).Data3 = OldMap.Tile(x, y).Data3
                NewMap.Tile(x, y).Data4 = OldMap.Tile(x, y).Data4
                NewMap.Tile(x, y).Data5 = OldMap.Tile(x, y).Data5
                NewMap.Tile(x, y).TileSheet_Ground = OldMap.Tile(x, y).TileSheet_Ground
                NewMap.Tile(x, y).TileSheet_Fringe = OldMap.Tile(x, y).TileSheet_Fringe
                NewMap.Tile(x, y).TileSheet_Anim = OldMap.Tile(x, y).TileSheet_Anim
                NewMap.Tile(x, y).TileSheet_Mask = OldMap.Tile(x, y).TileSheet_Mask
            Next x
        Next y
        
        For x = 1 To MAX_MAP_NPCS
            NewMap.Npc(x) = OldMap.Npc(x)
        Next x
        
        ' Set new values to 0 or null
        NewMap.street = 0
        
        ' Save the new map
        f = FreeFile
        Open filename For Binary As #f
            Put #f, , NewMap
        Close #f
    Next i
End Sub

Sub CheckMaps()
Dim filename As String
Dim x As Long
Dim y As Long
Dim i As Long
Dim n As Long

    Call ClearMaps
        
    For i = 1 To MAX_MAPS
        filename = "maps\map" & i & ".dat"
        
        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(filename) Then
            Call SaveMap(i)
        End If
    Next i
End Sub

Sub AddLog(ByVal Text As String, ByVal FN As String, Optional ByVal index As Long = 0)
'If FN = "raw" Then Exit Sub
Dim file As String
Dim filename As String
Dim f As Long
file = "\logs\" & FN & getFileNameForLogs()
    'If ServerLog = True Then
        filename = App.Path & file
        'Debug.Print getFileNameForLogs()
        'Debug.Print FileExist(file)
        If Not FileExist(file) Then
            f = FreeFile
            Open filename For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open filename For Append As #f
            If index > 0 Then
                Print #f, frmServer.Socket(index).RemoteHostIP & " - " & Time & ": " & Text & vbCrLf
            Else
                Print #f, Time & ": " & Text & vbCrLf
            End If
        Close #f
    'End If
End Sub

Private Function getFileNameForLogs() As String
'function returns the same value for 7 days then a new value for 7 days
'only approx but make new log file every 7 days
    Dim dateStr As String
    Dim lngDay As Long
    Dim lngMonth As Long
    dateStr = Now
    lngDay = Left(dateStr, 2)
    lngMonth = Mid(dateStr, 4, 2)
    'MsgBox lngDay
    While (lngDay / 7) <> Int(lngDay / 7)
        lngDay = lngDay - 1
    Wend
    getFileNameForLogs = "." & lngDay & "_" & lngMonth & ".log"
    DoEvents
End Function


Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
Dim filename, IP As String
Dim f As Long, i As Long

    filename = App.Path & "\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("banlist.txt") Then
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
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", RGB_GlobalColor)
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


Sub LoadInitData()
Dim filename As String
filename = App.Path & "\info.ini"
GAME_NAME = GetVar(filename, "INFO", "Name")
GAME_PORT = Val(GetVar(filename, "INFO", "Port"))
GAME_UPDATE_PORT = 0

If GAME_PORT = 0 Then GAME_PORT = 10001
End Sub
