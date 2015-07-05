Attribute VB_Name = "modDatabase"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public START_MAP As Long
Public START_X As Long
Public START_Y As Long

Public temp As Integer

Public Const ADMIN_LOG = "logs\admin.txt"
Public Const PLAYER_LOG = "logs\player.txt"

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

Sub LoadExps()
Dim FileName As String
Dim I As Long

    Call CheckExps
    
    FileName = App.Path & "\experience.ini"
    
    For I = 1 To MAX_LEVEL
        temp = I / MAX_LEVEL * 100
        Call SetStatus("Loading exp... " & temp & "%")
        Experience(I) = GetVar(FileName, "EXPERIENCE", "Exp" & I)
        
        DoEvents
    Next I
End Sub

Sub CheckExps()
    If Not FileExist("experience.ini") Then
        Dim I As Long
    
        For I = 1 To MAX_LEVEL
            temp = I / MAX_LEVEL * 100
            Call SetStatus("Saving exp... " & temp & "%")
            DoEvents
            Call PutVar(App.Path & "\experience.ini", "EXPERIENCE", "Exp" & I, I * 1500)
        Next I
    End If
End Sub

Sub ClearExps()
Dim I As Long

    For I = 1 To MAX_LEVEL
        Experience(I) = 0
    Next I
End Sub

Sub LoadEmos()
Dim FileName As String
Dim I As Long

    Call CheckEmos
    
    FileName = App.Path & "\emoticons.ini"
    
    For I = 0 To MAX_EMOTICONS
        temp = I / MAX_EMOTICONS * 100
        Call SetStatus("Loading emoticons... " & temp & "%")
        Emoticons(I).Pic = GetVar(FileName, "EMOTICONS", "Emoticon" & I)
        Emoticons(I).Command = GetVar(FileName, "EMOTICONS", "EmoticonC" & I)
        
        DoEvents
    Next I
End Sub

Sub CheckEmos()
    If Not FileExist("emoticons.ini") Then
        Dim I As Long
    
        For I = 0 To MAX_EMOTICONS
            temp = I / MAX_LEVEL * 100
            Call SetStatus("Saving emoticons... " & temp & "%")
            DoEvents
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "Emoticon" & I, 0)
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "EmoticonC" & I, "")
        Next I
    End If
End Sub

Sub ClearEmos()
Dim I As Long

    For I = 0 To MAX_EMOTICONS
        Emoticons(I).Pic = 0
        Emoticons(I).Command = ""
    Next I
End Sub

Sub SaveEmoticon(ByVal EmoNum As Long)
Dim FileName As String

    FileName = App.Path & "\emoticons.ini"
    
    Call PutVar(FileName, "EMOTICONS", "EmoticonC" & EmoNum, Trim(Emoticons(EmoNum).Command))
    Call PutVar(FileName, "EMOTICONS", "Emoticon" & EmoNum, Val(Emoticons(EmoNum).Pic))
End Sub

Function FileExist(ByVal FileName As String) As Boolean
    If Dir(App.Path & "\" & FileName) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Sub SavePlayer(ByVal Index As Long)
Dim FileName As String
Dim I As Long
Dim n As Long

    FileName = App.Path & "\accounts\" & Trim(Player(Index).Login) & ".ini"
    
    Call PutVar(FileName, "GENERAL", "Login", Trim(Player(Index).Login))
    Call PutVar(FileName, "GENERAL", "Password", Trim(Player(Index).Password))
    Call PutVar(FileName, "GENERAL", "Email", Trim(Player(Index).Email))

    For I = 1 To MAX_CHARS
        ' General
        Call PutVar(FileName, "CHAR" & I, "Name", Trim(Player(Index).Char(I).Name))
        Call PutVar(FileName, "CHAR" & I, "Class", STR(Player(Index).Char(I).Class))
        Call PutVar(FileName, "CHAR" & I, "Sex", STR(Player(Index).Char(I).Sex))
        Call PutVar(FileName, "CHAR" & I, "Sprite", STR(Player(Index).Char(I).Sprite))
        Call PutVar(FileName, "CHAR" & I, "Level", STR(Player(Index).Char(I).Level))
        Call PutVar(FileName, "CHAR" & I, "Exp", STR(Player(Index).Char(I).Exp))
        Call PutVar(FileName, "CHAR" & I, "Access", STR(Player(Index).Char(I).Access))
        Call PutVar(FileName, "CHAR" & I, "PK", STR(Player(Index).Char(I).PK))
        Call PutVar(FileName, "CHAR" & I, "Guild", Trim(Player(Index).Char(I).Guild))
        Call PutVar(FileName, "CHAR" & I, "Guildaccess", STR(Player(Index).Char(I).Guildaccess))

        
        ' Vitals
        Call PutVar(FileName, "CHAR" & I, "HP", STR(Player(Index).Char(I).HP))
        Call PutVar(FileName, "CHAR" & I, "MP", STR(Player(Index).Char(I).MP))
        Call PutVar(FileName, "CHAR" & I, "SP", STR(Player(Index).Char(I).SP))
        
        ' Stats
        Call PutVar(FileName, "CHAR" & I, "STR", STR(Player(Index).Char(I).STR))
        Call PutVar(FileName, "CHAR" & I, "DEF", STR(Player(Index).Char(I).DEF))
        Call PutVar(FileName, "CHAR" & I, "LUCK", STR(Player(Index).Char(I).Luck))
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
            Player(Index).Char(I).X = START_X
            Player(Index).Char(I).Y = START_Y
        End If
            
        ' Position
        Call PutVar(FileName, "CHAR" & I, "Map", STR(Player(Index).Char(I).Map))
        Call PutVar(FileName, "CHAR" & I, "X", STR(Player(Index).Char(I).X))
        Call PutVar(FileName, "CHAR" & I, "Y", STR(Player(Index).Char(I).Y))
        Call PutVar(FileName, "CHAR" & I, "Dir", STR(Player(Index).Char(I).Dir))
        
        ' Inventory
        For n = 1 To MAX_INV
            Call PutVar(FileName, "CHAR" & I, "InvItemNum" & n, STR(Player(Index).Char(I).Inv(n).num))
            Call PutVar(FileName, "CHAR" & I, "InvItemVal" & n, STR(Player(Index).Char(I).Inv(n).Value))
            Call PutVar(FileName, "CHAR" & I, "InvItemDur" & n, STR(Player(Index).Char(I).Inv(n).Dur))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Call PutVar(FileName, "CHAR" & I, "Spell" & n, STR(Player(Index).Char(I).Spell(n)))
        Next n
    Next I
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
Dim FileName As String
Dim I As Long
Dim n As Long

    Call ClearPlayer(Index)
    
    FileName = App.Path & "\accounts\" & Trim(Name) & ".ini"

    Player(Index).Login = GetVar(FileName, "GENERAL", "Login")
    Player(Index).Password = GetVar(FileName, "GENERAL", "Password")
    Player(Index).Email = GetVar(FileName, "GENERAL", "Email")

    For I = 1 To MAX_CHARS
        ' General
        Player(Index).Char(I).Name = GetVar(FileName, "CHAR" & I, "Name")
        Player(Index).Char(I).Sex = Val(GetVar(FileName, "CHAR" & I, "Sex"))
        Player(Index).Char(I).Class = Val(GetVar(FileName, "CHAR" & I, "Class"))
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
        Player(Index).Char(I).STR = Val(GetVar(FileName, "CHAR" & I, "STR"))
        Player(Index).Char(I).DEF = Val(GetVar(FileName, "CHAR" & I, "DEF"))
        Player(Index).Char(I).Luck = Val(GetVar(FileName, "CHAR" & I, "LUCK"))
        Player(Index).Char(I).Magi = Val(GetVar(FileName, "CHAR" & I, "MAGI"))
        Player(Index).Char(I).POINTS = Val(GetVar(FileName, "CHAR" & I, "POINTS"))
        
        ' Worn equipment
        Player(Index).Char(I).ArmorSlot = Val(GetVar(FileName, "CHAR" & I, "ArmorSlot"))
        Player(Index).Char(I).WeaponSlot = Val(GetVar(FileName, "CHAR" & I, "WeaponSlot"))
        Player(Index).Char(I).HelmetSlot = Val(GetVar(FileName, "CHAR" & I, "HelmetSlot"))
        Player(Index).Char(I).ShieldSlot = Val(GetVar(FileName, "CHAR" & I, "ShieldSlot"))
        
        ' Position
        Player(Index).Char(I).Map = Val(GetVar(FileName, "CHAR" & I, "Map"))
        Player(Index).Char(I).X = Val(GetVar(FileName, "CHAR" & I, "X"))
        Player(Index).Char(I).Y = Val(GetVar(FileName, "CHAR" & I, "Y"))
        Player(Index).Char(I).Dir = Val(GetVar(FileName, "CHAR" & I, "Dir"))
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(I).Map = 0 Then
            Player(Index).Char(I).Map = START_MAP
            Player(Index).Char(I).X = START_X
            Player(Index).Char(I).Y = START_Y
        End If
        
        ' Inventory
        For n = 1 To MAX_INV
            Player(Index).Char(I).Inv(n).num = Val(GetVar(FileName, "CHAR" & I, "InvItemNum" & n))
            Player(Index).Char(I).Inv(n).Value = Val(GetVar(FileName, "CHAR" & I, "InvItemVal" & n))
            Player(Index).Char(I).Inv(n).Dur = Val(GetVar(FileName, "CHAR" & I, "InvItemDur" & n))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(I).Spell(n) = Val(GetVar(FileName, "CHAR" & I, "Spell" & n))
        Next n
    Next I
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

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If Trim(Player(Index).Char(CharNum).Name) <> "" Then
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

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String, ByVal Email)
Dim I As Long

    Player(Index).Login = Name
    Player(Index).Password = Password
    Player(Index).Email = Email
    
    For I = 1 To MAX_CHARS
        Call ClearChar(Index, I)
    Next I
    
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
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).MaleSprite
        Else
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).FemaleSprite
        End If
        
        Player(Index).Char(CharNum).Level = 1
                    
        Player(Index).Char(CharNum).STR = Class(ClassNum).STR
        Player(Index).Char(CharNum).DEF = Class(ClassNum).DEF
        Player(Index).Char(CharNum).Luck = Class(ClassNum).Luck
        Player(Index).Char(CharNum).Magi = Class(ClassNum).Magi
        
        If Class(ClassNum).Map <= 0 Then Class(ClassNum).Map = 1
        If Class(ClassNum).X < 0 Or Class(ClassNum).X > MAX_MAPX Then Class(ClassNum).X = Int(Class(ClassNum).X / 2)
        If Class(ClassNum).Y < 0 Or Class(ClassNum).Y > MAX_MAPY Then Class(ClassNum).Y = Int(Class(ClassNum).Y / 2)
        Player(Index).Char(CharNum).Map = Class(ClassNum).Map
        Player(Index).Char(CharNum).X = Class(ClassNum).X
        Player(Index).Char(CharNum).Y = Class(ClassNum).Y
            
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
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SavePlayer(I)
        End If
    Next I
End Sub

Sub LoadClasses()
Dim FileName As String
Dim I As Long

    Call CheckClasses
    
    FileName = App.Path & "\Classes\info.ini"
    
    MAX_CLASSES = Val(GetVar(FileName, "INFO", "MaxClasses"))
    
    ReDim Class(0 To MAX_CLASSES) As ClassRec
    
    Call ClearClasses
    
    For I = 0 To MAX_CLASSES
        temp = I / MAX_CLASSES * 100
        Call SetStatus("Loading classes... " & temp & "%")
        FileName = App.Path & "\Classes\Class" & I & ".ini"
        Class(I).Name = GetVar(FileName, "CLASS", "Name")
        Class(I).MaleSprite = GetVar(FileName, "CLASS", "MaleSprite")
        Class(I).FemaleSprite = GetVar(FileName, "CLASS", "FemaleSprite")
        Class(I).STR = Val(GetVar(FileName, "CLASS", "STR"))
        Class(I).DEF = Val(GetVar(FileName, "CLASS", "DEF"))
        Class(I).Luck = Val(GetVar(FileName, "CLASS", "LUCK"))
        Class(I).Magi = Val(GetVar(FileName, "CLASS", "MAGI"))
        Class(I).Map = Val(GetVar(FileName, "CLASS", "MAP"))
        Class(I).X = Val(GetVar(FileName, "CLASS", "X"))
        Class(I).Y = Val(GetVar(FileName, "CLASS", "Y"))
        Class(I).Locked = Val(GetVar(FileName, "CLASS", "Locked"))
        
        DoEvents
    Next I
End Sub

Sub SaveClasses()
Dim FileName As String
Dim I As Long

    FileName = App.Path & "\Classes\info.ini"
    
    If Not FileExist("Classes\info.ini") Then
        Call PutVar(FileName, "INFO", "MaxClasses", 3)
        Call PutVar(FileName, "INFO", "MaxSkills", 25)
        Call PutVar(FileName, "INFO", "StatPoints", 0)
        Call PutVar(FileName, "INFO", "SkillPoints", 0)
    End If
    
    For I = 0 To MAX_CLASSES
        temp = I / MAX_CLASSES * 100
        Call SetStatus("Saving classes... " & temp & "%")
        DoEvents
        FileName = App.Path & "\Classes\Class" & I & ".ini"
        If Not FileExist("Classes\Class" & I & ".ini") Then
            Call PutVar(FileName, "CLASS", "Name", Trim(Class(I).Name))
            Call PutVar(FileName, "CLASS", "MaleSprite", STR(Class(I).MaleSprite))
            Call PutVar(FileName, "CLASS", "FemaleSprite", STR(Class(I).FemaleSprite))
            Call PutVar(FileName, "CLASS", "STR", STR(Class(I).STR))
            Call PutVar(FileName, "CLASS", "DEF", STR(Class(I).DEF))
            Call PutVar(FileName, "CLASS", "LUCK", STR(Class(I).Luck))
            Call PutVar(FileName, "CLASS", "MAGI", STR(Class(I).Magi))
            Call PutVar(FileName, "CLASS", "MAP", STR(Class(I).Map))
            Call PutVar(FileName, "CLASS", "X", STR(Class(I).X))
            Call PutVar(FileName, "CLASS", "Y", STR(Class(I).Y))
            Call PutVar(FileName, "CLASS", "Locked", STR(Class(I).Locked))
        End If
    Next I
End Sub

Sub CheckClasses()
    If Not FileExist("Classes\info.ini") Then
        Call SaveClasses
    End If
End Sub

Sub LoadClasses2()
Dim FileName As String
Dim I As Long

    Call CheckClasses2
    
    FileName = App.Path & "\FirstClassAdvancement.ini"
    
    MAX_CLASSES = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class2(0 To MAX_CLASSES) As ClassRec
    
    Call ClearClasses2
    
    For I = 0 To MAX_CLASSES
        temp = I / MAX_CLASSES * 100
        Call SetStatus("Loading first class advandcement... " & temp & "%")
        Class2(I).Name = GetVar(FileName, "CLASS" & I, "Name")
        Class2(I).AdvanceFrom = Val(GetVar(FileName, "CLASS" & I, "AdvanceFrom"))
        Class2(I).MaleSprite = GetVar(FileName, "CLASS" & I, "MaleSprite")
        Class2(I).FemaleSprite = GetVar(FileName, "CLASS" & I, "FemaleSprite")
        Class2(I).STR = Val(GetVar(FileName, "CLASS" & I, "STR"))
        Class2(I).DEF = Val(GetVar(FileName, "CLASS" & I, "DEF"))
        Class2(I).Luck = Val(GetVar(FileName, "CLASS" & I, "LUCK"))
        Class2(I).Magi = Val(GetVar(FileName, "CLASS" & I, "MAGI"))
        
        DoEvents
    Next I
End Sub

Sub SaveClasses2()
Dim FileName As String
Dim I As Long

    FileName = App.Path & "\FirstClassAdvancement.ini"
    
    For I = 0 To MAX_CLASSES
        temp = I / MAX_CLASSES * 100
        Call SetStatus("Saving first class advandcement... " & temp & "%")
        DoEvents
        Call PutVar(FileName, "CLASS" & I, "Name", Trim(Class2(I).Name))
        Call PutVar(FileName, "CLASS" & I, "AdvanceFrom", STR(Class2(I).AdvanceFrom))
        Call PutVar(FileName, "CLASS" & I, "LevelReq", STR(Class2(I).LevelReq))
        Call PutVar(FileName, "CLASS" & I, "MaleSprite", STR(Class2(I).MaleSprite))
        Call PutVar(FileName, "CLASS" & I, "FemaleSprite", STR(Class2(I).FemaleSprite))
        Call PutVar(FileName, "CLASS" & I, "STR", STR(Class2(I).STR))
        Call PutVar(FileName, "CLASS" & I, "DEF", STR(Class2(I).DEF))
        Call PutVar(FileName, "CLASS" & I, "LUCK", STR(Class2(I).Luck))
        Call PutVar(FileName, "CLASS" & I, "MAGI", STR(Class2(I).Magi))
    Next I
End Sub

Sub CheckClasses2()
    If Not FileExist("FirstClassAdvancement.ini") Then
        Call SaveClasses2
    End If
End Sub

Sub Loadclasses3()
Dim FileName As String
Dim I As Long

    Call Checkclasses3
    
    FileName = App.Path & "\SecondClassAdvancement.ini"
    
    MAX_CLASSES = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class3(0 To MAX_CLASSES) As ClassRec
    
    Call ClearClasses3
    
    For I = 0 To MAX_CLASSES
        temp = I / MAX_CLASSES * 100
        Call SetStatus("Loading second class advandcement... " & temp & "%")
        Class3(I).Name = GetVar(FileName, "CLASS" & I, "Name")
        Class3(I).AdvanceFrom = Val(GetVar(FileName, "CLASS" & I, "AdvanceFrom"))
        Class3(I).LevelReq = Val(GetVar(FileName, "CLASS" & I, "LevelReq"))
        Class3(I).MaleSprite = GetVar(FileName, "CLASS" & I, "MaleSprite")
        Class3(I).FemaleSprite = GetVar(FileName, "CLASS" & I, "FemaleSprite")
        Class3(I).STR = Val(GetVar(FileName, "CLASS" & I, "STR"))
        Class3(I).DEF = Val(GetVar(FileName, "CLASS" & I, "DEF"))
        Class3(I).Luck = Val(GetVar(FileName, "CLASS" & I, "LUCK"))
        Class3(I).Magi = Val(GetVar(FileName, "CLASS" & I, "MAGI"))
        
        DoEvents
    Next I
End Sub

Sub Saveclasses3()
Dim FileName As String
Dim I As Long

    FileName = App.Path & "\SecondClassAdvancement.ini"
    
    For I = 0 To MAX_CLASSES
        temp = I / MAX_CLASSES * 100
        Call SetStatus("Saving second class advandcement... " & temp & "%")
        DoEvents
        Call PutVar(FileName, "CLASS" & I, "Name", Trim(Class3(I).Name))
        Call PutVar(FileName, "CLASS" & I, "AdvanceFrom", STR(Class3(I).AdvanceFrom))
        Call PutVar(FileName, "CLASS" & I, "MaleSprite", STR(Class3(I).MaleSprite))
        Call PutVar(FileName, "CLASS" & I, "FemaleSprite", STR(Class3(I).FemaleSprite))
        Call PutVar(FileName, "CLASS" & I, "STR", STR(Class3(I).STR))
        Call PutVar(FileName, "CLASS" & I, "DEF", STR(Class3(I).DEF))
        Call PutVar(FileName, "CLASS" & I, "LUCK", STR(Class3(I).Luck))
        Call PutVar(FileName, "CLASS" & I, "MAGI", STR(Class3(I).Magi))
    Next I
End Sub

Sub Checkclasses3()
    If Not FileExist("SecondClassAdvancement.ini") Then
        Call Saveclasses3
    End If
End Sub

Sub SaveItems()
Dim I As Long
        
    Call SetStatus("Saving items... ")
    For I = 1 To MAX_ITEMS
        If Not FileExist("items\item" & I & ".dat") Then
            temp = I / MAX_ITEMS * 100
            Call SetStatus("Saving items... " & temp & "%")
            DoEvents
            Call SaveItem(I)
        End If
    Next I
End Sub

Sub SaveItem(ByVal ItemNum As Long)
Dim FileName As String
Dim f  As Long
FileName = App.Path & "\items\item" & ItemNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Item(ItemNum)
    Close #f
End Sub

Sub LoadItems()
Dim FileName As String
Dim I As Long
Dim f As Long

    Call CheckItems
    
    For I = 1 To MAX_ITEMS
        temp = I / MAX_ITEMS * 100
        Call SetStatus("Loading items... " & temp & "%")
        
        FileName = App.Path & "\Items\Item" & I & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Item(I)
        Close #f
        
        DoEvents
    Next I
End Sub

Sub CheckItems()
    Call SaveItems
End Sub

Sub SaveShops()
Dim I As Long

    Call SetStatus("Saving shops... ")
    For I = 1 To MAX_SHOPS
        If Not FileExist("shops\shop" & I & ".dat") Then
            temp = I / MAX_SHOPS * 100
            Call SetStatus("Saving shops... " & temp & "%")
            DoEvents
            Call SaveShop(I)
        End If
    Next I
End Sub

Sub SaveShop(ByVal ShopNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\shops\shop" & ShopNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Shop(ShopNum)
    Close #f
End Sub

Sub LoadShops()
Dim FileName As String
Dim I As Long, f As Long

    Call CheckShops
    
    For I = 1 To MAX_SHOPS
        temp = I / MAX_SHOPS * 100
        Call SetStatus("Loading shops... " & temp & "%")
        FileName = App.Path & "\shops\shop" & I & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Shop(I)
        Close #f
        
        DoEvents
    Next I
End Sub

Sub CheckShops()
    Call SaveShops
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\spells\spells" & SpellNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Spell(SpellNum)
    Close #f
End Sub

Sub SaveSpells()
Dim I As Long

    Call SetStatus("Saving spells... ")
    For I = 1 To MAX_SPELLS
        If Not FileExist("spells\spells" & I & ".dat") Then
            temp = I / MAX_SPELLS * 100
            Call SetStatus("Saving spells... " & temp & "%")
            DoEvents
            Call SaveSpell(I)
        End If
    Next I
End Sub

Sub LoadSpells()
Dim FileName As String
Dim I As Long
Dim f As Long

    Call CheckSpells
    
    For I = 1 To MAX_SPELLS
        temp = I / MAX_SPELLS * 100
        Call SetStatus("Loading spells... " & temp & "%")
        
        FileName = App.Path & "\spells\spells" & I & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Spell(I)
        Close #f
        
        DoEvents
    Next I
End Sub

Sub CheckSpells()
    Call SaveSpells
End Sub

Sub SaveNpcs()
Dim I As Long

    Call SetStatus("Saving npcs... ")
    
    For I = 1 To MAX_NPCS
        If Not FileExist("npcs\npc" & I & ".dat") Then
            temp = I / MAX_NPCS * 100
            Call SetStatus("Saving npcs... " & temp & "%")
            DoEvents
            Call SaveNpc(I)
        End If
    Next I
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
Dim FileName As String
Dim f As Long
FileName = App.Path & "\npcs\npc" & NpcNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Npc(NpcNum)
    Close #f
End Sub

Sub LoadNpcs()
Dim FileName As String
Dim I As Long
Dim z As Long
Dim f As Long

    Call CheckNpcs
        
    For I = 1 To MAX_NPCS
        temp = I / MAX_NPCS * 100
        Call SetStatus("Loading npcs... " & temp & "%")
        FileName = App.Path & "\npcs\npc" & I & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Npc(I)
        Close #f
        
        DoEvents
    Next I
End Sub

Sub CheckNpcs()
    Call SaveNpcs
End Sub

Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub LoadMaps()
Dim FileName As String
Dim I As Long
Dim f As Long

    Call CheckMaps
    
    For I = 1 To MAX_MAPS
        temp = I / MAX_MAPS * 100
        Call SetStatus("Loading maps... " & temp & "%")
        FileName = App.Path & "\maps\map" & I & ".dat"
        
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Map(I)
        Close #f
    
        DoEvents
    Next I
End Sub

Sub CheckMaps()
Dim FileName As String
Dim X As Long
Dim Y As Long
Dim I As Long
Dim n As Long

    Call ClearMaps
        
    For I = 1 To MAX_MAPS
        FileName = "maps\map" & I & ".dat"
        
        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(FileName) Then
            temp = I / MAX_MAPS * 100
            Call SetStatus("Saving maps... " & temp & "%")
            DoEvents
            
            Call SaveMap(I)
        End If
    Next I
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
Dim f As Long, I As Long

    FileName = App.Path & "\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)
            
    For I = Len(IP) To 1 Step -1
        If Mid(IP, I, 1) = "." Then
            Exit For
        End If
    Next I
    IP = Mid(IP, 1, I)
            
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

Sub BanByServer(ByVal BanPlayerIndex As Long, ByVal Reason As String)
Dim FileName, IP As String
Dim f As Long, I As Long

    FileName = App.Path & "\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)
    
    For I = Len(IP) To 1 Step -1
        If Mid(IP, I, 1) = "." Then
            Exit For
        End If
    Next I
    IP = Mid(IP, 1, I)
            
    f = FreeFile
    Open FileName For Append As #f
        Print #f, IP & "," & "Server"
    Close #f
    
    If Trim(Reason) <> "" Then
        Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by the server! Reason(" & Reason & ")", White)
        Call AlertMsg(BanPlayerIndex, "You have been banned by the server!  Reason(" & Reason & ")")
        Call AddLog("The server has banned " & GetPlayerName(BanPlayerIndex) & ".  Reason(" & Reason & ")", ADMIN_LOG)
    Else
        Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by the server!", White)
        Call AlertMsg(BanPlayerIndex, "You have been banned by the server!")
        Call AddLog("The server has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    End If
End Sub

Private Function Replace(strWord, strFind, strReplace, charAmount) As String
Dim a  As Integer

    a = InStr(1, UCase(strWord), UCase(strFind))
    On Error Resume Next
    strWord = Mid(strWord, 1, a - 1) & strReplace & Right(strWord, Len(strWord) - a - charAmount + 1)
    Replace = strWord
End Function

Sub SaveLogs()
Dim FileName As String
Dim I As String, c As String

    If LCase(Dir(App.Path & "\logs", vbDirectory)) <> "logs" Then
        Call MkDir(App.Path & "\Logs")
    End If
    
    c = Time
    c = Replace(c, ":", ".", 1)
    c = Replace(c, ":", ".", 1)
    c = Replace(c, ":", ".", 1)
    
    I = Date
    I = Replace(I, "/", ".", 1)
    I = Replace(I, "/", ".", 1)
    I = Replace(I, "/", ".", 1)
    
    If LCase(Dir(App.Path & "\logs\" & I, vbDirectory)) <> I Then
        Call MkDir(App.Path & "\Logs\" & I & "\")
    End If
    
    If LCase(Dir(App.Path & "\logs\" & I & "\" & c, vbDirectory)) <> c Then
        Call MkDir(App.Path & "\Logs\" & I & "\" & c & "\")
    End If
        
    FileName = App.Path & "\Logs\" & I & "\" & c & "\Main.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(0).Text
    Close #1
    
    FileName = App.Path & "\Logs\" & I & "\" & c & "\Broadcast.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(1).Text
    Close #1
    
    FileName = App.Path & "\Logs\" & I & "\" & c & "\Global.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(2).Text
    Close #1
    
    FileName = App.Path & "\Logs\" & I & "\" & c & "\Map.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(3).Text
    Close #1
    
    FileName = App.Path & "\Logs\" & I & "\" & c & "\Private.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(4).Text
    Close #1
    
    FileName = App.Path & "\Logs\" & I & "\" & c & "\Admin.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(5).Text
    Close #1
    
    FileName = App.Path & "\Logs\" & I & "\" & c & "\Emote.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(6).Text
    Close #1
End Sub

Sub LoadArrows()
Dim FileName As String
Dim I As Long

    Call CheckArrows
    
    FileName = App.Path & "\Arrows.ini"
    
    For I = 1 To MAX_ARROWS
        temp = I / MAX_ARROWS * 100
        Call SetStatus("Loading Arrows... " & temp & "%")
        Arrows(I).Name = GetVar(FileName, "Arrow" & I, "ArrowName")
        Arrows(I).Pic = GetVar(FileName, "Arrow" & I, "ArrowPic")
        Arrows(I).Range = GetVar(FileName, "Arrow" & I, "ArrowRange")
        Arrows(I).Amount = GetVar(FileName, "Arrow" & I, "ArrowAmount")
        Arrows(I).Ammo = GetVar(FileName, "Arrow" & I, "ArrowAmmo")
        
        DoEvents
    Next I
End Sub

Sub CheckArrows()
    If Not FileExist("Arrows.ini") Then
        Dim I As Long
    
        For I = 1 To MAX_ARROWS
            temp = I / MAX_ARROWS * 100
            Call SetStatus("Saving arrows... " & temp & "%")
            DoEvents
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & I, "ArrowName", "")
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & I, "ArrowPic", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & I, "ArrowRange", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & I, "ArrowAmount", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & I, "ArrowAmmo", 0)
        Next I
    End If
End Sub

Sub ClearArrows()
Dim I As Long

    For I = 1 To MAX_ARROWS
        Arrows(I).Name = ""
        Arrows(I).Pic = 0
        Arrows(I).Range = 0
        Arrows(I).Amount = 0
        Arrows(I).Ammo = 0
    Next I
End Sub

Sub SaveArrow(ByVal ArrowNum As Long)
Dim FileName As String

    FileName = App.Path & "\Arrows.ini"
    
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowName", Trim(Arrows(ArrowNum).Name))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowPic", Val(Arrows(ArrowNum).Pic))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowRange", Val(Arrows(ArrowNum).Range))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowAmount", Val(Arrows(ArrowNum).Amount))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowAmmo", Val(Arrows(ArrowNum).Ammo))
End Sub
