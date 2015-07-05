Attribute VB_Name = "modDatabase"
Option Explicit
Public Size As Byte

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public START_MAP As Long
Public START_X As Long
Public START_Y As Long

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

Sub PutVar(File As String, Header As String, Var As String, value As String)
    Call WritePrivateProfileString(Header, Var, value, File)
End Sub

Sub LoadExps()
Dim FileName As String
Dim i As Long

    Call CheckExps
    
    FileName = App.Path & "\experience.ini"
    
    For i = 1 To MAX_LEVEL
        Call SetStatus("Loading exp... " & i & "/" & MAX_LEVEL)
        Experience(i) = GetVar(FileName, "EXPERIENCE", "Exp" & i)
        
        DoEvents
    Next i
End Sub

Sub CheckExps()
    If Not FileExist("experience.ini") Then
        Dim i As Long
    
        For i = 1 To MAX_LEVEL
            Call SetStatus("Saving exp... " & i & "/" & MAX_LEVEL)
            DoEvents
            Call PutVar(App.Path & "\experience.ini", "EXPERIENCE", "Exp" & i, i * 1500)
        Next i
    End If
End Sub

Sub ClearExps()
Dim i As Long

    For i = 1 To MAX_LEVEL
        Experience(i) = 0
    Next i
End Sub

Sub LoadEmos()
Dim FileName As String
Dim i As Long

    Call CheckEmos
    
    FileName = App.Path & "\emoticons.ini"
    
    For i = 0 To MAX_EMOTICONS
        Call SetStatus("Loading emoticons... " & i & "/" & MAX_EMOTICONS)
        Emoticons(i).Pic = GetVar(FileName, "EMOTICONS", "Emoticon" & i)
        Emoticons(i).Command = GetVar(FileName, "EMOTICONS", "EmoticonC" & i)
        
        DoEvents
    Next i
End Sub

Sub CheckEmos()
    If Not FileExist("emoticons.ini") Then
        Dim i As Long
    
        For i = 0 To MAX_EMOTICONS
            Call SetStatus("Saving emoticons... " & i & "/" & MAX_EMOTICONS)
            DoEvents
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "Emoticon" & i, 0)
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "EmoticonC" & i, "")
        Next i
    End If
End Sub

Sub ClearEmos()
Dim i As Long

    For i = 0 To MAX_EMOTICONS
        Emoticons(i).Pic = 0
        Emoticons(i).Command = ""
    Next i
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
        Call PutVar(FileName, "CHAR" & i, "Guild", Trim(Player(index).Char(i).Guild))
        Call PutVar(FileName, "CHAR" & i, "Guildaccess", STR(Player(index).Char(i).Guildaccess))

        
        ' Vitals
        Call PutVar(FileName, "CHAR" & i, "HP", STR(Player(index).Char(i).HP))
        Call PutVar(FileName, "CHAR" & i, "MP", STR(Player(index).Char(i).MP))
        Call PutVar(FileName, "CHAR" & i, "SP", STR(Player(index).Char(i).SP))
        
        ' Stats
        Call PutVar(FileName, "CHAR" & i, "STR", STR(Player(index).Char(i).STR))
        Call PutVar(FileName, "CHAR" & i, "DEF", STR(Player(index).Char(i).DEF))
        Call PutVar(FileName, "CHAR" & i, "SPEED", STR(Player(index).Char(i).Speed))
        Call PutVar(FileName, "CHAR" & i, "MAGI", STR(Player(index).Char(i).Magi))
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
            Call PutVar(FileName, "CHAR" & i, "InvItemNum" & n, STR(Player(index).Char(i).Inv(n).num))
            Call PutVar(FileName, "CHAR" & i, "InvItemVal" & n, STR(Player(index).Char(i).Inv(n).value))
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
        Player(index).Char(i).Guild = GetVar(FileName, "CHAR" & i, "Guild")
        Player(index).Char(i).Guildaccess = Val(GetVar(FileName, "CHAR" & i, "Guildaccess"))
        
        ' Vitals
        Player(index).Char(i).HP = Val(GetVar(FileName, "CHAR" & i, "HP"))
        Player(index).Char(i).MP = Val(GetVar(FileName, "CHAR" & i, "MP"))
        Player(index).Char(i).SP = Val(GetVar(FileName, "CHAR" & i, "SP"))
        
        ' Stats
        Player(index).Char(i).STR = Val(GetVar(FileName, "CHAR" & i, "STR"))
        Player(index).Char(i).DEF = Val(GetVar(FileName, "CHAR" & i, "DEF"))
        Player(index).Char(i).Speed = Val(GetVar(FileName, "CHAR" & i, "SPEED"))
        Player(index).Char(i).Magi = Val(GetVar(FileName, "CHAR" & i, "MAGI"))
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
            Player(index).Char(i).Inv(n).num = Val(GetVar(FileName, "CHAR" & i, "InvItemNum" & n))
            Player(index).Char(i).Inv(n).value = Val(GetVar(FileName, "CHAR" & i, "InvItemVal" & n))
            Player(index).Char(i).Inv(n).Dur = Val(GetVar(FileName, "CHAR" & i, "InvItemDur" & n))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Player(index).Char(i).Spell(n) = Val(GetVar(FileName, "CHAR" & i, "Spell" & n))
        Next n
                For n = 1 To 49
           Player(index).Char(i).bank(n).num = Val(GetVar(FileName, "CHAR" & i, "BankItemNum" & n))
            Player(index).Char(i).bank(n).value = Val(GetVar(FileName, "CHAR" & i, "BankItemVal" & n))
            Player(index).Char(i).bank(n).Dur = Val(GetVar(FileName, "CHAR" & i, "BankItemDur" & n))
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
            Player(index).Char(CharNum).Sprite = Class(ClassNum).MaleSprite
        Else
            Player(index).Char(CharNum).Sprite = Class(ClassNum).FemaleSprite
        End If
        
        Player(index).Char(CharNum).Level = 1
                    
        Player(index).Char(CharNum).STR = Class(ClassNum).STR
        Player(index).Char(CharNum).DEF = Class(ClassNum).DEF
        Player(index).Char(CharNum).Speed = Class(ClassNum).Speed
        Player(index).Char(CharNum).Magi = Class(ClassNum).Magi
        
        If Class(ClassNum).Map <= 0 Then Class(ClassNum).Map = 1
        If Class(ClassNum).x < 0 Or Class(ClassNum).x > MAX_MAPX Then Class(ClassNum).x = Int(Class(ClassNum).x / 2)
        If Class(ClassNum).y < 0 Or Class(ClassNum).y > MAX_MAPY Then Class(ClassNum).y = Int(Class(ClassNum).y / 2)
        Player(index).Char(CharNum).Map = Class(ClassNum).Map
        Player(index).Char(CharNum).x = Class(ClassNum).x
        Player(index).Char(CharNum).y = Class(ClassNum).y
            
        Player(index).Char(CharNum).HP = GetPlayerMaxHP(index)
        Player(index).Char(CharNum).MP = GetPlayerMaxMP(index)
        Player(index).Char(CharNum).SP = GetPlayerMaxSP(index)
                
        Call PutVar(App.Path & "\defencespells.ini", Name, "StrengthUse", "0")
        Call PutVar(App.Path & "\defencespells.ini", Name, "DefenceUse", "0")
        Call PutVar(App.Path & "\defencespells.ini", Name, "AgilityUse", "0")
        Call PutVar(App.Path & "\defencespells.ini", Name, "WisdomUse", "0")
                
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
Dim S As String

    Call DeleteName(Player(index).Char(CharNum).Name)
    Call ClearChar(index, CharNum)
    Call SavePlayer(index)
End Sub

Function FindChar(ByVal Name As String) As Boolean
Dim f As Long
Dim S As String

    FindChar = False
    
    f = FreeFile
    Open App.Path & "\accounts\charlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, S
            
            If Trim(LCase(S)) = Trim(LCase(Name)) Then
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
    
    FileName = App.Path & "\Classes\info.ini"
    
    Max_Classes = Val(GetVar(FileName, "INFO", "MaxClasses"))
    
    ReDim Class(0 To Max_Classes) As ClassRec
    
    Call ClearClasses
    
    For i = 0 To Max_Classes
        Call SetStatus("Loading classes... " & i & "/" & Max_Classes)
        FileName = App.Path & "\Classes\Class" & i & ".ini"
        Class(i).Name = GetVar(FileName, "CLASS", "Name")
        Class(i).MaleSprite = GetVar(FileName, "CLASS", "MaleSprite")
        Class(i).FemaleSprite = GetVar(FileName, "CLASS", "FemaleSprite")
        Class(i).STR = Val(GetVar(FileName, "CLASS", "STR"))
        Class(i).DEF = Val(GetVar(FileName, "CLASS", "DEF"))
        Class(i).Speed = Val(GetVar(FileName, "CLASS", "SPEED"))
        Class(i).Magi = Val(GetVar(FileName, "CLASS", "MAGI"))
        Class(i).Map = Val(GetVar(FileName, "CLASS", "MAP"))
        Class(i).x = Val(GetVar(FileName, "CLASS", "X"))
        Class(i).y = Val(GetVar(FileName, "CLASS", "Y"))
        Class(i).Locked = Val(GetVar(FileName, "CLASS", "Locked"))
        
        DoEvents
    Next i
End Sub

Sub SaveClasses()
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\Classes\info.ini"
    
    If Not FileExist("Classes\info.ini") Then
        Call PutVar(FileName, "INFO", "MaxClasses", 0)
        Call PutVar(FileName, "INFO", "MaxSkills", 25)
        Call PutVar(FileName, "INFO", "StatPoints", 0)
        Call PutVar(FileName, "INFO", "SkillPoints", 0)
    End If
    
    For i = 0 To Max_Classes
        Call SetStatus("Saving classes... " & i & "/" & Max_Classes)
        DoEvents
        FileName = App.Path & "\Classes\Class" & i & ".ini"
        If Not FileExist("Classes\Class" & i & ".ini") Then
            Call PutVar(FileName, "CLASS", "Name", Trim(Class(i).Name))
            Call PutVar(FileName, "CLASS", "MaleSprite", STR(Class(i).MaleSprite))
            Call PutVar(FileName, "CLASS", "FemaleSprite", STR(Class(i).FemaleSprite))
            Call PutVar(FileName, "CLASS", "STR", STR(Class(i).STR))
            Call PutVar(FileName, "CLASS", "DEF", STR(Class(i).DEF))
            Call PutVar(FileName, "CLASS", "SPEED", STR(Class(i).Speed))
            Call PutVar(FileName, "CLASS", "MAGI", STR(Class(i).Magi))
            Call PutVar(FileName, "CLASS", "MAP", STR(Class(i).Map))
            Call PutVar(FileName, "CLASS", "X", STR(Class(i).x))
            Call PutVar(FileName, "CLASS", "Y", STR(Class(i).y))
            Call PutVar(FileName, "CLASS", "Locked", STR(Class(i).Locked))
        End If
    Next i
End Sub

Sub CheckClasses()
    If Not FileExist("Classes\info.ini") Then
        Call SaveClasses
    End If
End Sub

Sub LoadClasses2()
Dim FileName As String
Dim i As Long

    Call CheckClasses2
    
    FileName = App.Path & "\FirstClassAdvancement.ini"
    
    Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class2(0 To Max_Classes) As ClassRec
    
    Call ClearClasses2
    
    For i = 0 To Max_Classes
        Call SetStatus("Loading first class advandcement... " & i & "/" & Max_Classes)
        Class2(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class2(i).AdvanceFrom = Val(GetVar(FileName, "CLASS" & i, "AdvanceFrom"))
        Class2(i).MaleSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        Class2(i).FemaleSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        Class2(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class2(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class2(i).Speed = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class2(i).Magi = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
        
        DoEvents
    Next i
End Sub

Sub SaveClasses2()
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\FirstClassAdvancement.ini"
    
    For i = 0 To Max_Classes
        Call SetStatus("Saving first class advandcement... " & i & "/" & Max_Classes)
        DoEvents
        Call PutVar(FileName, "CLASS" & i, "Name", Trim(Class2(i).Name))
        Call PutVar(FileName, "CLASS" & i, "AdvanceFrom", STR(Class2(i).AdvanceFrom))
        Call PutVar(FileName, "CLASS" & i, "LevelReq", STR(Class2(i).LevelReq))
        Call PutVar(FileName, "CLASS" & i, "MaleSprite", STR(Class2(i).MaleSprite))
        Call PutVar(FileName, "CLASS" & i, "FemaleSprite", STR(Class2(i).FemaleSprite))
        Call PutVar(FileName, "CLASS" & i, "STR", STR(Class2(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR(Class2(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR(Class2(i).Speed))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR(Class2(i).Magi))
    Next i
End Sub

Sub CheckClasses2()
    If Not FileExist("FirstClassAdvancement.ini") Then
        Call SaveClasses2
    End If
End Sub

Sub Loadclasses3()
Dim FileName As String
Dim i As Long

    Call Checkclasses3
    
    FileName = App.Path & "\SecondClassAdvancement.ini"
    
    Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class3(0 To Max_Classes) As ClassRec
    
    Call ClearClasses3
    
    For i = 0 To Max_Classes
        Call SetStatus("Loading second class advandcement... " & i & "/" & Max_Classes)
        Class3(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class3(i).AdvanceFrom = Val(GetVar(FileName, "CLASS" & i, "AdvanceFrom"))
        Class3(i).LevelReq = Val(GetVar(FileName, "CLASS" & i, "LevelReq"))
        Class3(i).MaleSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        Class3(i).FemaleSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        Class3(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class3(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class3(i).Speed = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class3(i).Magi = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
        
        DoEvents
    Next i
End Sub

Sub Saveclasses3()
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\SecondClassAdvancement.ini"
    
    For i = 0 To Max_Classes
        Call SetStatus("Saving second class advandcement... " & i & "/" & Max_Classes)
        DoEvents
        Call PutVar(FileName, "CLASS" & i, "Name", Trim(Class3(i).Name))
        Call PutVar(FileName, "CLASS" & i, "AdvanceFrom", STR(Class3(i).AdvanceFrom))
        Call PutVar(FileName, "CLASS" & i, "MaleSprite", STR(Class3(i).MaleSprite))
        Call PutVar(FileName, "CLASS" & i, "FemaleSprite", STR(Class3(i).FemaleSprite))
        Call PutVar(FileName, "CLASS" & i, "STR", STR(Class3(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR(Class3(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR(Class3(i).Speed))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR(Class3(i).Magi))
    Next i
End Sub

Sub Checkclasses3()
    If Not FileExist("SecondClassAdvancement.ini") Then
        Call Saveclasses3
    End If
End Sub

Sub SaveItems()
Dim i As Long
        
    Call SetStatus("Saving items... ")
    For i = 1 To MAX_ITEMS
        If Not FileExist("items\item" & i & ".dat") Then
            Call SetStatus("Saving items... " & i & "/" & MAX_ITEMS)
            DoEvents
            Call SaveItem(i)
        End If
    Next i
End Sub

Sub SaveItem(ByVal ItemNum As Long)
Dim FileName As String
Dim f  As Long
FileName = App.Path & "\items\item" & ItemNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , item(ItemNum)
    Close #f
End Sub

Sub LoadItems()
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckItems
    
    For i = 1 To MAX_ITEMS
        Call SetStatus("Loading items... " & i & "/" & MAX_ITEMS)
        
        FileName = App.Path & "\Items\Item" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , item(i)
        Close #f
        
        DoEvents
    Next i
End Sub

Sub CheckItems()
    Call SaveItems
End Sub

Sub SaveShops()
Dim i As Long

    Call SetStatus("Saving shops... ")
    For i = 1 To MAX_SHOPS
        If Not FileExist("shops\shop" & i & ".dat") Then
            Call SetStatus("Saving shops... " & i & "/" & MAX_SHOPS)
            DoEvents
            Call SaveShop(i)
        End If
    Next i
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
Dim i As Long, f As Long

    Call CheckShops
    
    For i = 1 To MAX_SHOPS
        Call SetStatus("Loading shops... " & i & "/" & MAX_SHOPS)
        FileName = App.Path & "\shops\shop" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Shop(i)
        Close #f
        
        DoEvents
    Next i
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
Dim i As Long

    Call SetStatus("Saving spells... ")
    For i = 1 To MAX_SPELLS
        If Not FileExist("spells\spells" & i & ".dat") Then
            Call SetStatus("Saving spells... " & i & "/" & MAX_SPELLS)
            DoEvents
            Call SaveSpell(i)
        End If
    Next i
End Sub

Sub LoadSpells()
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckSpells
    
    For i = 1 To MAX_SPELLS
        Call SetStatus("Loading spells... " & i & "/" & MAX_SPELLS)
        
        FileName = App.Path & "\spells\spells" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Spell(i)
        Close #f
        
        DoEvents
    Next i
End Sub

Sub CheckSpells()
    Call SaveSpells
End Sub

Sub SaveNpcs()
Dim i As Long

    Call SetStatus("Saving npcs... ")
    
    For i = 1 To MAX_NPCS
        If Not FileExist("npcs\npc" & i & ".dat") Then
            Call SetStatus("Saving npcs... " & i & "/" & MAX_NPCS)
            DoEvents
            Call SaveNpc(i)
        End If
    Next i
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
Dim i As Long
Dim z As Long
Dim f As Long

    Call CheckNpcs
        
    For i = 1 To MAX_NPCS
        Call SetStatus("Loading npcs... " & i & "/" & MAX_NPCS)
        FileName = App.Path & "\npcs\npc" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Npc(i)
        Close #f
        
        DoEvents
    Next i
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
Dim i As Long
Dim f As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        Call SetStatus("Loading maps... " & i & "/" & MAX_MAPS)
        FileName = App.Path & "\maps\map" & i & ".dat"
        
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Map(i)
        Close #f
    
        DoEvents
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
        FileName = "maps\map" & i & ".dat"
        
        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(FileName) Then
            Call SetStatus("Saving maps... " & i & "/" & MAX_MAPS)
            DoEvents
            
            Call SaveMap(i)
        End If
    Next i
End Sub

Sub AddLog(ByVal text As String, ByVal FN As String)
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
            Print #f, Time & ": " & text
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
Dim S As String

    Call FileCopy(App.Path & "\accounts\charlist.txt", App.Path & "\accounts\chartemp.txt")
    
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\accounts\charlist.txt" For Output As #f2
        
    Do While Not EOF(f1)
        Input #f1, S
        If Trim(LCase(S)) <> Trim(LCase(Name)) Then
            Print #f2, S
        End If
    Loop
    
    Close #f1
    Close #f2
    
    Call Kill(App.Path & "\accounts\chartemp.txt")
End Sub

Sub BanByServer(ByVal BanPlayerIndex As Long, ByVal Reason As String)
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
    strWord = Mid(strWord, 1, a - 1) & strReplace & Right(strWord, Len(strWord) - a - charAmount + 1)
    Replace = strWord
End Function

Sub SaveLogs()
Dim FileName As String
Dim i As String, C As String

    If LCase(Dir(App.Path & "\logs", vbDirectory)) <> "logs" Then
        Call MkDir(App.Path & "\Logs")
    End If
    
    C = Time
    C = Replace(C, ":", ".", 1)
    C = Replace(C, ":", ".", 1)
    C = Replace(C, ":", ".", 1)
    
    i = Date
    i = Replace(i, "/", ".", 1)
    i = Replace(i, "/", ".", 1)
    i = Replace(i, "/", ".", 1)
    
    If LCase(Dir(App.Path & "\logs\" & i, vbDirectory)) <> i Then
        Call MkDir(App.Path & "\Logs\" & i & "\")
    End If
    
    If LCase(Dir(App.Path & "\logs\" & i & "\" & C, vbDirectory)) <> C Then
        Call MkDir(App.Path & "\Logs\" & i & "\" & C & "\")
    End If
        
    FileName = App.Path & "\Logs\" & i & "\" & C & "\Main.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(0).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & C & "\Broadcast.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(1).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & C & "\Global.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(2).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & C & "\Map.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(3).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & C & "\Private.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(4).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & C & "\Admin.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(5).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & C & "\Emote.txt"
    Open FileName For Output As #1
        Print #1, frmServer.txtText(6).text
    Close #1
End Sub

Sub LoadArrows()
Dim FileName As String
Dim i As Long

    Call CheckArrows
    
    FileName = App.Path & "\Arrows.ini"
    
    For i = 1 To MAX_ARROWS
        Call SetStatus("Loading Arrows... " & i & "/" & MAX_ARROWS)
        Arrows(i).Name = GetVar(FileName, "Arrow" & i, "ArrowName")
        Arrows(i).Pic = GetVar(FileName, "Arrow" & i, "ArrowPic")
        Arrows(i).Range = GetVar(FileName, "Arrow" & i, "ArrowRange")
        Arrows(i).amount = GetVar(FileName, "Arrow" & i, "ArrowAmount")

        DoEvents
    Next i
End Sub

Sub CheckArrows()
    If Not FileExist("Arrows.ini") Then
        Dim i As Long
    
        For i = 1 To MAX_ARROWS
            Call SetStatus("Saving arrows... " & i & "/" & MAX_ARROWS)
            DoEvents
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowName", "")
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowPic", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowRange", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowAmount", 0)
        Next i
    End If
End Sub

Sub ClearArrows()
Dim i As Long

    For i = 1 To MAX_ARROWS
        Arrows(i).Name = ""
        Arrows(i).Pic = 0
        Arrows(i).Range = 0
        Arrows(i).amount = 0
    Next i
End Sub

Sub SaveArrow(ByVal ArrowNum As Long)
Dim FileName As String

    FileName = App.Path & "\Arrows.ini"
    
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowName", Trim(Arrows(ArrowNum).Name))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowPic", Val(Arrows(ArrowNum).Pic))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowRange", Val(Arrows(ArrowNum).Range))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowAmount", Val(Arrows(ArrowNum).amount))
End Sub

