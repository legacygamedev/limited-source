Attribute VB_Name = "modDatabase"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public START_MAP As Long
Public START_X As Long
Public START_Y As Long

Public Const ADMIN_LOG = "admin.txt"
Public Const PLAYER_LOG = "player.txt"

Function GetVar(File As String, Header As String, Var As String) As String
On Error GoTo ErrorHandler
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modDatabase.bas", "GetVar", Err.Number, Err.Description
End Function

Sub PutVar(File As String, Header As String, Var As String, Value As String)
On Error GoTo ErrorHandler
    Call WritePrivateProfileString(Header, Var, Value, File)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "PutVar", Err.Number, Err.Description
End Sub

Sub LoadExps()
On Error GoTo ErrorHandler
Dim FileName As String
Dim i As Long

    Call CheckExps
    
    FileName = App.Path & "\experience.ini"
    
    For i = 1 To MAX_LEVEL
        Call SetStatus("Loading exp... " & i & "/" & MAX_LEVEL)
        Experience(i) = GetVar(FileName, "EXPERIENCE", "Exp" & i)
        
        DoEvents
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "LoadExps", Err.Number, Err.Description
End Sub

Sub CheckExps()
On Error GoTo ErrorHandler
    If Not FileExist("experience.ini") Then
        Dim i As Long
    
        For i = 1 To MAX_LEVEL
            Call SetStatus("Saving exp... " & i & "/" & MAX_LEVEL)
            DoEvents
            Call PutVar(App.Path & "\experience.ini", "EXPERIENCE", "Exp" & i, i * 1500)
        Next i
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "CheckExps", Err.Number, Err.Description
End Sub

Sub ClearExps()
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_LEVEL
        Experience(i) = 0
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "ClearExps", Err.Number, Err.Description
End Sub

Sub LoadEmos()
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "LoadEmos", Err.Number, Err.Description
End Sub

Sub CheckEmos()
On Error GoTo ErrorHandler
    If Not FileExist("emoticons.ini") Then
        Dim i As Long
    
        For i = 0 To MAX_EMOTICONS
            Call SetStatus("Saving emoticons... " & i & "/" & MAX_EMOTICONS)
            DoEvents
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "Emoticon" & i, 0)
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "EmoticonC" & i, "")
        Next i
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "CheckEmos", Err.Number, Err.Description
End Sub

Sub ClearEmos()
On Error GoTo ErrorHandler
Dim i As Long

    For i = 0 To MAX_EMOTICONS
        Emoticons(i).Pic = 0
        Emoticons(i).Command = ""
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "ClearEmos", Err.Number, Err.Description
End Sub

Sub SaveEmoticon(ByVal EmoNum As Long)
On Error GoTo ErrorHandler
Dim FileName As String

    FileName = App.Path & "\emoticons.ini"
    
    Call PutVar(FileName, "EMOTICONS", "EmoticonC" & EmoNum, Trim(Emoticons(EmoNum).Command))
    Call PutVar(FileName, "EMOTICONS", "Emoticon" & EmoNum, Trim(Emoticons(EmoNum).Pic))
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveEmoticon", Err.Number, Err.Description
End Sub

Function FileExist(ByVal FileName As String) As Boolean
On Error GoTo ErrorHandler
    If Dir(App.Path & "\" & FileName) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modDatabase.bas", "FileExist", Err.Number, Err.Description
End Function

Sub SavePlayer(ByVal index As Long, ByVal CharNum As Long)
On Error GoTo ErrorHandler

Dim FileName As String
Dim i As Long
Dim n As Long
Dim c As Long

    FileName = App.Path & "\accounts\" & Trim(Player(index).Login) & ".ini"
    
    Call PutVar(FileName, "GENERAL", "Login", Trim(Player(index).Login))
    Call PutVar(FileName, "GENERAL", "Password", Trim(Player(index).Password))

    i = CharNum
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
    Call PutVar(FileName, "CHAR" & i, "SPEED", STR(Player(index).Char(i).SPEED))
    Call PutVar(FileName, "CHAR" & i, "MAGI", STR(Player(index).Char(i).MAGI))
    Call PutVar(FileName, "CHAR" & i, "VIT", STR(Player(index).Char(i).VIT))
    Call PutVar(FileName, "CHAR" & i, "POINTS", STR(Player(index).Char(i).POINTS))
    
    ' Worn equipment
    Call PutVar(FileName, "CHAR" & i, "ArmorSlot", STR(Player(index).Char(i).ArmorSlot))
    Call PutVar(FileName, "CHAR" & i, "WeaponSlot", STR(Player(index).Char(i).WeaponSlot))
    Call PutVar(FileName, "CHAR" & i, "HelmetSlot", STR(Player(index).Char(i).HelmetSlot))
    Call PutVar(FileName, "CHAR" & i, "ShieldSlot", STR(Player(index).Char(i).ShieldSlot))
    Call PutVar(FileName, "CHAR" & i, "BootsSlot", STR(Player(index).Char(i).BootsSlot))
    Call PutVar(FileName, "CHAR" & i, "GlovesSlot", STR(Player(index).Char(i).GlovesSlot))
    Call PutVar(FileName, "CHAR" & i, "RingSlot", STR(Player(index).Char(i).RingSlot))
    Call PutVar(FileName, "CHAR" & i, "AmuletSlot", STR(Player(index).Char(i).AmuletSlot))
    
    ' Check to make sure that they aren't on map 0, if so reset'm
    If Player(index).Char(i).Map = 0 Then
        Player(index).Char(i).Map = START_MAP
        Player(index).Char(i).X = START_X
        Player(index).Char(i).Y = START_Y
    End If
        
    ' Position
    Call PutVar(FileName, "CHAR" & i, "Map", STR(Player(index).Char(i).Map))
    Call PutVar(FileName, "CHAR" & i, "X", STR(Player(index).Char(i).X))
    Call PutVar(FileName, "CHAR" & i, "Y", STR(Player(index).Char(i).Y))
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
    
    FileName = App.Path & "\banks\" & Trim(Player(index).Login) & ".ini"
    ' Bank
    For n = 1 To MAX_BANK
        Call PutVar(FileName, "CHAR" & i, "BankItemNum" & n, STR(Player(index).Char(i).Bank(n).Num))
        Call PutVar(FileName, "CHAR" & i, "BankItemVal" & n, STR(Player(index).Char(i).Bank(n).Value))
        Call PutVar(FileName, "CHAR" & i, "BankItemDur" & n, STR(Player(index).Char(i).Bank(n).Dur))
    Next n
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SavePlayer", Err.Number, Err.Description
End Sub

Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
On Error GoTo ErrorHandler

Dim FileName As String
Dim i As Long
Dim n As Long

    Call ClearPlayer(index)
    
    FileName = App.Path & "\accounts\" & Trim(Name) & ".ini"

    Player(index).Login = GetVar(FileName, "GENERAL", "Login")
    Player(index).Password = GetVar(FileName, "GENERAL", "Password")

    For i = 1 To MAX_CHARS
        FileName = App.Path & "\accounts\" & Trim(Name) & ".ini"
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
        Player(index).Char(i).SPEED = Val(GetVar(FileName, "CHAR" & i, "SPEED"))
        Player(index).Char(i).MAGI = Val(GetVar(FileName, "CHAR" & i, "MAGI"))
        Player(index).Char(i).VIT = Val(GetVar(FileName, "CHAR" & i, "VIT"))
        Player(index).Char(i).POINTS = Val(GetVar(FileName, "CHAR" & i, "POINTS"))
        
        ' Worn equipment
        Player(index).Char(i).ArmorSlot = Val(GetVar(FileName, "CHAR" & i, "ArmorSlot"))
        Player(index).Char(i).WeaponSlot = Val(GetVar(FileName, "CHAR" & i, "WeaponSlot"))
        Player(index).Char(i).HelmetSlot = Val(GetVar(FileName, "CHAR" & i, "HelmetSlot"))
        Player(index).Char(i).ShieldSlot = Val(GetVar(FileName, "CHAR" & i, "ShieldSlot"))
        Player(index).Char(i).BootsSlot = Val(GetVar(FileName, "CHAR" & i, "BootsSlot"))
        Player(index).Char(i).GlovesSlot = Val(GetVar(FileName, "CHAR" & i, "GlovesSlot"))
        Player(index).Char(i).RingSlot = Val(GetVar(FileName, "CHAR" & i, "RingSlot"))
        Player(index).Char(i).AmuletSlot = Val(GetVar(FileName, "CHAR" & i, "AmuletSlot"))
        
        ' Position
        Player(index).Char(i).Map = Val(GetVar(FileName, "CHAR" & i, "Map"))
        Player(index).Char(i).X = Val(GetVar(FileName, "CHAR" & i, "X"))
        Player(index).Char(i).Y = Val(GetVar(FileName, "CHAR" & i, "Y"))
        Player(index).Char(i).Dir = Val(GetVar(FileName, "CHAR" & i, "Dir"))
        
        
        'Binding
        Player(index).Char(i).Binding.Map = START_MAP
        Player(index).Char(i).Binding.X = START_X
        Player(index).Char(i).Binding.Y = START_Y
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(index).Char(i).Map = 0 Then
            Player(index).Char(i).Map = START_MAP
            Player(index).Char(i).X = START_X
            Player(index).Char(i).Y = START_Y
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
        
        FileName = App.Path & "\banks\" & Trim(Name) & ".ini"
        ' Bank
        For n = 1 To MAX_BANK
            Player(index).Char(i).Bank(n).Num = Val(GetVar(FileName, "CHAR" & i, "BankItemNum" & n))
            Player(index).Char(i).Bank(n).Value = Val(GetVar(FileName, "CHAR" & i, "BankItemVal" & n))
            Player(index).Char(i).Bank(n).Dur = Val(GetVar(FileName, "CHAR" & i, "BankItemDur" & n))
        Next n
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "LoadPlayer", Err.Number, Err.Description
End Sub

Function AccountExist(ByVal Name As String) As Boolean
On Error GoTo ErrorHandler
Dim FileName As String

    FileName = "accounts\" & Trim(Name) & ".ini"
    
    If FileExist(FileName) Then
        AccountExist = True
    Else
        AccountExist = False
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modDatabase.bas", "AccountExist", Err.Number, Err.Description
End Function

Function CharExist(ByVal index As Long, ByVal CharNum As Long) As Boolean
On Error GoTo ErrorHandler
    If Trim(Player(index).Char(CharNum).Name) <> "" Then
        CharExist = True
    Else
        CharExist = False
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modDatabase.bas", "CharExist", Err.Number, Err.Description
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modDatabase.bas", "PasswordOK", Err.Number, Err.Description
End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String)
On Error GoTo ErrorHandler

Dim i As Long

    Player(index).Login = Name
    Player(index).Password = Password
    
    For i = 1 To MAX_CHARS
        Call ClearChar(index, i)
    Next i
    
    For i = 1 To MAX_CHARS
        Call SavePlayer(index, i)
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "AddAccount", Err.Number, Err.Description
End Sub

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
On Error GoTo ErrorHandler

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
        Player(index).Char(CharNum).SPEED = Class(ClassNum).SPEED
        Player(index).Char(CharNum).MAGI = Class(ClassNum).MAGI
        Player(index).Char(CharNum).VIT = Class(ClassNum).VIT
        
        If Class(ClassNum).Map <= 0 Then Class(ClassNum).Map = 1
        If Class(ClassNum).X < 0 Or Class(ClassNum).X > MAX_MAPX Then Class(ClassNum).X = Int(Class(ClassNum).X / 2)
        If Class(ClassNum).Y < 0 Or Class(ClassNum).Y > MAX_MAPY Then Class(ClassNum).Y = Int(Class(ClassNum).Y / 2)
        Player(index).Char(CharNum).Map = Class(ClassNum).Map
        Player(index).Char(CharNum).X = Class(ClassNum).X
        Player(index).Char(CharNum).Y = Class(ClassNum).Y
            
        Player(index).Char(CharNum).HP = GetPlayerMaxHP(index)
        Player(index).Char(CharNum).MP = GetPlayerMaxMP(index)
        Player(index).Char(CharNum).SP = GetPlayerMaxSP(index)
                
        ' Append name to file
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Append As #f
            Print #f, Name
        Close #f
        
        Call SavePlayer(index, CharNum)
            
        Exit Sub
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "AddChar", Err.Number, Err.Description
End Sub

Sub DelChar(ByVal index As Long, ByVal CharNum As Long)
On Error GoTo ErrorHandler

Dim f1 As Long, f2 As Long
Dim s As String

    Call DeleteName(Player(index).Char(CharNum).Name)
    Call ClearChar(index, CharNum)
    Call SavePlayer(index, CharNum)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "DelChar", Err.Number, Err.Description
End Sub

Function FindChar(ByVal Name As String) As Boolean
On Error GoTo ErrorHandler

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
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modDatabase.bas", "FindChar", Err.Number, Err.Description
End Function

Function FindHDSerial(ByVal HDS As Long) As Boolean
On Error GoTo ErrorHandler

Dim f As Long
Dim s As Long

    If Not FileExist("banHDlist.txt") Then
        f = FreeFile
        Open App.Path & "\banHDlist.txt" For Output As #f
        Close #f
    End If

    FindHDSerial = False
    
    f = FreeFile
    Open App.Path & "\banHDlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            
            If Trim(Val(s)) = Trim(Val(HDS)) Then
                FindHDSerial = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modDatabase.bas", "FindHDSerial", Err.Number, Err.Description
End Function

Sub SaveAllPlayersOnline()
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i, Player(i).CharNum)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveAllPlayersOnline", Err.Number, Err.Description
End Sub

Sub LoadClasses()
On Error GoTo ErrorHandler

Dim FileName As String
Dim i As Long

    Call CheckClasses
    
    FileName = App.Path & "\classes.ini"
    
    Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class(0 To Max_Classes) As ClassRec
    
    Call ClearClasses
    
    For i = 0 To Max_Classes
        Call SetStatus("Loading classes... " & i & "/" & Max_Classes)
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).MaleSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        Class(i).FemaleSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        Class(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class(i).SPEED = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class(i).MAGI = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
        Class(i).VIT = Val(GetVar(FileName, "CLASS" & i, "VIT"))
        Class(i).Map = Val(GetVar(FileName, "CLASS" & i, "MAP"))
        Class(i).X = Val(GetVar(FileName, "CLASS" & i, "X"))
        Class(i).Y = Val(GetVar(FileName, "CLASS" & i, "Y"))
        Class(i).Locked = Val(GetVar(FileName, "CLASS" & i, "Locked"))
        
        DoEvents
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "LoadClasses", Err.Number, Err.Description
End Sub

Sub SaveClasses()
On Error GoTo ErrorHandler

Dim FileName As String
Dim i As Long

    FileName = App.Path & "\classes.ini"
    
    For i = 0 To Max_Classes
        Call SetStatus("Saving classes... " & i & "/" & Max_Classes)
        DoEvents
        Call PutVar(FileName, "CLASS" & i, "Name", Trim(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "MaleSprite", STR(Class(i).MaleSprite))
        Call PutVar(FileName, "CLASS" & i, "FemaleSprite", STR(Class(i).FemaleSprite))
        Call PutVar(FileName, "CLASS" & i, "STR", STR(Class(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR(Class(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR(Class(i).SPEED))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR(Class(i).MAGI))
        Call PutVar(FileName, "CLASS" & i, "VIT", STR(Class(i).VIT))
        Call PutVar(FileName, "CLASS" & i, "MAP", STR(Class(i).Map))
        Call PutVar(FileName, "CLASS" & i, "X", STR(Class(i).X))
        Call PutVar(FileName, "CLASS" & i, "Y", STR(Class(i).Y))
        Call PutVar(FileName, "CLASS" & i, "Locked", STR(Class(i).Locked))
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveClasses", Err.Number, Err.Description
End Sub

Sub CheckClasses()
On Error GoTo ErrorHandler
    If Not FileExist("classes.ini") Then
        Call SaveClasses
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "CheckClasses", Err.Number, Err.Description
End Sub

Sub SaveItems()
On Error GoTo ErrorHandler

Dim i As Long
        
    Call SetStatus("Saving items... ")
    For i = 1 To MAX_ITEMS
        If Not FileExist("items\item" & i & ".dat") Then
            Call SetStatus("Saving items... " & i & "/" & MAX_ITEMS)
            DoEvents
            Call SaveItem(i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveItems", Err.Number, Err.Description
End Sub

Sub SaveItem(ByVal ItemNum As Long)
On Error GoTo ErrorHandler

Dim FileName As String
Dim f  As Long
FileName = App.Path & "\items\item" & ItemNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Item(ItemNum)
    Close #f
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveItem", Err.Number, Err.Description
End Sub

Sub LoadItems()
On Error GoTo ErrorHandler

Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckItems
    
    For i = 1 To MAX_ITEMS
        Call SetStatus("Loading items... " & i & "/" & MAX_ITEMS)
        
        FileName = App.Path & "\Items\Item" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Item(i)
        Close #f
        
        DoEvents
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "LoadItems", Err.Number, Err.Description
End Sub

Sub CheckItems()
On Error GoTo ErrorHandler
    Call SaveItems
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "CheckItems", Err.Number, Err.Description
End Sub

Sub SaveShops()
On Error GoTo ErrorHandler

Dim i As Long

    Call SetStatus("Saving shops... ")
    For i = 1 To MAX_SHOPS
        If Not FileExist("shops\shop" & i & ".dat") Then
            Call SetStatus("Saving shops... " & i & "/" & MAX_SHOPS)
            DoEvents
            Call SaveShop(i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveShops", Err.Number, Err.Description
End Sub

Sub SaveShop(ByVal ShopNum As Long)
On Error GoTo ErrorHandler

Dim FileName As String
Dim f As Long

    FileName = App.Path & "\shops\shop" & ShopNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Shop(ShopNum)
    Close #f
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveShop", Err.Number, Err.Description
End Sub

Sub LoadShops()
On Error GoTo ErrorHandler

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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "LoadShops", Err.Number, Err.Description
End Sub

Sub CheckShops()
On Error GoTo ErrorHandler
    Call SaveShops
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "CheckShops", Err.Number, Err.Description
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
On Error GoTo ErrorHandler
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\spells\spells" & SpellNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Spell(SpellNum)
    Close #f
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveSpell", Err.Number, Err.Description
End Sub

Sub SaveSpells()
On Error GoTo ErrorHandler

Dim i As Long

    Call SetStatus("Saving spells... ")
    For i = 1 To MAX_SPELLS
        If Not FileExist("spells\spells" & i & ".dat") Then
            Call SetStatus("Saving spells... " & i & "/" & MAX_SPELLS)
            DoEvents
            Call SaveSpell(i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveSpells", Err.Number, Err.Description
End Sub

Sub LoadSpells()
On Error GoTo ErrorHandler

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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "LoadSpells", Err.Number, Err.Description
End Sub

Sub CheckSpells()
On Error GoTo ErrorHandler
    Call SaveSpells
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "CheckSpells", Err.Number, Err.Description
End Sub

Sub SaveNpcs()
On Error GoTo ErrorHandler

Dim i As Long

    Call SetStatus("Saving npcs... ")
    
    For i = 1 To MAX_NPCS
        If Not FileExist("npcs\npc" & i & ".dat") Then
            Call SetStatus("Saving npcs... " & i & "/" & MAX_NPCS)
            DoEvents
            Call SaveNpc(i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveNpcs", Err.Number, Err.Description
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
On Error GoTo ErrorHandler

Dim FileName As String
Dim f As Long
FileName = App.Path & "\npcs\npc" & NpcNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Npc(NpcNum)
    Close #f
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveNpc", Err.Number, Err.Description
End Sub

Sub LoadNpcs()
On Error GoTo ErrorHandler

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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "LoadNpcs", Err.Number, Err.Description
End Sub

Sub CheckNpcs()
On Error GoTo ErrorHandler
    Call SaveNpcs
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "CheckNpcs", Err.Number, Err.Description
End Sub

Sub SaveMap(ByVal MapNum As Long)
On Error GoTo ErrorHandler
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveMap", Err.Number, Err.Description
End Sub

Sub SaveMaps()
On Error GoTo ErrorHandler
Dim FileName As String
Dim i As Long
Dim f As Long

    For i = 1 To MAX_MAPS
        Call SetStatus("Saving maps... " & i & "/" & MAX_MAPS)
        DoEvents
        Call SaveMap(i)
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "SaveMaps", Err.Number, Err.Description
End Sub

Sub LoadMaps()
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "LoadMaps", Err.Number, Err.Description
End Sub

Sub CheckMaps()
On Error GoTo ErrorHandler
Dim FileName As String
Dim X As Long
Dim Y As Long
Dim i As Long
Dim n As Long

    Call ClearMaps
        
    For i = 1 To MAX_MAPS
        FileName = "maps\map" & i & ".dat"
        
        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(FileName) Then
            Call SetStatus("Saving maps... " & i & "/" & MAX_NPCS)
            DoEvents
            
            Call SaveMap(i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "CheckMaps", Err.Number, Err.Description
End Sub

Sub AddLog(ByVal Text As String, ByVal FN As String)
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "AddLog", Err.Number, Err.Description
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
On Error GoTo ErrorHandler
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
        Print #f, IP & "," & GetPlayerName(BannedByIndex) & "," & GetPlayerName(BanPlayerIndex)
    Close #f
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "BanIndex", Err.Number, Err.Description
End Sub

Sub BanHDIndex(ByVal BanPlayerIndex As Long, ByVal HDSerial As Long, ByVal BannedByIndex As Long)
On Error GoTo ErrorHandler
Dim FileName, IP As String
Dim f As Long, i As Long

    FileName = App.Path & "\banHDlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("banHDlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
                
    f = FreeFile
    Open FileName For Append As #f
        Print #f, Trim(Val(HDSerial))
    Close #f
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "BanHDIndex", Err.Number, Err.Description
End Sub

Sub DeleteName(ByVal Name As String)
On Error GoTo ErrorHandler
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
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modDatabase.bas", "DeleteName", Err.Number, Err.Description
End Sub
