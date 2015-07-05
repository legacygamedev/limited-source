Attribute VB_Name = "modDatabase"
Option Explicit

Public Function GetVar(File As String, Header As String, Var As String) As String
'On Error GoTo errorhandler:
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modDatabase.bas", "", Err.Number, Err.Description)
End Function

Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
'On Error GoTo errorhandler:
    Call WritePrivateProfileString$(Header, Var, Value, File)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "PutVar", Err.Number, Err.Description)
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
'On Error GoTo errorhandler:
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/16/2005  Shannara   Optimized function.
'****************************************************************
    
    If RAW = False Then
        If Dir(App.Path & "\" & FileName) = "" Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
            Exit Function
        End If
    Else
        If Dir(FileName) = "" Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
        End If
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modDatabase.bas", "FileExist", Err.Number, Err.Description)
End Function

Sub SavePlayer(ByVal Index As Long)
'On Error GoTo errorhandler:
Dim FileName As String, f As Integer
Dim ClassNum As Byte
Dim START_MAP As Long, START_X As Long, START_Y As Long
Dim i As Long
Dim n As Long

    FileName = App.Path & "\Accounts\" & Trim$(Player(Index).Login) & ".ini"
    f = FreeFile
    
Open FileName For Output As #f
    Print #f, "[GENERAL]"
    Print #f, "Login=" & Trim$(Player(Index).Login)
    Print #f, "Password=" & Trim$(Player(Index).Password)
        
    For i = 1 To MAX_CHARS
        ClassNum = Player(Index).Char(i).Class
        START_MAP = Class(ClassNum).Map
        START_X = Class(ClassNum).x
        START_Y = Class(ClassNum).y
        
        Print #f, "[CHAR" & i & "]"
        
        ' General
        Print #f, "Name=" & Trim$(Player(Index).Char(i).Name)
        Print #f, "Class=" & STR(Player(Index).Char(i).Class)
        Print #f, "Sex=" & STR(Player(Index).Char(i).Sex)
        Print #f, "Sprite=" & STR(Player(Index).Char(i).Sprite)
        Print #f, "Level=" & STR(Player(Index).Char(i).Level)
        Print #f, "Exp=" & STR(Player(Index).Char(i).EXP)
        Print #f, "Access=" & STR(Player(Index).Char(i).Access)
        Print #f, "PK=" & STR(Player(Index).Char(i).PK)
        Print #f, "Guild=" & STR(Player(Index).Char(i).Guild)
        
        ' Vitals
        Print #f, "HP=" & STR(Player(Index).Char(i).HP)
        Print #f, "MP=" & STR(Player(Index).Char(i).MP)
        Print #f, "SP=" & STR(Player(Index).Char(i).SP)
        
        ' Stats
        Print #f, "STR=" & STR(Player(Index).Char(i).STR)
        Print #f, "DEF=" & STR(Player(Index).Char(i).DEF)
        Print #f, "SPEED=" & STR(Player(Index).Char(i).SPEED)
        Print #f, "MAGI=" & STR(Player(Index).Char(i).MAGI)
        Print #f, "POINTS=" & STR(Player(Index).Char(i).POINTS)
        
        ' Worn equipment
        Print #f, "ArmorSlot=" & STR(Player(Index).Char(i).ArmorSlot)
        Print #f, "WeaponSlot=" & STR(Player(Index).Char(i).WeaponSlot)
        Print #f, "HelmetSlot=" & STR(Player(Index).Char(i).HelmetSlot)
        Print #f, "ShieldSlot=" & STR(Player(Index).Char(i).ShieldSlot)
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(i).Map = 0 Then
            Player(Index).Char(i).Map = START_MAP
            Player(Index).Char(i).x = START_X
            Player(Index).Char(i).y = START_Y
        End If
            
        ' Position
        Print #f, "Map=" & STR(Player(Index).Char(i).Map)
        Print #f, "X=" & STR(Player(Index).Char(i).x)
        Print #f, "Y=" & STR(Player(Index).Char(i).y)
        Print #f, "Dir=" & STR(Player(Index).Char(i).Dir)
        
        'Set Player's Text
        Print #f, "Text=" & Player(Index).Char(i).Text
        
        'Set Friends
        For n = 1 To MAX_FRIENDS
            Print #f, "Friend" & n & "=" & Player(Index).Char(i).Friends(n)
        Next
        
        ' Inventory
        For n = 1 To MAX_INV
            Print #f, "InvItemNum" & n & "=" & STR(Player(Index).Char(i).Inv(n).Num)
            Print #f, "InvItemVal" & n & "=" & STR(Player(Index).Char(i).Inv(n).Value)
            Print #f, "InvItemDur" & n & "=" & STR(Player(Index).Char(i).Inv(n).Dur)
        Next n
        
        ' Bank
        For n = 1 To MAX_BANK_ITEMS
            Print #f, "BankItemNum" & n & "=" & STR(Player(Index).Char(i).BankInv(n).Num)
            Print #f, "BankItemVal" & n & "=" & STR(Player(Index).Char(i).BankInv(n).Value)
            Print #f, "BankItemDur" & n & "=" & STR(Player(Index).Char(i).BankInv(n).Dur)
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Print #f, "Spell" & n & "=" & STR(Player(Index).Char(i).Spell(n))
        Next n
    Next i
    
Close #f
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SavePlayer", Err.Number, Err.Description)
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
'On Error GoTo errorhandler:
Dim FileName As String
Dim ClassNum As Byte, START_MAP As Long, START_X As Long, START_Y As Long
Dim i As Long
Dim n As Long

    Call ClearPlayer(Index)
    
    FileName = App.Path & "\Accounts\" & Trim$(Name) & ".ini"

    Player(Index).Login = GetVar(FileName, "GENERAL", "Login")
    Player(Index).Password = GetVar(FileName, "GENERAL", "Password")

    For i = 1 To MAX_CHARS
        ClassNum = Player(Index).Char(i).Class
        START_MAP = Class(ClassNum).Map
        START_X = Class(ClassNum).x
        START_Y = Class(ClassNum).y
        
        ' General
        Player(Index).Char(i).Name = GetVar(FileName, "CHAR" & i, "Name")
        Player(Index).Char(i).Sex = Val(GetVar(FileName, "CHAR" & i, "Sex"))
        Player(Index).Char(i).Class = Val(GetVar(FileName, "CHAR" & i, "Class"))
        Player(Index).Char(i).Sprite = Val(GetVar(FileName, "CHAR" & i, "Sprite"))
        Player(Index).Char(i).Level = Val(GetVar(FileName, "CHAR" & i, "Level"))
        Player(Index).Char(i).EXP = Val(GetVar(FileName, "CHAR" & i, "Exp"))
        Player(Index).Char(i).Access = Val(GetVar(FileName, "CHAR" & i, "Access"))
        Player(Index).Char(i).PK = Val(GetVar(FileName, "CHAR" & i, "PK"))
        Player(Index).Char(i).Guild = Val(GetVar(FileName, "CHAR" & i, "Guild"))
        
        ' Vitals
        Player(Index).Char(i).HP = Val(GetVar(FileName, "CHAR" & i, "HP"))
        Player(Index).Char(i).MP = Val(GetVar(FileName, "CHAR" & i, "MP"))
        Player(Index).Char(i).SP = Val(GetVar(FileName, "CHAR" & i, "SP"))
        
        ' Stats
        Player(Index).Char(i).STR = Val(GetVar(FileName, "CHAR" & i, "STR"))
        Player(Index).Char(i).DEF = Val(GetVar(FileName, "CHAR" & i, "DEF"))
        Player(Index).Char(i).SPEED = Val(GetVar(FileName, "CHAR" & i, "SPEED"))
        Player(Index).Char(i).MAGI = Val(GetVar(FileName, "CHAR" & i, "MAGI"))
        Player(Index).Char(i).POINTS = Val(GetVar(FileName, "CHAR" & i, "POINTS"))
        
        ' Worn equipment
        Player(Index).Char(i).ArmorSlot = Val(GetVar(FileName, "CHAR" & i, "ArmorSlot"))
        Player(Index).Char(i).WeaponSlot = Val(GetVar(FileName, "CHAR" & i, "WeaponSlot"))
        Player(Index).Char(i).HelmetSlot = Val(GetVar(FileName, "CHAR" & i, "HelmetSlot"))
        Player(Index).Char(i).ShieldSlot = Val(GetVar(FileName, "CHAR" & i, "ShieldSlot"))
        
        ' Position
        Player(Index).Char(i).Map = Val(GetVar(FileName, "CHAR" & i, "Map"))
        Player(Index).Char(i).x = Val(GetVar(FileName, "CHAR" & i, "X"))
        Player(Index).Char(i).y = Val(GetVar(FileName, "CHAR" & i, "Y"))
        Player(Index).Char(i).Dir = Val(GetVar(FileName, "CHAR" & i, "Dir"))
        
        'Get Player's Text
        Player(Index).Char(i).Text = GetVar(FileName, "CHAR" & i, "Text")
        
        'Get Friends
        For n = 1 To MAX_FRIENDS
            Player(Index).Char(i).Friends(n) = GetVar(FileName, "CHAR" & i, "Friend" & n)
        Next
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(i).Map = 0 Then
            Player(Index).Char(i).Map = START_MAP
            Player(Index).Char(i).x = START_X
            Player(Index).Char(i).y = START_Y
        End If
        
        ' Inventory
        For n = 1 To MAX_INV
            Player(Index).Char(i).Inv(n).Num = Val(GetVar(FileName, "CHAR" & i, "InvItemNum" & n))
            Player(Index).Char(i).Inv(n).Value = Val(GetVar(FileName, "CHAR" & i, "InvItemVal" & n))
            Player(Index).Char(i).Inv(n).Dur = Val(GetVar(FileName, "CHAR" & i, "InvItemDur" & n))
        Next n
        
        ' Bank
        For n = 1 To MAX_BANK_ITEMS
            Player(Index).Char(i).BankInv(n).Num = Val(GetVar(FileName, "CHAR" & i, "BankItemNum" & n))
            Player(Index).Char(i).BankInv(n).Value = Val(GetVar(FileName, "CHAR" & i, "BankItemVal" & n))
            Player(Index).Char(i).BankInv(n).Dur = Val(GetVar(FileName, "CHAR" & i, "BankItemDur" & n))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(i).Spell(n) = Val(GetVar(FileName, "CHAR" & i, "Spell" & n))
        Next n
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "LoadPlayer", Err.Number, Err.Description)
End Sub

Function AccountExist(ByVal Name As String) As Boolean
'On Error GoTo errorhandler:
Dim FileName As String

    FileName = "Accounts\" & Trim$(Name) & ".ini"
    
    If FileExist(FileName) Then
        AccountExist = True
    Else
        AccountExist = False
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modDatabase.bas", "AccountExist", Err.Number, Err.Description)
End Function

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
'On Error GoTo errorhandler:
    If Trim$(Player(Index).Char(CharNum).Name) <> "" Then
        CharExist = True
    Else
        CharExist = False
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modDatabase.bas", "CharExist", Err.Number, Err.Description)
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
'On Error GoTo errorhandler:
Dim FileName As String
Dim RightPassword As String

    PasswordOK = False
    
    If AccountExist(Name) Then
        FileName = App.Path & "\Accounts\" & Trim$(Name) & ".ini"
        RightPassword = GetVar(FileName, "GENERAL", "Password")
        
        If Trim$(Password) = Trim$(RightPassword) Then
            PasswordOK = True
        End If
    End If
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modDatabase.bas", "PasswordOK", Err.Number, Err.Description)
End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
'On Error GoTo errorhandler:
Dim i As Long

    Player(Index).Login = Name
    Player(Index).Password = Password
    
    For i = 1 To MAX_CHARS
        Call ClearChar(Index, i)
    Next i
    
    Call SavePlayer(Index)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "AddAccount", Err.Number, Err.Description)
End Sub

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
'On Error GoTo errorhandler:
Dim f As Long
Dim START_MAP As Long, START_X As Long, START_Y As Long

        START_MAP = Class(ClassNum).Map
        START_X = Class(ClassNum).x
        START_Y = Class(ClassNum).y
        
    If Trim$(Player(Index).Char(CharNum).Name) = "" Then
        
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
        Player(Index).Char(CharNum).SPEED = Class(ClassNum).SPEED
        Player(Index).Char(CharNum).MAGI = Class(ClassNum).MAGI
        
        Player(Index).Char(CharNum).Map = START_MAP
        Player(Index).Char(CharNum).x = START_X
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
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "AddChar", Err.Number, Err.Description)
End Sub

Sub DelChar(ByVal Index As Long, ByVal CharNum As Long)
'On Error GoTo errorhandler:
Dim f1 As Long, f2 As Long
Dim s As String

    Call DeleteName(Player(Index).Char(CharNum).Name)
    Call ClearChar(Index, CharNum)
    Call SavePlayer(Index)
    'Debug.Print "DelChar"
    'Debug.Print Index
    'Debug.Print CharNum
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "DelChar", Err.Number, Err.Description)
End Sub

Function FindChar(ByVal Name As String) As Boolean
'On Error GoTo errorhandler:
Dim f As Long
Dim s As String

    FindChar = False
    
    f = FreeFile
    Open App.Path & "\Accounts\charlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            
            If Trim$(LCase(s)) = Trim$(LCase(Name)) Then
                FindChar = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
ErrorHandlerExit:
  Exit Function
errorhandler:
  Call ReportError("modDatabase.bas", "FindChar", Err.Number, Err.Description)
End Function

Sub SaveAllPlayersOnline()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveAllPlayersOnline", Err.Number, Err.Description)
End Sub

Sub LoadClasses()
'On Error GoTo errorhandler:
Dim FileName As String
Dim i As Long

    Call CheckClasses
    
    FileName = App.Path & "\data\classes.ini"
    
    Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
    Max_Visible_Classes = Val(GetVar(FileName, "INIT", "MaxVisibleClasses"))
    
    ReDim Class(0 To Max_Classes) As ClassRec
    
    Call ClearClasses
    
    For i = 0 To Max_Classes
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).Sprite = GetVar(FileName, "CLASS" & i, "Sprite")
        Class(i).HP = Val(GetVar(FileName, "CLASS" & i, "HP"))
        Class(i).MP = Val(GetVar(FileName, "CLASS" & i, "MP"))
        Class(i).SP = Val(GetVar(FileName, "CLASS" & i, "SP"))
        Class(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class(i).SPEED = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class(i).MAGI = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
        Class(i).Map = Val(GetVar(FileName, "CLASS" & i, "MAP"))
        Class(i).x = Val(GetVar(FileName, "CLASS" & i, "X"))
        Class(i).y = Val(GetVar(FileName, "CLASS" & i, "Y"))
        frmLoad.lblClasses.Caption = i
        DoEvents
    Next i
    frmLoad.lblClasses.ForeColor = &H8000&
    frmLoad.lblClasses.Caption = "DONE"
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "LoadClasses", Err.Number, Err.Description)
End Sub

Sub SaveClasses()
'On Error GoTo errorhandler:
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\data\classes.ini"
    
    For i = 0 To Max_Classes
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "Sprite", STR(Class(i).Sprite))
        Call PutVar(FileName, "CLASS" & i, "HP", STR(Class(i).HP))
        Call PutVar(FileName, "CLASS" & i, "MP", STR(Class(i).MP))
        Call PutVar(FileName, "CLASS" & i, "SP", STR(Class(i).SP))
        Call PutVar(FileName, "CLASS" & i, "STR", STR(Class(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR(Class(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR(Class(i).SPEED))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR(Class(i).MAGI))
        Call PutVar(FileName, "CLASS" & i, "MAP", STR(Class(i).Map))
        Call PutVar(FileName, "CLASS" & i, "X", STR(Class(i).x))
        Call PutVar(FileName, "CLASS" & i, "Y", STR(Class(i).y))
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveClasses", Err.Number, Err.Description)
End Sub

Sub SaveClass(ByVal ClassNum As Byte, Optional NewNumber As Byte = 0)
'On Error GoTo errorhandler:
Dim FileName As String
Dim i As Byte

    FileName = App.Path & "\data\classes.ini"
    
    i = ClassNum
    If NewNumber = 0 Then
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "Sprite", STR(Class(i).Sprite))
        Call PutVar(FileName, "CLASS" & i, "HP", STR(Class(i).HP))
        Call PutVar(FileName, "CLASS" & i, "MP", STR(Class(i).MP))
        Call PutVar(FileName, "CLASS" & i, "SP", STR(Class(i).SP))
        Call PutVar(FileName, "CLASS" & i, "STR", STR(Class(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR(Class(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR(Class(i).SPEED))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR(Class(i).MAGI))
        Call PutVar(FileName, "CLASS" & i, "MAP", STR(Class(i).Map))
        Call PutVar(FileName, "CLASS" & i, "X", STR(Class(i).x))
        Call PutVar(FileName, "CLASS" & i, "Y", STR(Class(i).y))
        Exit Sub
    Else
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class(NewNumber).Name))
        Call PutVar(FileName, "CLASS" & i, "Sprite", STR(Class(NewNumber).Sprite))
        Call PutVar(FileName, "CLASS" & i, "HP", STR(Class(NewNumber).HP))
        Call PutVar(FileName, "CLASS" & i, "MP", STR(Class(NewNumber).MP))
        Call PutVar(FileName, "CLASS" & i, "SP", STR(Class(NewNumber).SP))
        Call PutVar(FileName, "CLASS" & i, "STR", STR(Class(NewNumber).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR(Class(NewNumber).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR(Class(NewNumber).SPEED))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR(Class(NewNumber).MAGI))
        Call PutVar(FileName, "CLASS" & i, "MAP", STR(Class(NewNumber).Map))
        Call PutVar(FileName, "CLASS" & i, "X", STR(Class(NewNumber).x))
        Call PutVar(FileName, "CLASS" & i, "Y", STR(Class(NewNumber).y))
        Exit Sub
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveClass", Err.Number, Err.Description)
End Sub

Sub PrintClassINI()
'On Error GoTo errorhandler:
Dim f As Integer
Dim n As Byte
f = FreeFile
        
    ' Create a new instance and populate it
    Open App.Path & "\Data\classes.ini" For Output As #f
        Print #f, "; Distribute 20 points"
        Print #f, "; Make sure that STR DEF SPEED are all at least 1, MAGI can be 0"
        Print #f, "; Keep in mind that the number of Max Classes and Max Visible classes starts with 0 and goes to the specified"
        Print #f, "; The Map, X, and Y choices determine where the class starting location will be"
        Print #f, ""
        Print #f, "[INIT]"
        Print #f, "MaxClasses = " & CStr(Max_Classes)
        Print #f, "MaxVisibleClasses = " & CStr(Max_Visible_Classes)
    Close #f
            
    ' Loop through all classes and write data
    For n = 0 To Max_Classes
        Debug.Print "Class" & n & " " & Class(n).Name
        ' Add that nice spacing we love so much
        'Print #f, ""
        Call PrintClass(n, True)
    Next n
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "PrintClassINI", Err.Number, Err.Description)
End Sub

Sub PrintClass(ByVal ClassNum As Byte, Optional AddSpacing As Boolean = False)
'On Error GoTo errorhandler:
Dim FileName As String
Dim i As Byte
Dim f As Integer

    FileName = App.Path & "\data\classes.ini"
    f = FreeFile
    
    i = ClassNum
    Open FileName For Append As #f
        If AddSpacing = True Then
            Print #f, ""
        End If
        Print #f, "[CLASS" & i & "]"
        Print #f, "Name=" & Trim$(Class(i).Name)
        Print #f, "Sprite=" & STR(Class(i).Sprite)
        Print #f, "HP=" & STR(Class(i).HP)
        Print #f, "MP=" & STR(Class(i).MP)
        Print #f, "SP=" & STR(Class(i).SP)
        Print #f, "STR=" & STR(Class(i).STR)
        Print #f, "DEF=" & STR(Class(i).DEF)
        Print #f, "SPEED=" & STR(Class(i).SPEED)
        Print #f, "MAGI=" & STR(Class(i).MAGI)
        Print #f, "MAP=" & STR(Class(i).Map)
        Print #f, "X=" & STR(Class(i).x)
        Print #f, "Y=" & STR(Class(i).y)
    Close #f
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "PrintClass", Err.Number, Err.Description)
End Sub

Sub CheckClasses()
'On Error GoTo errorhandler:
    If Not FileExist("data\classes.ini") Then
        Call SaveClasses
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "CheckClasses", Err.Number, Err.Description)
End Sub

Sub SaveItems()
'On Error GoTo errorhandler:
Dim i As Long
    
    For i = 1 To MAX_ITEMS
        frmLoad.lblStatus.Caption = "Saving Item " & i
        Call SaveItem(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveItems", Err.Number, Err.Description)
End Sub

Sub SaveItem(ByVal ItemNum As Long)
'On Error GoTo errorhandler:
Dim FileName As String
Dim f As Integer

    FileName = App.Path & "\items\item" & ItemNum & ".dat"
    f = FreeFile
    
    Open FileName For Binary As #f
        frmLoad.lblStatus.Caption = "Creating Item " & ItemNum
        Put #f, , Item(ItemNum)
    Close #f
    DoEvents
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveItem", Err.Number, Err.Description)
End Sub

Sub LoadItems()
'On Error GoTo errorhandler:
Dim FileName As String
Dim f As Integer
Dim i As Long

    Call CheckItems
    
    For i = 1 To MAX_ITEMS
    FileName = App.Path & "\items\item" & i & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Item(i)
        frmLoad.lblItems.Caption = i
    Close #f
        DoEvents
    Next i
    frmLoad.lblItems.ForeColor = &H8000&
    frmLoad.lblItems.Caption = "DONE"
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "LoadItems", Err.Number, Err.Description)
End Sub

Sub CheckItems()
'On Error GoTo errorhandler:
Dim File As String
Dim x As Long

For x = 1 To MAX_ITEMS
File = "items\item" & x & ".dat"
    If Not FileExist(File) Then
        Call SaveItems
    End If
Next x
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "CheckItems", Err.Number, Err.Description)
End Sub

Sub SaveShops()
'On Error GoTo errorhandler:
Dim i As Long
    For i = 1 To MAX_SHOPS
        frmLoad.lblStatus.Caption = "Saving Shop " & i
        Call SaveShop(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveShops", Err.Number, Err.Description)
End Sub

Sub SaveShop(ByVal ShopNum As Long)
'On Error GoTo errorhandler:
Dim FileName As String
Dim f As Integer
Dim i As Long

    FileName = App.Path & "\shops\shop" & ShopNum & ".dat"
    f = FreeFile
    
    Open FileName For Binary As #f
        frmLoad.lblStatus.Caption = "Creating Shop " & ShopNum
        Put #f, , Shop(ShopNum)
    Close #f
    DoEvents
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveShop", Err.Number, Err.Description)
End Sub

Sub LoadShops()
'On Error GoTo errorhandler:
On Error Resume Next

Dim FileName As String
Dim f As Integer
Dim x As Long, y As Long

    Call CheckShops

For y = 1 To MAX_SHOPS
    FileName = App.Path & "\shops\shop" & y & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Shop(y)
        frmLoad.lblShops.Caption = y
    Close #f
    DoEvents
Next y
    frmLoad.lblShops.ForeColor = &H8000&
    frmLoad.lblShops.Caption = "DONE"
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "LoadShops", Err.Number, Err.Description)
End Sub

Sub CheckShops()
'On Error GoTo errorhandler:
Dim File As String
Dim x As Long

For x = 1 To MAX_SHOPS
File = "shops\shop" & x & ".dat"
    If Not FileExist(File) Then
        Call SaveShops
    End If
Next x
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "CheckShops", Err.Number, Err.Description)
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
'On Error GoTo errorhandler:
Dim FileName As String
Dim f As Integer
Dim i As Long

    FileName = App.Path & "\spells\spell" & SpellNum & ".dat"
    f = FreeFile
    
    Open FileName For Binary As #f
        frmLoad.lblStatus.Caption = "Creating Spell " & SpellNum
        Put #f, , Spell(SpellNum)
    Close #f
    DoEvents
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveSpell", Err.Number, Err.Description)
End Sub

Sub SaveSpells()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_SPELLS
        frmLoad.lblStatus.Caption = "Saving Spell " & i
        Call SaveSpell(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveSpells", Err.Number, Err.Description)
End Sub

Sub LoadSpells()
'On Error GoTo errorhandler:
Dim FileName As String
Dim f As Integer
Dim i As Long

    Call CheckSpells
    
    For i = 1 To MAX_SPELLS
        FileName = App.Path & "\spells\spell" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Spell(i)
            frmLoad.lblSpells.Caption = i
        Close #f
        DoEvents
    Next i
    frmLoad.lblSpells.ForeColor = &H8000&
    frmLoad.lblSpells.Caption = "DONE"
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "LoadSpells", Err.Number, Err.Description)
End Sub

Sub CheckSpells()
'On Error GoTo errorhandler:
Dim File As String
Dim x As Long

For x = 1 To MAX_SPELLS
File = "spells\spell" & x & ".dat"
    If Not FileExist(File) Then
        Call SaveSpells
    End If
Next x
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "CheckSpells", Err.Number, Err.Description)
End Sub

Sub SaveNpcs()
'On Error GoTo errorhandler:
Dim i As Long
    
    For i = 1 To MAX_NPCS
        frmLoad.lblStatus.Caption = "Saving NPC " & i
        Call SaveNpc(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveNpcs", Err.Number, Err.Description)
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
'On Error GoTo errorhandler:
Dim FileName As String
Dim f As Integer

    FileName = App.Path & "\npcs\npc" & NpcNum & ".dat"
    f = FreeFile
    
    Open FileName For Binary As #f
        frmLoad.lblStatus.Caption = "Creating Npc " & NpcNum
        Put #f, , Npc(NpcNum)
    Close #f
    DoEvents
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveNpc", Err.Number, Err.Description)
End Sub

Sub LoadNpcs()
'On Error GoTo errorhandler:
On Error Resume Next

Dim FileName As String
Dim f As Integer
Dim i As Long

    Call CheckNpcs
    
    For i = 1 To MAX_NPCS
    FileName = App.Path & "\npcs\npc" & i & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Npc(i)
        frmLoad.lblNpcs.Caption = i
    Close #f
        DoEvents
    Next i
    frmLoad.lblNpcs.ForeColor = &H8000&
    frmLoad.lblNpcs.Caption = "DONE"
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "LoadNpcs", Err.Number, Err.Description)
End Sub

Sub CheckNpcs()
'On Error GoTo errorhandler:
Dim File As String
Dim x As Long

For x = 1 To MAX_NPCS
File = "npcs\npc" & x & ".dat"
    If Not FileExist(File) Then
        Call SaveNpcs
    End If
Next x
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "CheckNpcs", Err.Number, Err.Description)
End Sub

Sub LoadExps()
'On Error GoTo errorhandler:
Dim FileName As String
Dim i As Long

    Call CheckExps
    
    FileName = App.Path & "\Data\experience_curve.ini"
    
    For i = 1 To MAX_EXPERIENCE
        Experience(i) = CLng(GetVar(FileName, "EXPERIENCE", "Exp" & i))
        Call SetStatus("Loading EXP")
        frmLoad.lblEXP.Caption = i
        DoEvents
    Next i
    frmLoad.lblEXP.ForeColor = &H8000&
    frmLoad.lblEXP.Caption = "DONE"
    DoEvents
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "LoadExps", Err.Number, Err.Description)
End Sub

Sub CheckExps()
'On Error GoTo errorhandler:
    If Not FileExist("Data\experience_curve.ini") Then
        Dim i As Long
    
        For i = 1 To MAX_EXPERIENCE
            Call PutVar(App.Path & "\Data\experience_curve.ini", "EXPERIENCE", "Exp" & i, i * 1500)
            Call SetStatus("Creating exp file! Writing exp #" & i)
            DoEvents
        Next i
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "CheckExps", Err.Number, Err.Description)
End Sub

Sub ClearExps()
'On Error GoTo errorhandler:
Dim i As Long

    For i = 1 To MAX_EXPERIENCE
        Experience(i) = 0
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "ClearExps", Err.Number, Err.Description)
End Sub

Sub SaveMap(ByVal MapNum As Long)
'On Error GoTo errorhandler:
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        frmLoad.lblStatus.Caption = "Creating Map " & MapNum
        Put #f, , Map(MapNum)
    Close #f
        DoEvents
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveMap", Err.Number, Err.Description)
End Sub

Sub SaveMaps()
'On Error GoTo errorhandler:
Dim FileName As String
Dim i As Long
Dim f As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "SaveMaps", Err.Number, Err.Description)
End Sub

Sub LoadMaps()
'On Error GoTo errorhandler:
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        FileName = App.Path & "\maps\map" & i & ".dat"
        
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Map(i)
            frmLoad.lblMaps.Caption = i
        Close #f
    
        DoEvents
    Next i
    frmLoad.lblMaps.ForeColor = &H8000&
    frmLoad.lblMaps.Caption = "DONE"
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "LoadMaps", Err.Number, Err.Description)
End Sub

Sub CheckMaps()
'On Error GoTo errorhandler:
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
            Call SaveMap(i)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "CheckMaps", Err.Number, Err.Description)
End Sub

Sub AddLog(ByVal Text As String, ByVal FN As String)
'On Error GoTo errorhandler:
Dim FileName As String
Dim f As Long

    If ServerLog = True Then
        FileName = App.Path & "\logs\" & FN
    
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
errorhandler:
  Call ReportError("modDatabase.bas", "AddLog", Err.Number, Err.Description)
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
'On Error GoTo errorhandler:
Dim FileName, IP As String
Dim f As Long, i As Long

    FileName = App.Path & "\data\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
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
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "BanIndex", Err.Number, Err.Description)
End Sub

Sub DeleteName(ByVal Name As String)
'On Error GoTo errorhandler:
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
        If Trim$(LCase(s)) <> Trim$(LCase(Name)) Then
            Print #f2, s
        End If
    Loop
    
    Close #f1
    Close #f2
    
    Call Kill(App.Path & "\accounts\chartemp.txt")
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "DeleteName", Err.Number, Err.Description)
End Sub

Public Sub LoadLibrary()
    Dim sFilename As String
    Dim i As Long
    Dim FileName As String
    
    '  Initialize the Dir$ function (and get the first filename if it exists).
    sFilename = Dir$(App.Path & "\Library\*.txt")
    
    '  Loop on the text files in the directory.
    Do While sFilename <> ""
        '  Add the current filename to the listbox
        frmLibrary.lstLibrary.AddItem sFilename
        '  Advance to the next filename in the directory.
        sFilename = Dir$
    Loop
    
    For i = 0 To frmLibrary.lstLibrary.ListCount - 1
        FileName = App.Path & "\Library\" & frmLibrary.lstLibrary.List(i)
        If GetVar(FileName, "DATA", "Enabled") = "True" Then
            frmLibrary.lstLibrary.Selected(i) = True
        End If
    Next i
    
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "LoadLibrary", Err.Number, Err.Description)
End Sub

Public Sub LoadScripts()
Dim i As Long
Dim FileName As String
Dim Msg As String

    Msg = vbNullString
    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands
    For i = 0 To frmLibrary.lstLibrary.ListCount - 1
        FileName = App.Path & "\Library\" & frmLibrary.lstLibrary.List(i)
        If GetVar(FileName, "DATA", "Enabled") = "True" Then
            MyScript.ReadInCode App.Path & "\Library\" & frmLibrary.lstLibrary.List(i), frmLibrary.lstLibrary.List(i), MyScript.SControl, False
            If Msg = "" Then
                Msg = frmLibrary.lstLibrary.List(i)
            Else
                Msg = Msg + ", " + frmLibrary.lstLibrary.List(i)
            End If
        End If
    Next i
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    
    Call TextAdd(frmServer.txtText, "Loaded Scripts: " & Msg & ".", True)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modDatabase.bas", "LoadScripts", Err.Number, Err.Description)
End Sub
