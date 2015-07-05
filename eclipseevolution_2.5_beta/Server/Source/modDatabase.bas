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

' ---------------------------------------------------------------------------------------
' Procedure : GetVar
' Purpose   :  Reads a variable from an INI file
' ---------------------------------------------------------------------------------------
Function GetVar(file As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    
    On Error GoTo GetVar_Error
    
    szReturn = ""
    
    sSpaces = Space(5000)
    
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), file)
    
    GetVar = RTrim$(sSpaces)
    GetVar = left$(GetVar, Len(GetVar) - 1)
    
    On Error GoTo 0
    Exit Function
    
GetVar_Error:
    
    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure GetVar of Module modDatabase"
End Function

' ---------------------------------------------------------------------------------------
' Procedure : PutVar
' Purpose   : Writes a file to an INI file
' ---------------------------------------------------------------------------------------
Sub PutVar(file As String, Header As String, Var As String, Value As String)
    On Error GoTo PutVar_Error
    
    Call WritePrivateProfileString(Header, Var, Value, file)
    
    On Error GoTo 0
    Exit Sub
    
PutVar_Error:
    
    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure PutVar of Module modDatabase"
End Sub

Sub LoadExps()
    On Error GoTo ExpErr
    Dim FileName As String
    Dim i As Integer
    
    Call CheckExps
    
    FileName = App.Path & "\experience.ini"
    
    For i = 1 To MAX_LEVEL
        temp = i / MAX_LEVEL * 100
        Call SetStatus("Loading exp... " & temp & "%")
        Experience(i) = GetVar(FileName, "EXPERIENCE", "Exp" & i)
        
        DoEvents
    Next i
    Exit Sub
    
ExpErr:
    Call MsgBox("Error loading EXP for level " & i & ". Make sure experience.ini has the correct variables! ERR: " & Err.number & ", Desc: " & Err.Description, vbCritical)
    Call DestroyServer
    End
End Sub

Sub CheckExps()
    If Not FileExist("experience.ini") Then
        Dim i As Integer
        
        For i = 1 To MAX_LEVEL
            temp = i / MAX_LEVEL * 100
            Call SetStatus("Saving exp... " & temp & "%")
            Call PutVar(App.Path & "\experience.ini", "EXPERIENCE", "Exp" & i, i * 1500)
        Next i
    End If
End Sub

Sub ClearExps()
    Dim i As Integer
    
    For i = 1 To MAX_LEVEL
        Experience(i) = 0
    Next i
End Sub

Sub LoadEmos()
    Dim FileName As String
    Dim i As Integer
    
    Call CheckEmos
    
    FileName = App.Path & "\emoticons.ini"
    
    For i = 0 To MAX_EMOTICONS
        temp = i / MAX_EMOTICONS * 100
        Call SetStatus("Loading emoticons... " & temp & "%")
        Emoticons(i).Pic = GetVar(FileName, "EMOTICONS", "Emoticon" & i)
        Emoticons(i).Command = GetVar(FileName, "EMOTICONS", "EmoticonC" & i)
    Next i
End Sub
Sub LoadElements()
    On Error GoTo ElementErr
    Dim FileName As String
    Dim i As Integer
    
    Call CheckElements
    
    FileName = App.Path & "\elements.ini"
    
    For i = 0 To MAX_ELEMENTS
        temp = i / MAX_ELEMENTS * 100
        Call SetStatus("Loading elements... " & temp & "%")
        Element(i).Name = GetVar(FileName, "ELEMENTS", "ElementName" & i)
        Element(i).Strong = Val(GetVar(FileName, "ELEMENTS", "ElementStrong" & i))
        Element(i).Weak = Val(GetVar(FileName, "ELEMENTS", "ElementWeak" & i))
    Next i
    Exit Sub
    
ElementErr:
    Call MsgBox("Error loading element " & i & ". Make sure all the variables in elements.ini are correct!", vbCritical)
    Call DestroyServer
    End
End Sub

Sub CheckEmos()
    If Not FileExist("emoticons.ini") Then
        Dim i As Integer
        
        For i = 0 To MAX_EMOTICONS
            temp = i / MAX_LEVEL * 100
            Call SetStatus("Saving emoticons... " & temp & "%")
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "Emoticon" & i, 0)
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "EmoticonC" & i, "")
        Next i
    End If
End Sub
Sub CheckElements()
    If Not FileExist("elements.ini") Then
        Dim i As Integer
        
        For i = 0 To MAX_ELEMENTS
            temp = i / MAX_ELEMENTS * 100
            Call SetStatus("Saving elements... " & temp & "%")
            DoEvents
            Call PutVar(App.Path & "\elements.ini", "ELEMENTS", "ElementName" & i, "")
            Call PutVar(App.Path & "\elements.ini", "ELEMENTS", "ElementStrong" & i, 0)
            Call PutVar(App.Path & "\elements.ini", "ELEMENTS", "ElementWeak" & i, 0)
        Next i
    End If
End Sub

Sub ClearEmos()
    Dim i As Integer
    
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
Sub SaveElement(ByVal ElementNum As Long)
    Dim FileName As String
    
    FileName = App.Path & "\elements.ini"
    
    Call PutVar(FileName, "ELEMENTS", "ElementName" & ElementNum, Trim(Element(ElementNum).Name))
    Call PutVar(FileName, "ELEMENTS", "ElementStrong" & ElementNum, Val(Element(ElementNum).Strong))
    Call PutVar(FileName, "ELEMENTS", "ElementWeak" & ElementNum, Val(Element(ElementNum).Weak))
End Sub
Function FileExist(FileName As String) As Boolean
    On Error GoTo ErrorHandler
'    get the attributes and ensure that it isn't a directory
    FileExist = (GetAttr(App.Path & "\" & FileName) And vbDirectory) = 0
ErrorHandler:
'    if an error occurs, this function returns False
End Function
Sub SavePlayer(ByVal index As Long)
    Dim FileName As String
    Dim f As Long 'File
    Dim i As Integer
    On Error Resume Next
'    Save login information first
    FileName = App.Path & "\accounts\" & Trim$(player(index).Login) & "_info.ini"
    
    Call PutVar(FileName, "ACCESS", "Login", Trim$(player(index).Login))
    Call PutVar(FileName, "ACCESS", "Password", Trim$(player(index).Password))
    Call PutVar(FileName, "ACCESS", "Email", Trim$(player(index).Email))
    
'    Make the directory
    If LCase(Dir(App.Path & "\accounts\" & Trim$(player(index).Login), vbDirectory)) <> LCase(Trim$(player(index).Login)) Then
        Call MkDir(App.Path & "\accounts\" & Trim$(player(index).Login))
    End If
    
'    Now save their characters
    For i = 1 To MAX_CHARS
        FileName = App.Path & "\accounts\" & Trim$(player(index).Login) & "\char" & i & ".dat"
        
'        Save the character
        f = FreeFile
        Open FileName For Binary As #f
        Put #f, , player(index).Char(i)
        Close #f
        
    Next i
End Sub

Public Sub Findcharfile(ByVal RName As String)
    Dim fso, fldr, folders, fldrnm, s
    Dim i As String
    Dim f As Long
    Dim c As Long
    Dim charpath As String
    Dim FileName As String
    Dim g As Long
    Dim FName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set fldr = fso.GetFolder(App.Path & "\accounts")
    
    Set folders = fldr.subfolders
    
    Call CleartempPlayer
    
    For Each fldrnm In folders
        charpath = App.Path & "\accounts\" & fldrnm.Name
'        MsgBox (fldrnm.Name & vbCrLf & charpath & vbCrLf & vbCrLf & "Searching for: " & RName)
        For c = 1 To MAX_CHARS
            f = FreeFile
            FileName = charpath & "\char" & c & ".dat"
            Open FileName For Binary As #f
            Get #f, , tempplayer.Char(c)
'            MsgBox ("Looking for: " & RName & vbCrLf & vbCrLf & FileName & vbCrLf & vbCrLf & "CharName: " & tempplayer.Char(c).Name)
            FName = Trim$(tempplayer.Char(c).Name)
            RName = Trim$(RName)
            If LCase(FName) = LCase(RName) Then
                tempplayer.Login = fldrnm.Name
                Close #f
                Exit Sub
            Else
                Close #f
                Call CleartempPlayer
            End If
        Next c
    Next
End Sub


Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
    On Error GoTo PlayerErr
    Dim f As Long 'File
    Dim i As Integer
    Dim FileName As String
    
    Call ClearPlayer(index)
    
    FileName = App.Path & "\accounts\" & Trim(Name) & ".ini"
    
    If FileExist("\accounts\" & Trim(Name) & ".ini") Then
        Call LoadPlayerFromINI(index, Name)
'        Delete the old file
        Kill FileName
    Else
'        Load the account settings
        FileName = App.Path & "\accounts\" & Trim$(Name) & "_info.ini"
        
        player(index).Login = Name
        player(index).Password = GetVar(FileName, "ACCESS", "Password")
        player(index).Email = GetVar(FileName, "ACCESS", "Email")
        
'        Load the .dat
        For i = 1 To MAX_CHARS
            
            FileName = App.Path & "\accounts\" & Trim$(player(index).Login) & "\char" & i & ".dat"
            
            f = FreeFile
            Open FileName For Binary As #f
            Get #f, , player(index).Char(i)
            Close #f
            
        Next i
    End If
    
    Exit Sub
    
PlayerErr:
    Call MsgBox("Error loading player " & index & ". Make sure all variables are correct!", vbCritical)
    Call DestroyServer
    End
End Sub

' Loads a player from an INI file as opposed to a .dat
Sub LoadPlayerFromINI(ByVal index As Integer, ByVal Name As String)
    
    On Error GoTo PlayerErr
    Dim FileName As String
    Dim i As Long
    Dim j As Long
    Dim n As Long
    
    FileName = App.Path & "\accounts\" & Name & ".ini"
    
    player(index).Login = GetVar(FileName, "GENERAL", "Login")
    player(index).Password = GetVar(FileName, "GENERAL", "Password")
    
    For i = 1 To MAX_CHARS
        
'        General
        player(index).Char(i).Name = GetVar(FileName, "CHAR" & i, "Name")
        player(index).Char(i).Sex = Val(GetVar(FileName, "CHAR" & i, "Sex"))
        player(index).Char(i).Class = Val(GetVar(FileName, "CHAR" & i, "Class"))
        player(index).Char(i).Sprite = Val(GetVar(FileName, "CHAR" & i, "Sprite"))
        player(index).Char(i).Level = Val(GetVar(FileName, "CHAR" & i, "Level"))
        player(index).Char(i).Exp = Val(GetVar(FileName, "CHAR" & i, "Exp"))
        player(index).Char(i).access = Val(GetVar(FileName, "CHAR" & i, "Access"))
        player(index).Char(i).PK = Val(GetVar(FileName, "CHAR" & i, "PK"))
        player(index).Char(i).Guild = GetVar(FileName, "CHAR" & i, "Guild")
        player(index).Char(i).Guildaccess = Val(GetVar(FileName, "CHAR" & i, "Guildaccess"))
        player(index).Char(i).head = Val(GetVar(FileName, "CHAR" & i, "Head"))
        player(index).Char(i).body = Val(GetVar(FileName, "CHAR" & i, "Body"))
        player(index).Char(i).leg = Val(GetVar(FileName, "CHAR" & i, "Leg"))
        
'        Vitals
        player(index).Char(i).HP = Val(GetVar(FileName, "CHAR" & i, "HP"))
        player(index).Char(i).MP = Val(GetVar(FileName, "CHAR" & i, "MP"))
        player(index).Char(i).SP = Val(GetVar(FileName, "CHAR" & i, "SP"))
        
'        Stats
        player(index).Char(i).STR = Val(GetVar(FileName, "CHAR" & i, "STR"))
        player(index).Char(i).DEF = Val(GetVar(FileName, "CHAR" & i, "DEF"))
        player(index).Char(i).Speed = Val(GetVar(FileName, "CHAR" & i, "SPEED"))
        player(index).Char(i).Magi = Val(GetVar(FileName, "CHAR" & i, "MAGI"))
        player(index).Char(i).POINTS = Val(GetVar(FileName, "CHAR" & i, "POINTS"))
        
'        Worn equipment
        player(index).Char(i).ArmorSlot = Val(GetVar(FileName, "CHAR" & i, "ArmorSlot"))
        player(index).Char(i).WeaponSlot = Val(GetVar(FileName, "CHAR" & i, "WeaponSlot"))
        player(index).Char(i).HelmetSlot = Val(GetVar(FileName, "CHAR" & i, "HelmetSlot"))
        player(index).Char(i).ShieldSlot = Val(GetVar(FileName, "CHAR" & i, "ShieldSlot"))
        player(index).Char(i).LegsSlot = Val(GetVar(FileName, "CHAR" & i, "LegsSlot"))
        player(index).Char(i).RingSlot = Val(GetVar(FileName, "CHAR" & i, "RingSlot"))
        player(index).Char(i).NecklaceSlot = Val(GetVar(FileName, "CHAR" & i, "NecklaceSlot"))
        
'        sprite stuff
        player(index).Char(i).head = Val(GetVar(FileName, "CHAR" & i, "Head"))
        player(index).Char(i).body = Val(GetVar(FileName, "CHAR" & i, "Body"))
        player(index).Char(i).leg = Val(GetVar(FileName, "CHAR" & i, "Leg"))
        
'        Paperdoll
        If GetVar(FileName, "CHAR" & i, "Paperdoll") = "" Then
            Call PutVar(FileName, "CHAR" & i, "Paperdoll", 1)
            player(index).Char(i).Paperdoll = 1
        Else
            player(index).Char(i).Paperdoll = Val(GetVar(FileName, "CHAR" & i, "Paperdoll"))
        End If
        
'        skill
        For j = 1 To MAX_SKILLS
            player(index).Char(i).SkillLvl(j) = Val(GetVar(FileName, "CHAR" & i, "Skill" & j & "lvl"))
            player(index).Char(i).SkillExp(j) = Val(GetVar(FileName, "CHAR" & i, "Skill" & j & "exp"))
        Next j
        
'        Position
        player(index).Char(i).Map = Val(GetVar(FileName, "CHAR" & i, "Map"))
        player(index).Char(i).x = Val(GetVar(FileName, "CHAR" & i, "X"))
        player(index).Char(i).y = Val(GetVar(FileName, "CHAR" & i, "Y"))
        player(index).Char(i).Dir = Val(GetVar(FileName, "CHAR" & i, "Dir"))
        
'        Check to make sure that they aren't on map 0, if so reset'm
        If player(index).Char(i).Map = 0 Then
            player(index).Char(i).Map = START_MAP
            player(index).Char(i).x = START_X
            player(index).Char(i).y = START_Y
        End If
        
'        Inventory
        For n = 1 To MAX_INV
            player(index).Char(i).Inv(n).num = Val(GetVar(FileName, "CHAR" & i, "InvItemNum" & n))
            player(index).Char(i).Inv(n).Value = Val(GetVar(FileName, "CHAR" & i, "InvItemVal" & n))
            player(index).Char(i).Inv(n).Dur = Val(GetVar(FileName, "CHAR" & i, "InvItemDur" & n))
        Next n
        
'        Spells
        For n = 1 To MAX_PLAYER_SPELLS
            player(index).Char(i).Spell(n) = Val(GetVar(FileName, "CHAR" & i, "Spell" & n))
        Next n
        
        FileName = App.Path & "\banks\" & Name & ".ini"
'        Bank
        For n = 1 To MAX_BANK
            player(index).Char(i).Bank(n).num = Val(GetVar(FileName, "CHAR" & i, "BankItemNum" & n))
            player(index).Char(i).Bank(n).Value = Val(GetVar(FileName, "CHAR" & i, "BankItemVal" & n))
            player(index).Char(i).Bank(n).Dur = Val(GetVar(FileName, "CHAR" & i, "BankItemDur" & n))
        Next n
    Next i
    Exit Sub
    
PlayerErr:
    Call MsgBox("Error loading player " & i & ". Make sure all variables are correct!", vbCritical)
    Call DestroyServer
    End
End Sub

' Loads a tempplayer from an INI file as opposed to a .dat
Sub LoadTempPlayerFromINI(ByVal Name As String)
    
    On Error GoTo PlayerErr
    Dim fso, fldr, files, file, s
    Dim FileName As String
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim check As Long
    
    Call CleartempPlayer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    check = 0
    
    Set fldr = fso.GetFolder(App.Path & "\accounts")
    
    Set files = fldr.files
    
    For Each file In files
        If check = 0 Then
            i = 1
            FileName = App.Path & "\accounts\" & file.Name
            If GetVar(FileName, "CHAR" & i, "Name") = Name Then
                check = 1
                tempplayer.Login = GetVar(FileName, "GENERAL", "Login")
                tempplayer.Password = GetVar(FileName, "GENERAL", "Password")
                
                
                
'                General
                tempplayer.Char(i).Name = GetVar(FileName, "CHAR" & i, "Name")
                tempplayer.Char(i).Sex = Val(GetVar(FileName, "CHAR" & i, "Sex"))
                tempplayer.Char(i).Class = Val(GetVar(FileName, "CHAR" & i, "Class"))
                tempplayer.Char(i).Sprite = Val(GetVar(FileName, "CHAR" & i, "Sprite"))
                tempplayer.Char(i).Level = Val(GetVar(FileName, "CHAR" & i, "Level"))
                tempplayer.Char(i).Exp = Val(GetVar(FileName, "CHAR" & i, "Exp"))
                tempplayer.Char(i).access = Val(GetVar(FileName, "CHAR" & i, "Access"))
                tempplayer.Char(i).PK = Val(GetVar(FileName, "CHAR" & i, "PK"))
                tempplayer.Char(i).Guild = GetVar(FileName, "CHAR" & i, "Guild")
                tempplayer.Char(i).Guildaccess = Val(GetVar(FileName, "CHAR" & i, "Guildaccess"))
'                Vitals
                tempplayer.Char(i).MAXHP = Val(GetVar(FileName, "CHAR" & i, "HP"))
                tempplayer.Char(i).MAXMP = Val(GetVar(FileName, "CHAR" & i, "MP"))
                tempplayer.Char(i).MAXSP = Val(GetVar(FileName, "CHAR" & i, "SP"))
                
'                Stats
                tempplayer.Char(i).STR = Val(GetVar(FileName, "CHAR" & i, "STR"))
                tempplayer.Char(i).DEF = Val(GetVar(FileName, "CHAR" & i, "DEF"))
                tempplayer.Char(i).Speed = Val(GetVar(FileName, "CHAR" & i, "SPEED"))
                tempplayer.Char(i).Magi = Val(GetVar(FileName, "CHAR" & i, "MAGI"))
                tempplayer.Char(i).POINTS = Val(GetVar(FileName, "CHAR" & i, "POINTS"))
                
'                Worn equipment
                tempplayer.Char(i).ArmorSlot = Val(GetVar(FileName, "CHAR" & i, "ArmorSlot"))
                tempplayer.Char(i).WeaponSlot = Val(GetVar(FileName, "CHAR" & i, "WeaponSlot"))
                tempplayer.Char(i).HelmetSlot = Val(GetVar(FileName, "CHAR" & i, "HelmetSlot"))
                tempplayer.Char(i).ShieldSlot = Val(GetVar(FileName, "CHAR" & i, "ShieldSlot"))
                tempplayer.Char(i).LegsSlot = Val(GetVar(FileName, "CHAR" & i, "LegsSlot"))
                tempplayer.Char(i).RingSlot = Val(GetVar(FileName, "CHAR" & i, "RingSlot"))
                tempplayer.Char(i).NecklaceSlot = Val(GetVar(FileName, "CHAR" & i, "NecklaceSlot"))
                
'                skill
                
'                Position
                tempplayer.Char(i).Map = Val(GetVar(FileName, "CHAR" & i, "Map"))
                tempplayer.Char(i).x = Val(GetVar(FileName, "CHAR" & i, "X"))
                tempplayer.Char(i).y = Val(GetVar(FileName, "CHAR" & i, "Y"))
                tempplayer.Char(i).Dir = Val(GetVar(FileName, "CHAR" & i, "Dir"))
                
'                Check to make sure that they aren't on map 0, if so reset'm
                If tempplayer.Char(i).Map = 0 Then
                    tempplayer.Char(i).Map = START_MAP
                    tempplayer.Char(i).x = START_X
                    tempplayer.Char(i).y = START_Y
                End If
                
'                Inventory
                For n = 1 To MAX_INV
                    tempplayer.Char(i).Inv(n).num = Val(GetVar(FileName, "CHAR" & i, "InvItemNum" & n))
                    tempplayer.Char(i).Inv(n).Value = Val(GetVar(FileName, "CHAR" & i, "InvItemVal" & n))
                    tempplayer.Char(i).Inv(n).Dur = Val(GetVar(FileName, "CHAR" & i, "InvItemDur" & n))
                Next n
                
'                Spells
                For n = 1 To MAX_PLAYER_SPELLS
                    tempplayer.Char(i).Spell(n) = Val(GetVar(FileName, "CHAR" & i, "Spell" & n))
                Next n
                
                FileName = App.Path & "\Banks\" & file.Name & ".ini"
'                Bank
                For n = 1 To MAX_BANK
                    tempplayer.Char(i).Bank(n).num = Val(GetVar(FileName, "CHAR" & i, "BankItemNum" & n))
                    tempplayer.Char(i).Bank(n).Value = Val(GetVar(FileName, "CHAR" & i, "BankItemVal" & n))
                    tempplayer.Char(i).Bank(n).Dur = Val(GetVar(FileName, "CHAR" & i, "BankItemDur" & n))
                Next n
                Exit Sub
            End If
        End If
    Next file
    
    If check = 0 Then
        
        Set fldr = fso.GetFolder(App.Path & "\Banks")
        
        Set files = fldr.files
        
        For Each file In files
            For i = 2 To MAX_CHARS
                FileName = App.Path & "\Banks\" & file.Name
                If GetVar(FileName, "CHAR" & i, "Name") = Name Then
                    check = 1
                    tempplayer.Login = GetVar(FileName, "GENERAL", "Login")
                    tempplayer.Password = GetVar(FileName, "GENERAL", "Password")
                    
                    
                    
'                    General
                    tempplayer.Char(i).Name = GetVar(FileName, "CHAR" & i, "Name")
                    tempplayer.Char(i).Sex = Val(GetVar(FileName, "CHAR" & i, "Sex"))
                    tempplayer.Char(i).Class = Val(GetVar(FileName, "CHAR" & i, "Class"))
                    tempplayer.Char(i).Sprite = Val(GetVar(FileName, "CHAR" & i, "Sprite"))
                    tempplayer.Char(i).Level = Val(GetVar(FileName, "CHAR" & i, "Level"))
                    tempplayer.Char(i).Exp = Val(GetVar(FileName, "CHAR" & i, "Exp"))
                    tempplayer.Char(i).access = Val(GetVar(FileName, "CHAR" & i, "Access"))
                    tempplayer.Char(i).PK = Val(GetVar(FileName, "CHAR" & i, "PK"))
                    tempplayer.Char(i).Guild = GetVar(FileName, "CHAR" & i, "Guild")
                    tempplayer.Char(i).Guildaccess = Val(GetVar(FileName, "CHAR" & i, "Guildaccess"))
'                    Vitals
                    tempplayer.Char(i).MAXHP = Val(GetVar(FileName, "CHAR" & i, "HP"))
                    tempplayer.Char(i).MAXMP = Val(GetVar(FileName, "CHAR" & i, "MP"))
                    tempplayer.Char(i).MAXSP = Val(GetVar(FileName, "CHAR" & i, "SP"))
                    
'                    Stats
                    tempplayer.Char(i).STR = Val(GetVar(FileName, "CHAR" & i, "STR"))
                    tempplayer.Char(i).DEF = Val(GetVar(FileName, "CHAR" & i, "DEF"))
                    tempplayer.Char(i).Speed = Val(GetVar(FileName, "CHAR" & i, "SPEED"))
                    tempplayer.Char(i).Magi = Val(GetVar(FileName, "CHAR" & i, "MAGI"))
                    tempplayer.Char(i).POINTS = Val(GetVar(FileName, "CHAR" & i, "POINTS"))
                    
'                    Worn equipment
                    tempplayer.Char(i).ArmorSlot = Val(GetVar(FileName, "CHAR" & i, "ArmorSlot"))
                    tempplayer.Char(i).WeaponSlot = Val(GetVar(FileName, "CHAR" & i, "WeaponSlot"))
                    tempplayer.Char(i).HelmetSlot = Val(GetVar(FileName, "CHAR" & i, "HelmetSlot"))
                    tempplayer.Char(i).ShieldSlot = Val(GetVar(FileName, "CHAR" & i, "ShieldSlot"))
                    tempplayer.Char(i).LegsSlot = Val(GetVar(FileName, "CHAR" & i, "LegsSlot"))
                    tempplayer.Char(i).RingSlot = Val(GetVar(FileName, "CHAR" & i, "RingSlot"))
                    tempplayer.Char(i).NecklaceSlot = Val(GetVar(FileName, "CHAR" & i, "NecklaceSlot"))
                    
                    
'                    Position
                    tempplayer.Char(i).Map = Val(GetVar(FileName, "CHAR" & i, "Map"))
                    tempplayer.Char(i).x = Val(GetVar(FileName, "CHAR" & i, "X"))
                    tempplayer.Char(i).y = Val(GetVar(FileName, "CHAR" & i, "Y"))
                    tempplayer.Char(i).Dir = Val(GetVar(FileName, "CHAR" & i, "Dir"))
                    
'                    Check to make sure that they aren't on map 0, if so reset'm
                    If tempplayer.Char(i).Map = 0 Then
                        tempplayer.Char(i).Map = START_MAP
                        tempplayer.Char(i).x = START_X
                        tempplayer.Char(i).y = START_Y
                    End If
                    
'                    Inventory
                    For n = 1 To MAX_INV
                        tempplayer.Char(i).Inv(n).num = Val(GetVar(FileName, "CHAR" & i, "InvItemNum" & n))
                        tempplayer.Char(i).Inv(n).Value = Val(GetVar(FileName, "CHAR" & i, "InvItemVal" & n))
                        tempplayer.Char(i).Inv(n).Dur = Val(GetVar(FileName, "CHAR" & i, "InvItemDur" & n))
                    Next n
                    
'                    Spells
                    For n = 1 To MAX_PLAYER_SPELLS
                        tempplayer.Char(i).Spell(n) = Val(GetVar(FileName, "CHAR" & i, "Spell" & n))
                    Next n
                    
                    FileName = App.Path & "\Banks\" & file.Name
'                    Bank
                    For n = 1 To MAX_BANK
                        tempplayer.Char(i).Bank(n).num = Val(GetVar(FileName, "CHAR" & i, "BankItemNum" & n))
                        tempplayer.Char(i).Bank(n).Value = Val(GetVar(FileName, "CHAR" & i, "BankItemVal" & n))
                        tempplayer.Char(i).Bank(n).Dur = Val(GetVar(FileName, "CHAR" & i, "BankItemDur" & n))
                    Next n
                    Exit Sub
                End If
            Next i
        Next file
    End If
PlayerErr:
    Call MsgBox("Error loading tempplayer " & i & ". Make sure all variables are correct!", vbCritical)
    Call DestroyServer
    End
End Sub



Function AccountExist(ByVal Name As String) As Boolean
    Dim FileName As String
    
    FileName = "accounts\" & Trim(Name) & ".ini"
    
    If FileExist(FileName) Then
        AccountExist = True
        Exit Function
    Else
        AccountExist = False
    End If
    
    FileName = "accounts\" & Trim(Name) & "_info.ini"
    
    If FileExist(FileName) Then
        AccountExist = True
        Exit Function
    Else
        AccountExist = False
    End If
    
End Function

Function CharExist(ByVal index As Long, ByVal charnum As Long) As Boolean
    If Trim(player(index).Char(charnum).Name) <> "" Then
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
'        Since we're using the new character save/load we have to check both ways
        If FileExist("\accounts\" & Trim$(Name) & "_info.ini") Then
            
            FileName = App.Path & "\accounts\" & Trim$(Name) & "_info.ini"
            RightPassword = GetVar(FileName, "ACCESS", "Password")
            
            If Trim$(Password) = Trim(RightPassword) Then
                PasswordOK = True
            End If
        Else
            
            FileName = App.Path & "\accounts\" & Trim$(Name) & ".ini"
            RightPassword = GetVar(FileName, "GENERAL", "Password")
            
            If Trim$(Password) = Trim$(RightPassword) Then
                PasswordOK = True
            End If
        End If
    End If
End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String, ByVal Email As String)
    Dim i As Long
    
    player(index).Login = Name
    player(index).Password = Password
    player(index).Email = Email
    
    For i = 1 To MAX_CHARS
        Call ClearChar(index, i)
    Next i
    
    Call SavePlayer(index)
    
    If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "verified")) = 1 Then
        Call PutVar(App.Path & "\accounts\" & Trim(player(index).Login) & "_info.ini", "ACCESS", "verified", 0)
    End If
    
End Sub

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal charnum As Long, ByVal headc As Long, ByVal bodyc As Long, ByVal logc As Long)
    Dim f As Long
    
    If Trim(player(index).Char(charnum).Name) = "" Then
        player(index).charnum = charnum
        
        player(index).Char(charnum).Name = Name
        player(index).Char(charnum).Sex = Sex
        player(index).Char(charnum).Class = ClassNum
        
        If player(index).Char(charnum).Sex = SEX_MALE Then
            player(index).Char(charnum).Sprite = Class(ClassNum).MaleSprite
        Else
            player(index).Char(charnum).Sprite = Class(ClassNum).FemaleSprite
        End If
        
        player(index).Char(charnum).Level = 1
        
        player(index).Char(charnum).STR = Class(ClassNum).STR
        player(index).Char(charnum).DEF = Class(ClassNum).DEF
        player(index).Char(charnum).Speed = Class(ClassNum).Speed
        player(index).Char(charnum).Magi = Class(ClassNum).Magi
        
        If Class(ClassNum).Map <= 0 Then Class(ClassNum).Map = 1
        If Class(ClassNum).x < 0 Or Class(ClassNum).x > MAX_MAPX Then Class(ClassNum).x = Int(Class(ClassNum).x / 2)
        If Class(ClassNum).y < 0 Or Class(ClassNum).y > MAX_MAPY Then Class(ClassNum).y = Int(Class(ClassNum).y / 2)
        player(index).Char(charnum).Map = Class(ClassNum).Map
        player(index).Char(charnum).x = Class(ClassNum).x
        player(index).Char(charnum).y = Class(ClassNum).y
        
        player(index).Char(charnum).HP = GetPlayerMaxHP(index)
        player(index).Char(charnum).MP = GetPlayerMaxMP(index)
        player(index).Char(charnum).SP = GetPlayerMaxSP(index)
        
        player(index).Char(charnum).MAXHP = GetPlayerMaxHP(index)
        player(index).Char(charnum).MAXMP = GetPlayerMaxMP(index)
        player(index).Char(charnum).MAXSP = GetPlayerMaxSP(index)
        
        player(index).Char(charnum).head = headc
        player(index).Char(charnum).body = bodyc
        player(index).Char(charnum).leg = logc
        
        player(index).Char(charnum).Paperdoll = 1
        
'        Append name to file
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Append As #f
        Print #f, Name
        Close #f
        
        Call SavePlayer(index)
        
        Exit Sub
    End If
End Sub

Sub DelChar(ByVal index As Long, ByVal charnum As Long)
    
    Call DeleteName(player(index).Char(charnum).Name)
    Call ClearChar(index, charnum)
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
    Dim i As Integer
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next i
End Sub

Sub LoadClasses()
    Dim FileName As String
    Dim i As Long
    On Error GoTo ClassErr
    Call CheckClasses
    
    FileName = App.Path & "\Classes\info.ini"
    
    MAX_CLASSES = Val(GetVar(FileName, "INFO", "MaxClasses"))
    
    ReDim Class(0 To MAX_CLASSES) As ClassRec
    
    Call ClearClasses
    
    For i = 0 To MAX_CLASSES
        temp = i / MAX_CLASSES * 100
        Call SetStatus("Loading classes... " & temp & "%")
        FileName = App.Path & "\Classes\Class" & i & ".ini"
        
'        Check if class exists
        If Not FileExist("\Classes\Class" & i & ".ini") Then
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
            Call PutVar(FileName, "CLASS", "Locked", STR(Class(i).locked))
            Call MsgBox("Class " & i & " not found, created.", vbInformation)
        End If
        
        
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
        Class(i).locked = Val(GetVar(FileName, "CLASS", "Locked"))
        
        DoEvents
    Next i
    Exit Sub
    
ClassErr:
    Call MsgBox("Error loading class " & i & ". Check that all the variables in your class files exist!", vbCritical)
    Call DestroyServer
    End
End Sub

Sub SaveClasses()
    Dim FileName As String
    Dim i As Long
    
    FileName = App.Path & "\Classes\info.ini"
    
    If Not FileExist("Classes\info.ini") Then
        Call PutVar(FileName, "INFO", "MaxClasses", 3)
        Call PutVar(FileName, "INFO", "MaxSkills", 25)
        Call PutVar(FileName, "INFO", "StatPoints", 0)
        Call PutVar(FileName, "INFO", "SkillPoints", 0)
    End If
    
    For i = 0 To MAX_CLASSES
        temp = i / MAX_CLASSES * 100
        Call SetStatus("Saving classes... " & temp & "%")
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
            Call PutVar(FileName, "CLASS", "Locked", STR(Class(i).locked))
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
    
    MAX_CLASSES = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class2(0 To MAX_CLASSES) As ClassRec
    
    Call ClearClasses2
    
    For i = 0 To MAX_CLASSES
        temp = i / MAX_CLASSES * 100
        Call SetStatus("Loading first class advandcement... " & temp & "%")
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
    
    For i = 0 To MAX_CLASSES
        temp = i / MAX_CLASSES * 100
        Call SetStatus("Saving first class advandcement... " & temp & "%")
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
    
    MAX_CLASSES = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class3(0 To MAX_CLASSES) As ClassRec
    
    Call ClearClasses3
    
    For i = 0 To MAX_CLASSES
        temp = i / MAX_CLASSES * 100
        Call SetStatus("Loading second class advandcement... " & temp & "%")
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
    
    For i = 0 To MAX_CLASSES
        temp = i / MAX_CLASSES * 100
        Call SetStatus("Saving second class advandcement... " & temp & "%")
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

Sub SaveQuests()
    Dim i As Long
    
    Call SetStatus("Saving quests... ")
    For i = 1 To MAX_QUESTS
        If Not FileExist("Quests\Quest" & i & ".dat") Then
            temp = i / MAX_QUESTS * 100
            Call SetStatus("Saving quest... " & temp & "%")
            DoEvents
            Call SaveQuest(i)
        End If
    Next i
End Sub

Sub SaveQuest(ByVal QuestNum As Long)
    Dim FileName As String
    Dim f  As Long
    FileName = App.Path & "\quests\Quest" & QuestNum & ".dat"
    
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Quest(QuestNum)
    Close #f
End Sub

Sub LoadQuests()
    Dim FileName As String
    Dim i As Long
    Dim f As Long
    
    Call CheckQuests
    
    For i = 1 To MAX_QUESTS
        temp = i / MAX_QUESTS * 100
        Call SetStatus("Loading quests... " & temp & "%")
        
        FileName = App.Path & "\Quests\Quest" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Quest(i)
        Close #f
        
        DoEvents
    Next i
End Sub

Sub CheckQuests()
    Call SaveQuests
End Sub


Sub SaveSkills()
    Dim i As Long
    
    Call SetStatus("Saving skills... ")
    For i = 1 To MAX_SKILLS
        If Not FileExist("Skills\Skill" & i & ".dat") Then
            temp = i / MAX_SKILLS * 100
            Call SetStatus("Saving skills... " & temp & "%")
            DoEvents
            Call SaveSkill(i)
        End If
    Next i
End Sub

Sub SaveSkill(ByVal skillNum As Long)
    Dim FileName As String
    Dim f  As Long
    FileName = App.Path & "\skills\Skill" & skillNum & ".dat"
    
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , skill(skillNum)
    Close #f
End Sub

Sub LoadSkills()
    Dim FileName As String
    Dim i As Long
    Dim f As Long
    
    Call CheckSkills
    
    For i = 1 To MAX_SKILLS
        temp = i / MAX_SKILLS * 100
        Call SetStatus("Loading skills... " & temp & "%")
        
        FileName = App.Path & "\Skills\Skill" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , skill(i)
        Close #f
        
        DoEvents
    Next i
End Sub

Sub CheckSkills()
    Call SaveSkills
End Sub

Sub SaveItems()
    Dim i As Long
    
    Call SetStatus("Saving items... ")
    For i = 1 To MAX_ITEMS
        If Not FileExist("items\item" & i & ".dat") Then
            temp = i / MAX_ITEMS * 100
            Call SetStatus("Saving items... " & temp & "%")
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
    Put #f, , Item(ItemNum)
    Close #f
End Sub

Sub LoadItems()
    Dim FileName As String
    Dim i As Long
    Dim f As Long
    
    Call CheckItems
    
    For i = 1 To MAX_ITEMS
        temp = i / MAX_ITEMS * 100
        Call SetStatus("Loading items... " & temp & "%")
        
        FileName = App.Path & "\Items\Item" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Item(i)
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
            temp = i / MAX_SHOPS * 100
            Call SetStatus("Saving shops... " & temp & "%")
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
        temp = i / MAX_SHOPS * 100
        Call SetStatus("Loading shops... " & temp & "%")
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
            temp = i / MAX_SPELLS * 100
            Call SetStatus("Saving spells... " & temp & "%")
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
        temp = i / MAX_SPELLS * 100
        Call SetStatus("Loading spells... " & temp & "%")
        
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
            temp = i / MAX_NPCS * 100
            Call SetStatus("Saving npcs... " & temp & "%")
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
    Dim i As Integer
    Dim f As Long
    
    Call CheckNpcs
    
    For i = 1 To MAX_NPCS
        temp = i / MAX_NPCS * 100
        Call SetStatus("Loading npcs... " & temp & "%")
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
    Dim i As Integer
    Dim f As Integer
    
    If 0 + GetVar(App.Path & "\Data.ini", "CONFIG", "NonToScroll") = 1 Then
        For i = 1 To MAX_MAPS
            Call ClearMapScroll(i)
        Next i
        Call PutVar(App.Path & "\Data.ini", "CONFIG", "NonToScroll", 0)
    End If
    
    
    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        temp = i / MAX_MAPS * 100
        Call SetStatus("Loading maps... " & temp & "%")
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
    Dim i As Integer
    
    Call ClearMaps
    
    For i = 1 To MAX_MAPS
        FileName = "maps\map" & i & ".dat"
        
'        Check to see if map exists, if it doesn't, create it.
        If Not FileExist(FileName) Then
            temp = i / MAX_MAPS * 100
            Call SetStatus("Saving maps... " & temp & "%")
            DoEvents
            
            Call SaveMap(i)
        End If
    Next i
End Sub

Sub AddLog(ByVal text As String, ByVal FN As String)
    On Error Resume Next
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
    
'    Make sure the file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
'    Cut off last portion of ip
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
    
'    Destroy name from charlist
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
    Dim f As Long, i As Long
    
    FileName = App.Path & "\banlist.txt"
    
'    Make sure the file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
'    Cut off last portion of ip
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


Sub SaveLogs()
    Dim FileName As String
    Dim i As String, c As String
    On Error Resume Next
    
    If LCase(Dir(App.Path & "\logs", vbDirectory)) <> "logs" Then
        Call MkDir(App.Path & "\Logs")
    End If
    
    c = Time
    c = Replace(c, ":", "-", 1, -1, vbTextCompare)
    c = Replace(c, ":", "-", 1, -1, vbTextCompare)
    c = Replace(c, ":", "-", 1, -1, vbTextCompare)
    
    i = Date
    
    If LCase(Dir(App.Path & "\logs\" & i, vbDirectory)) <> i Then
        Call MkDir(App.Path & "\Logs\" & i & "\")
    End If
    
    Call MkDir(App.Path & "\Logs\" & i & "\" & c & "\")
    
    FileName = App.Path & "\Logs\" & i & "\" & c & "\Main.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(0).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & c & "\Broadcast.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(1).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & c & "\Global.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(2).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & c & "\Map.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(3).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & c & "\Private.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(4).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & c & "\Admin.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(5).text
    Close #1
    
    FileName = App.Path & "\Logs\" & i & "\" & c & "\Emote.txt"
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
        temp = i / MAX_ARROWS * 100
        Call SetStatus("Loading Arrows... " & temp & "%")
        Arrows(i).Name = GetVar(FileName, "Arrow" & i, "ArrowName")
        Arrows(i).Pic = GetVar(FileName, "Arrow" & i, "ArrowPic")
        Arrows(i).Range = GetVar(FileName, "Arrow" & i, "ArrowRange")
        Arrows(i).Amount = GetVar(FileName, "Arrow" & i, "ArrowAmount")
        
        DoEvents
    Next i
End Sub

Sub CheckArrows()
    If Not FileExist("Arrows.ini") Then
        Dim i As Long
        
        For i = 1 To MAX_ARROWS
            temp = i / MAX_ARROWS * 100
            Call SetStatus("Saving arrows... " & temp & "%")
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
        Arrows(i).Name = vbNullString
        Arrows(i).Pic = 0
        Arrows(i).Range = 0
        Arrows(i).Amount = 0
    Next i
End Sub

Sub SaveArrow(ByVal ArrowNum As Long)
    Dim FileName As String
    
    FileName = App.Path & "\Arrows.ini"
    
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowName", Trim(Arrows(ArrowNum).Name))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowPic", Val(Arrows(ArrowNum).Pic))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowRange", Val(Arrows(ArrowNum).Range))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowAmount", Val(Arrows(ArrowNum).Amount))
End Sub


