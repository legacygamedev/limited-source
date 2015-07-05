Attribute VB_Name = "modDatabase"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Const START_MAP = 1
Public Const START_X = MAX_MAPX / 2
Public Const START_Y = MAX_MAPY / 2

Public Const ADMIN_LOG = "admin.txt"
Public Const PLAYER_LOG = "player.txt"

'Database Stuff
Public Const strCONN = "DRIVER={MySQL};SERVER=127.0.0.1;DATABASE=mirage_online;UID=mo;PWD=mo;OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
'Connections....
Public Conn_Server As ADODB.Connection
Public Conn_Client As ADODB.Connection
'RecordSets
Public RS_LoadPlayer As ADODB.Recordset
Public RS_LoadChar As ADODB.Recordset
Public RS_SavePlayerNEW As ADODB.Recordset
Public RS_SavePlayerUPDATE As ADODB.Recordset
Public RS_SaveCharNEW As ADODB.Recordset
Public RS_SaveCharUPDATE As ADODB.Recordset
Dim RS_LOGS As ADODB.Recordset

Public Sub InitDB()
'Setup Connections
Set Conn_Client = New ADODB.Connection
Conn_Client.CursorLocation = adUseClient
Conn_Client.ConnectionString = strCONN
Conn_Client.Open

Set Conn_Server = New ADODB.Connection
Conn_Server.CursorLocation = adUseServer
Conn_Server.ConnectionString = strCONN
Conn_Server.Open

'Setup Connections

'Accounts
Set RS_SavePlayerNEW = New ADODB.Recordset
Set RS_SavePlayerUPDATE = New ADODB.Recordset
Set RS_LoadPlayer = New ADODB.Recordset
RS_SavePlayerNEW.Open "SELECT * FROM accounts LIMIT 1;", Conn_Client, adOpenStatic, adLockOptimistic

'Characters
Set RS_SaveCharNEW = New ADODB.Recordset
Set RS_SaveCharUPDATE = New ADODB.Recordset
Set RS_LoadChar = New ADODB.Recordset
RS_SaveCharNEW.Open "SELECT * FROM characters LIMIT 1;", Conn_Client, adOpenStatic, adLockOptimistic

'Logs
Set RS_LOGS = New ADODB.Recordset
RS_LOGS.Open "SELECT * FROM logs LIMIT 1;", Conn_Client, adOpenStatic, adLockOptimistic
End Sub

Public Sub CloseDB()
'Kill RecordSets

'Kill Logs
RS_LOGS.Close
Set RS_LOGS = Nothing

'Kill Characters
RS_SaveCharNEW.Close
Set RS_SaveCharNEW = Nothing
Set RS_SaveCharUPDATE = Nothing

'Kill Accounts
RS_SavePlayerNEW.Close
Set RS_SavePlayerNEW = Nothing
Set RS_SavePlayerUPDATE = Nothing

'Kill Connections
Conn_Server.Close
Set Conn_Server = Nothing
Conn_Client.Close
Set Conn_Client = Nothing
End Sub


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

Sub SavePlayer(ByVal Index As Long)
Dim RS As ADODB.Recordset
Dim I As Long
Dim X As Long
Dim Y As Long
Dim N As Long
Dim SPELLString As String
Dim INVString As String

'Save Account
With Player(Index)
    If .FKey = 0 Then
        'It is a new class, so we Insert.
        Set RS = RS_SavePlayerNEW
        RS.AddNew
    Else
        'It is an old class, so we update all fields.
        Set RS = RS_SavePlayerUPDATE
        RS.Open "SELECT * FROM accounts WHERE `FKey`=" & .FKey & ";", Conn_Client, adOpenStatic, adLockOptimistic
    End If
    RS!Login = Trim(.Login)
    RS!Password = Encrypt(DBKEY, Trim(.Password))
    RS!HDModel = Trim(.HDModel)
    RS!HDSerial = Trim(.HDSerial)
    RS.Update
    
    If .FKey = 0 Then
        'Grab New FKey
        .FKey = RS!FKey
    Else
        RS.Close
    End If
End With


'Do Characters
For I = 1 To MAX_CHARS
    INVString = ""
    SPELLString = ""
    With Player(Index).Char(I)
        If Trim(.Name) <> "" Then
            If .FKey = 0 Then
                'It is a new class, so we Insert.
                Set RS = RS_SaveCharNEW
                RS.AddNew
            Else
                'It is an old class, so we update all fields.
                Set RS = RS_SaveCharUPDATE
                RS.Open "SELECT * FROM characters WHERE `FKey`=" & .FKey & ";", Conn_Client, adOpenStatic, adLockOptimistic
            End If
            RS!Account = Player(Index).FKey
            RS!Name = Trim(.Name)
            RS!Sex = .Sex
            RS!Class = .Class
            RS!Sprite = .Sprite
            RS!Level = .Level
            RS!Exp = .Exp
            RS!Access = .Access
            RS!PK = .PK
            RS!Guild = .Guild
            RS!HP = .HP
            RS!MP = .MP
            RS!SP = .SP
            RS!str = .str
            RS!DEF = .DEF
            RS!SPEED = .SPEED
            RS!MAGI = .MAGI
            RS!POINTS = .POINTS
            RS!ArmorSlot = .ArmorSlot
            RS!WeaponSlot = .WeaponSlot
            RS!HelmetSlot = .HelmetSlot
            RS!ShieldSlot = .ShieldSlot
            RS!Map = .Map
            RS!X = .X
            RS!Y = .Y
            
            'Inventory
            For X = 1 To MAX_INV
                INVString = INVString & .Inv(X).Num & "-"
                INVString = INVString & .Inv(X).Value & "-"
                INVString = INVString & .Inv(X).Dur & "|"
            Next X
            RS!Inventory = INVString
    
            'Spells
            For X = 1 To MAX_PLAYER_SPELLS
                If X <> MAX_PLAYER_SPELLS Then
                    SPELLString = SPELLString & .Spell(X) & "-"
                Else
                    SPELLString = SPELLString & .Spell(X)
                End If
            Next X
            RS!Spells = SPELLString
            
            RS.Update
            
            If .FKey = 0 Then
                'Grab New FKey
                .FKey = RS!FKey
            Else
                RS.Close
            End If
        End If
    End With
Next I
End Sub

Public Function CharCount(Account As Long) As Long
'Setup database stuff
Dim RS As ADODB.Recordset
Dim I As Long
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM characters WHERE `account`=" & Account & " LIMIT " & MAX_CHARS & ";", Conn_Client, adOpenStatic, adLockReadOnly
RS.MoveFirst
I = RS.RecordCount
RS.Close
Set RS = Nothing
CharCount = I
End Function

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
Dim I As Long
Dim N As Long
Dim cCount As Long
Dim chCount As Long
Dim RS As ADODB.Recordset
Dim Inv1() As String
Dim Inv2() As String
Dim Spells() As String

'ClearPlayer
Call ClearPlayer(Index)

cCount = 0
'Setup database stuff
Set RS = RS_LoadPlayer
RS.Open "SELECT * FROM accounts WHERE `login`='" & Trim(Name) & "' LIMIT 1;", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst

'Load Account
With Player(Index)
    .FKey = CLng(RS("FKey"))
    .Login = Trim(CStr(RS("Login")))
    .Password = Decrypt(DBKEY, CStr(RS("Password")))
    .HDModel = Trim(CStr(RS("HDModel")))
    .HDSerial = Trim(CStr(RS("HDSerial")))
End With
RS.Close

'Load Characters
chCount = CharCount(Player(Index).FKey)

'Setup database stuff
Set RS = RS_LoadChar
RS.Open "SELECT * FROM characters WHERE `account`=" & Player(Index).FKey & " LIMIT " & MAX_CHARS & ";", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
For I = 1 To chCount
    With Player(Index).Char(I)
        ' General
        .FKey = CLng(RS("Fkey"))
        .Name = Trim(CStr(RS("Name")))
        .Sex = CByte(RS("Sex"))
        .Class = CByte(RS("Class"))
        .Sprite = CInt(RS("Sprite"))
        .Level = CByte(RS("Level"))
        .Exp = CLng(RS("Exp"))
        .Access = CByte(RS("Access"))
        .PK = CByte(RS("PK"))
        .Guild = CByte(RS("Guild"))
        
        ' Vitals
        .HP = CLng(RS("HP"))
        .MP = CLng(RS("MP"))
        .SP = CLng(RS("SP"))
        
        ' Stats
        .str = CByte(RS("STR"))
        .DEF = CByte(RS("DEF"))
        .SPEED = CByte(RS("SPEED"))
        .MAGI = CByte(RS("MAGI"))
        .POINTS = CByte(RS("POINTS"))
        
        ' Worn equipment
        .ArmorSlot = CByte(RS("ArmorSlot"))
        .WeaponSlot = CByte(RS("WeaponSlot"))
        .HelmetSlot = CByte(RS("HelmetSlot"))
        .ShieldSlot = CByte(RS("ShieldSlot"))
        
        ' Position
        .Map = CInt(RS("Map"))
        .X = CByte(RS("X"))
        .Y = CByte(RS("Y"))
        .Dir = CByte(RS("Dir"))
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If .Map = 0 Then
            .Map = START_MAP
            .X = START_X
            .Y = START_Y
        End If
        
        ' Inventory
        Inv1 = Split(RS("Inventory"), "|")
                
        For N = 1 To MAX_INV
            Inv2 = Split(Inv1(N - 1), "-")
            .Inv(N).Num = CByte(Inv2(0))
            .Inv(N).Value = CLng(Inv2(1))
            .Inv(N).Dur = CInt(Inv2(2))
        Next N
        
        ' Spells
        Spells = Split(RS("Spells"), "-")
        For N = 1 To MAX_PLAYER_SPELLS
            .Spell(N) = Spells(N - 1)
        Next N
    End With
    If I <> chCount Then
        RS.MoveNext
    End If
Next I
RS.Close
End Sub

Function AccountExist(ByVal Name As String) As Boolean
Dim cLogin As String
Dim iExist As Boolean
Dim RS As ADODB.Recordset

'Cleanse the strings
cLogin = Cleanse(Trim(Name))

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM accounts WHERE `login`='" & cLogin & "';", Conn_Server, adOpenStatic, adLockReadOnly

iExist = False
If Not (RS.BOF And RS.EOF) Then
    RS.MoveFirst
    iExist = True
Else
    iExist = False
End If
RS.Close
Set RS = Nothing
AccountExist = iExist
End Function

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If Trim(Player(Index).Char(CharNum).Name) <> "" Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Public Function SerialsBlank(ByVal Login As String) As Boolean
Dim RS As ADODB.Recordset
Dim cLogin As String

SerialsBlank = True

'Cleanse the strings
cLogin = Cleanse(Trim(Login))

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM accounts WHERE `login`='" & cLogin & "';", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
HD = CStr(RS("HDModel"))
If Trim(HD) = "" Then
    SerialsBlank = True
    Exit Function
End If
HD = CStr(RS("HDSerial"))
If Trim(HD) = "" Then
    SerialsBlank = True
    Exit Function
End If
DoEvents
RS.Close
Set RS = Nothing
SerialsBlank = False

End Function

Public Function SerialsOK(ByVal Login As String, ByVal HDModel As String, ByVal HDSerial As String) As Boolean
Dim cLogin As String
Dim cPassword As String
Dim iExist As Boolean
Dim RS As ADODB.Recordset


'Cleanse the strings
cLogin = Cleanse(Trim(Login))

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM accounts WHERE `Login` ='" & cLogin & "' AND `HDModel`='" & HDModel & "' AND `HDSerial`='" & HDSerial & "';", Conn_Server, adOpenStatic, adLockReadOnly

iExist = False
If Not (RS.BOF And RS.EOF) Then
    RS.MoveFirst
    iExist = True
Else
    iExist = False
End If
RS.Close

Set RS = Nothing
SerialsOK = iExist
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim cLogin As String
Dim cPassword As String
Dim iExist As Boolean
Dim RS As ADODB.Recordset


'Cleanse the strings
cLogin = Cleanse(Trim(Name))
cPassword = Encrypt(DBKEY, Trim(Password))

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM accounts WHERE `login`='" & cLogin & "' AND `password`='" & cPassword & "';", Conn_Server, adOpenStatic, adLockReadOnly

iExist = False
If Not (RS.BOF And RS.EOF) Then
    RS.MoveFirst
    iExist = True
Else
    iExist = False
End If
RS.Close

Set RS = Nothing
PasswordOK = iExist
End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String, ByVal HDModel As String, ByVal HDSerial As String)
Dim I As Long

With Player(Index)
    .Login = Name
    .Password = Password
    .HDModel = HDModel
    .HDSerial = HDSerial
End With

For I = 1 To MAX_CHARS
    Call ClearChar(Index, I)
Next I

Call SavePlayer(Index)
End Sub

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
Dim F As Long

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
                    
        Player(Index).Char(CharNum).str = Class(ClassNum).str
        Player(Index).Char(CharNum).DEF = Class(ClassNum).DEF
        Player(Index).Char(CharNum).SPEED = Class(ClassNum).SPEED
        Player(Index).Char(CharNum).MAGI = Class(ClassNum).MAGI
        
        Player(Index).Char(CharNum).Map = START_MAP
        Player(Index).Char(CharNum).X = START_X
        Player(Index).Char(CharNum).Y = START_Y
            
        Player(Index).Char(CharNum).HP = GetPlayerMaxHP(Index)
        Player(Index).Char(CharNum).MP = GetPlayerMaxMP(Index)
        Player(Index).Char(CharNum).SP = GetPlayerMaxSP(Index)
                
        
        'Call SavePlayer(Index)
            
        Exit Sub
    End If
End Sub

Sub DelChar(ByVal Index As Long, ByVal CharNum As Long)
Dim f1 As Long, f2 As Long
Dim s As String

Call DeleteChar(Player(Index).Char(CharNum).FKey)
Call ClearChar(Index, CharNum)
'No need to save player now :)
End Sub

Function FindChar(ByVal Name As String) As Boolean
'NOTE: To find a character in the database list, to see if it already exists :)
Dim cLogin As String
Dim iExist As Boolean
Dim RS As ADODB.Recordset

'Cleanse the strings
cLogin = Cleanse(Trim(Name))

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM characters WHERE `name`='" & cLogin & "';", Conn_Server, adOpenStatic, adLockReadOnly

iExist = False
If Not (RS.BOF And RS.EOF) Then
    RS.MoveFirst
    iExist = True
Else
    iExist = False
End If
RS.Close
Set RS = Nothing
FindChar = iExist
End Function

Sub SaveAllPlayersOnline()
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SavePlayer(I)
            DoEvents
        End If
    Next I
End Sub

Sub LoadClasses()
Dim I As Long
Dim cCount As Long
Dim RS As ADODB.Recordset

cCount = GetRecordCountByFKey("classes")
Max_Classes = cCount - 1
ReDim Class(0 To Max_Classes) As ClassRec
Call ClearClasses
    
'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM classes;", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
    
For I = 1 To cCount
    Call SetStatus("Loading Class " & I & "/" & cCount)
    With Class(I - 1)
        .FKey = CLng(RS("FKey"))
        .Name = CStr(RS("Name"))
        .Sprite = CLng(RS("Sprite"))
        .str = CLng(RS("STR"))
        .DEF = CLng(RS("DEF"))
        .SPEED = CLng(RS("SPEED"))
        .MAGI = CLng(RS("MAGI"))
    End With
    DoEvents
    If I <> cCount Then
        RS.MoveNext
    End If
Next I

RS.Close
Set RS = Nothing
End Sub

Sub SaveClasses()
Dim I As Long
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
For I = 0 To Max_Classes
    With Class(I)
        If .FKey = 0 Then
            'It is a new class, so we Insert.
            RS.Open "SELECT * FROM classes;", Conn_Client, adOpenStatic, adLockOptimistic
            RS.AddNew
        Else
            'It is an old class, so we update all fields.
            RS.Open "SELECT * FROM classes WHERE `FKey`=" & .FKey & ";", Conn_Client, adOpenStatic, adLockOptimistic
        End If
        
        If Trim(.Name) = "" Then .Name = " "
        RS!Name = .Name
        RS!Sprite = .Sprite
        RS!str = .str
        RS!DEF = .DEF
        RS!SPEED = .SPEED
        RS!MAGI = .MAGI
        
        RS.Update
        
        If .FKey = 0 Then
            'Grab New FKey
            .FKey = RS!FKey
        End If
    
        RS.Close
    End With
Next I
Set RS = Nothing
End Sub

Sub SaveItemArray(lFROM As Long, lTO As Long)
Dim I As Long
    
For I = lFROM To lTO
    Call SaveItem(I)
Next I
End Sub


Sub SaveItems()
Dim I As Long
    
    For I = 1 To MAX_ITEMS
        Call SaveItem(I)
    Next I
End Sub

Sub SaveItem(ByVal ItemNum As Long)
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
With Item(ItemNum)
    If .FKey = 0 Then
        'It is a new class, so we Insert.
        RS.Open "SELECT * FROM items;", Conn_Client, adOpenStatic, adLockOptimistic
        RS.AddNew
    Else
        'It is an old class, so we update all fields.
        RS.Open "SELECT * FROM items WHERE `FKey`=" & .FKey & ";", Conn_Client, adOpenStatic, adLockOptimistic
    End If
    If Trim(.Name) = "" Then .Name = " "
    RS!Name = .Name
    RS!Pic = .Pic
    RS!Type = .Type
    RS!Data1 = .Data1
    RS!Data2 = .Data2
    RS!Data3 = .Data3
    RS!Unbreakable = .Unbreakable
    RS!Locked = .Locked
    RS!Disabled = .Disabled
    RS!Assigned = .Assigned
    
    RS.Update
    
    If .FKey = 0 Then
        'Grab New FKey
        .FKey = RS!FKey
    End If

    RS.Close
End With
Set RS = Nothing
End Sub

Sub SaveSerials(ByVal Login As String, ByVal HDModel As String, ByVal HDSerial As String)
Dim RS As ADODB.Recordset
Dim cLogin As String
cLogin = Cleanse(Login)
Set RS = New ADODB.Recordset

'It is an old class, so we update all fields.
RS.Open "SELECT * FROM accounts WHERE `Login`='" & .cLogin & "';", Conn_Client, adOpenStatic, adLockOptimistic
RS!HDModel = HDModel
RS!HDSerial = HDSerial
RS.Update
RS.Close
Set RS = Nothing
End Sub


Public Sub CheckItems()
'Check to see if max items is how many items there are in the database.
Dim RecCount As Long
Dim I As Long
Dim OffSet As Long
Dim RS As ADODB.Recordset

RecCount = GetRecordCount("Items")
OffSet = MAX_ITEMS - RecCount
If OffSet = 0 Then Exit Sub
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM items;", Conn_Client, adOpenStatic, adLockOptimistic
RS.MoveFirst
For I = 1 To OffSet
    'Debug.Print "ADDING Item " & I & "/" & OffSet
    RS.AddNew
    
    RS!Name = "  "
    RS!Pic = 0
    RS!Type = 0
    RS!Data1 = 0
    RS!Data2 = 0
    RS!Data3 = 0
    RS!Unbreakable = 0
    RS!Locked = 0
    RS!Disabled = 0
    RS!Assigned = 0
        
    RS.Update
Next I
RS.Close
Set RS = Nothing
End Sub

Sub CreateCache(CacheNum As Long)
'Load the Cache
Dim I As Long
Dim ItemNum As Long
Dim C As String
Dim D As String

Select Case CacheNum
    Case 1 'Items
        With Cache(CacheNum)
            .nCache = ""
            .nCache = "ITEMCACHE" & SEP_CHAR
            For I = 1 To MAX_ITEMS
                ItemNum = I
                If I <> MAX_ITEMS Then
                    C = C & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR
                Else
                    C = C & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type
                End If
            Next I
            'Add C to cache
            .nCache = .nCache & SEP_CHAR & C & SEP_CHAR & END_CHAR
            .nDate = ""
            .nDate = Date$
            .nTime = ""
            .nTime = Time$
        End With
    
    Case 2 'npc
        With Cache(CacheNum)
            .nCache = ""
            .nCache = "NPCCACHE" & SEP_CHAR
            For I = 1 To MAX_NPCS
                If I <> MAX_NPCS Then
                    C = C & Trim(Npc(I).Name) & SEP_CHAR & Npc(I).Sprite & SEP_CHAR
                Else
                    C = C & Trim(Npc(I).Name) & SEP_CHAR & Npc(I).Sprite
                End If
            Next I
            'Add C to cache
            .nCache = .nCache & SEP_CHAR & C & SEP_CHAR & END_CHAR
            .nDate = ""
            .nDate = Date$
            .nTime = ""
            .nTime = Time$
        End With
    
    Case 3 'shop
    
    Case 4 'spell
    
    Case 5 'classes
    
End Select
'Call SaveCache(CacheNum)
End Sub

Public Sub SaveCache(CacheNum As Long)
Dim EC As Long
Dim I As Long
'On Error GoTo ERR:
'Save the Cache
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
'RS.Open "SELECT * FROM cache WHERE `FKey`=" & CacheNum & ";", Conn_Client, adOpenStatic, adLockOptimistic
RS.Open "SELECT * FROM cache;", Conn_Client, adOpenStatic, adLockOptimistic '  WHERE `FKey`=" & CacheNum & ";", Conn_Client, adOpenStatic, adLockOptimistic
RS.AddNew
With Cache(CacheNum)
    RS!nDate = .nDate
    RS!nTime = .nTime
    RS!nCache = .nCache
    Debug.Print "DB: " & .nCache
End With
RS.Update
RS.Close
Set RS = Nothing
Exit Sub

ERR:
EC = Conn_Client.Errors.Count
Debug.Print "ERR Count: " & EC
EC = EC - 1
For I = 0 To EC
    Debug.Print "ERROR #" & I & ": " & Conn_Client.Errors.Item(I).Source
    
Next I
End Sub

Sub LoadCache()
Dim RS As ADODB.Recordset
Dim cCount As Long
Dim I As Long
Set RS = New ADODB.Recordset
cCount = GetRecordCount("cache")
RS.Open "SELECT * FROM cache;", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
For I = 1 To cCount
    Call SetStatus("Loading Cache " & I & "/" & cCount)
    With Cache(I)
        .nDate = Trim(CStr(RS("nDate")))
        .nTime = Trim(CStr(RS("nTime")))
        .nCache = Trim(CStr(RS("Cache")))
    End With
    DoEvents
    If I <> cCount Then
        RS.MoveNext
    End If
Next I
End Sub

Sub LoadItems()
Dim I As Long
Dim cCount As Long
Dim RS As ADODB.Recordset

'Check Items :)
Call CheckItems

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM items;", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
    
For I = 1 To MAX_ITEMS
    Call SetStatus("Loading Item " & I & "/" & MAX_ITEMS)
    With Item(I)
        .FKey = CLng(RS("FKey"))
        .Name = Trim(CStr(RS("Name")))
        .Pic = CLng(RS("Pic"))
        .Type = CLng(RS("Type"))
        .Data1 = CLng(RS("Data1"))
        .Data2 = CLng(RS("Data2"))
        .Data3 = CLng(RS("Data3"))
        .Unbreakable = CLng(RS("UnBreakAble"))
        .Locked = CLng(RS("Locked"))
        .Disabled = CLng(RS("Disabled"))
        .Assigned = CLng(RS("Assigned"))
    End With
    DoEvents
    If I <> MAX_ITEMS Then
        RS.MoveNext
    End If
Next I
RS.Close
Set RS = Nothing
'Create Item Cache
Call CreateCache(1)

End Sub

Sub SaveShops()
Dim I As Long

    For I = 1 To MAX_SHOPS
        Call SaveShop(I)
    Next I
End Sub

Sub SaveShop(ByVal ShopNum As Long)
Dim I As Long
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
                
With Shop(ShopNum)
    If .FKey = 0 Then
        'It is a new class, so we Insert.
        RS.Open "SELECT * FROM shops;", Conn_Client, adOpenStatic, adLockOptimistic
        RS.AddNew
    Else
        'It is an old class, so we update all fields.
        RS.Open "SELECT * FROM shops WHERE `FKey`=" & .FKey & ";", Conn_Client, adOpenStatic, adLockOptimistic
    End If
    If Trim(.Name) = "" Then .Name = " "
    RS!Name = .Name
    If Trim(.JoinSay) = "" Then .JoinSay = " "
    RS!JoinSay = .JoinSay
    If Trim(.LeaveSay) = "" Then .LeaveSay = " "
    RS!LeaveSay = .LeaveSay
    RS!FixesItems = .FixesItems
    
    For I = 1 To MAX_TRADES
        RS("GiveItem" & CStr(I)) = .TradeItem(I).GiveItem
        RS("GiveValue" & CStr(I)) = .TradeItem(I).GiveValue
        RS("GetItem" & CStr(I)) = .TradeItem(I).GetItem
        RS("GetValue" & CStr(I)) = .TradeItem(I).GetValue
    Next I
     
    RS.Update
    
    If .FKey = 0 Then
        'Grab New FKey
        .FKey = RS!FKey
    End If

    RS.Close
End With
Set RS = Nothing
End Sub

Sub LoadShops()
Dim I As Long
Dim X As Long
Dim cCount As Long
Dim RS As ADODB.Recordset

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM shops;", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
    
For I = 1 To MAX_SHOPS
    Call SetStatus("Loading Shop " & I & "/" & MAX_SHOPS)
    With Shop(I)
        .FKey = CLng(RS("FKey"))
        .Name = Trim(CStr(RS("Name")))
        .JoinSay = Trim(CStr(RS("JoinSay")))
        .LeaveSay = Trim(CStr(RS("LeaveSay")))
        .FixesItems = CLng(RS("FixesItems"))
    
        For X = 1 To MAX_TRADES
            .TradeItem(X).GiveItem = CLng(RS("GiveItem" & CStr(X)))
            .TradeItem(X).GiveValue = CLng(RS("GiveValue" & CStr(X)))
            .TradeItem(X).GetItem = CLng(RS("GetItem" & CStr(X)))
            .TradeItem(X).GetValue = CLng(RS("GetValue" & CStr(X)))
        Next X
    End With
    DoEvents
    If I <> MAX_SHOPS Then
        RS.MoveNext
    End If
Next I
RS.Close
Set RS = Nothing
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
                
With Spell(SpellNum)
    If .FKey = 0 Then
        'It is a new class, so we Insert.
        RS.Open "SELECT * FROM spells;", Conn_Client, adOpenStatic, adLockOptimistic
        RS.AddNew
    Else
        'It is an old class, so we update all fields.
        RS.Open "SELECT * FROM spells WHERE `FKey`=" & .FKey & ";", Conn_Client, adOpenStatic, adLockOptimistic
    End If
    If Trim(.Name) = "" Then .Name = " "
    RS!Name = .Name
    RS!ClassReq = .ClassReq
    RS!Type = .Type
    RS!Data1 = .Data1
    RS!Data2 = .Data2
    RS!Data3 = .Data3
    
    RS.Update
    
    If .FKey = 0 Then
        'Grab New FKey
        .FKey = RS!FKey
    End If

    RS.Close
End With
Set RS = Nothing
End Sub

Sub SaveSpells()
Dim I As Long

    For I = 1 To MAX_SPELLS
        Call SaveSpell(I)
    Next I
End Sub

Sub LoadSpells()
Dim I As Long
Dim cCount As Long
Dim RS As ADODB.Recordset

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM spells;", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
    
For I = 1 To MAX_SPELLS
    Call SetStatus("Loading Spell " & I & "/" & MAX_SPELLS)
    With Spell(I)
        .FKey = CLng(RS("FKey"))
        .Name = Trim(CStr(RS("Name")))
        .ClassReq = CLng(RS("ClassReq"))
        .Type = CLng(RS("Type"))
        .Data1 = CLng(RS("Data1"))
        .Data2 = CLng(RS("Data2"))
        .Data3 = CLng(RS("Data3"))
    End With
    DoEvents
    If I <> MAX_SPELLS Then
        RS.MoveNext
    End If
Next I
RS.Close
Set RS = Nothing
End Sub

Sub SaveNpcs()
Dim I As Long
    
    For I = 1 To MAX_NPCS
        Call SaveNpc(I)
    Next I
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
                
With Npc(NpcNum)
    If .FKey = 0 Then
        'It is a new class, so we Insert.
        RS.Open "SELECT * FROM npcs;", Conn_Client, adOpenStatic, adLockOptimistic
        RS.AddNew
    Else
        'It is an old class, so we update all fields.
        RS.Open "SELECT * FROM npcs WHERE `FKey`=" & .FKey & ";", Conn_Client, adOpenStatic, adLockOptimistic
    End If
    If Trim(.Name) = "" Then .Name = " "
    RS!Name = .Name
    If Trim(.AttackSay) = "" Then .AttackSay = " "
    RS!AttackSay = .AttackSay
    RS!Sprite = .Sprite
    RS!SpawnSecs = .SpawnSecs
    RS!Behavior = .Behavior
    RS!Range = .Range
    RS!DropChance = .DropChance
    RS!DropItem = .DropItem
    RS!DropItemValue = .DropItemValue
    RS!str = .str
    RS!DEF = .DEF
    RS!SPEED = .SPEED
    RS!MAGI = .MAGI
    
    RS.Update
    
    If .FKey = 0 Then
        'Grab New FKey
        .FKey = RS!FKey
    End If

    RS.Close
End With
Set RS = Nothing
End Sub

Sub LoadNpcs()
Dim I As Long
Dim cCount As Long
Dim RS As ADODB.Recordset

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM npcs;", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
    
For I = 1 To MAX_NPCS
    Call SetStatus("Loading NPC " & I & "/" & MAX_NPCS)
    With Npc(I)
        .FKey = CLng(RS("FKey"))
        .Name = Trim(CStr(RS("Name")))
        .AttackSay = Trim(CStr(RS("AttackSay")))
        .Sprite = CLng(RS("Sprite"))
        .SpawnSecs = CLng(RS("SpawnSecs"))
        .Behavior = CLng(RS("Behavior"))
        .Range = CLng(RS("Range"))
        .DropChance = CLng(RS("DropChance"))
        .DropItem = CLng(RS("DropItem"))
        .DropItemValue = CLng(RS("DropItemValue"))
        .str = CLng(RS("STR"))
        .DEF = CLng(RS("DEF"))
        .SPEED = CLng(RS("SPEED"))
        .MAGI = CLng(RS("MAGI"))
    End With
    DoEvents
    If I <> MAX_NPCS Then
        RS.MoveNext
    End If
Next I
RS.Close
Set RS = Nothing
'Create NPC Cache
Call CreateCache(2)

End Sub

Sub SaveMap(ByVal MapNum As Long)
Dim X As Long
Dim Y As Long
Dim MapString As String
Dim NPCString As String
Dim RS As ADODB.Recordset
'Set Conn = New ADODB.Connection
Set RS = New ADODB.Recordset
                
With Map(MapNum)
    If .FKey = 0 Then
        'It is a new class, so we Insert.
        RS.Open "SELECT * FROM maps;", Conn_Client, adOpenStatic, adLockOptimistic
        RS.AddNew
    Else
        'It is an old class, so we update all fields.
        RS.Open "SELECT * FROM maps WHERE `FKey`=" & .FKey & ";", Conn_Client, adOpenStatic, adLockOptimistic
    End If
    If Trim(.Name) = "" Then .Name = " "
    RS!Name = .Name
    RS!Revision = .Revision
    RS!Moral = .Moral
    RS!Up = .Up
    RS!Down = .Down
    RS!mLeft = .Left
    RS!mRight = .Right
    RS!Music = .Music
    RS!BootMap = .BootMap
    RS!BootX = .BootX
    RS!BootY = .BootY
    RS!Shop = .Shop
    RS!Indoors = .Indoors
    
    'Tiles
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            If .Tile(X, Y).Ground < 0 Then .Tile(X, Y).Ground = 0
            MapString = MapString & .Tile(X, Y).Ground & "-"
            If .Tile(X, Y).Mask < 0 Then .Tile(X, Y).Mask = 0
            MapString = MapString & .Tile(X, Y).Mask & "-"
            If .Tile(X, Y).Anim < 0 Then .Tile(X, Y).Anim = 0
            MapString = MapString & .Tile(X, Y).Anim & "-"
            If .Tile(X, Y).Fringe < 0 Then .Tile(X, Y).Fringe = 0
            MapString = MapString & .Tile(X, Y).Fringe & "-"
            If .Tile(X, Y).Type < 0 Then .Tile(X, Y).Type = 0
            MapString = MapString & .Tile(X, Y).Type & "-"
            If .Tile(X, Y).Data1 < 0 Then .Tile(X, Y).Data1 = 0
            MapString = MapString & .Tile(X, Y).Data1 & "-"
            If .Tile(X, Y).Data2 < 0 Then .Tile(X, Y).Data2 = 0
            MapString = MapString & .Tile(X, Y).Data2 & "-"
            If .Tile(X, Y).Data3 < 0 Then .Tile(X, Y).Data3 = 0
            MapString = MapString & .Tile(X, Y).Data3 & "|"
        Next Y
    Next X
    
    RS!Tiles = MapString
    'NPCs
    For X = 1 To MAX_MAP_NPCS
        If X <> MAX_MAP_NPCS Then
            NPCString = NPCString & CStr(.Npc(X)) & "-"
        Else
            NPCString = NPCString & CStr(.Npc(X))
        End If
    Next X
    RS!NPCs = NPCString
    
    
    RS.Update
    
    If .FKey = 0 Then
        'Grab New FKey
        .FKey = RS!FKey
    End If

    RS.Close
End With
Set RS = Nothing
End Sub

Sub SaveMaps()
Dim FileName As String
Dim I As Long
Dim F As Long

    For I = 1 To MAX_MAPS
        Call SaveMap(I)
    Next I
End Sub

Sub LoadMaps()
Dim I As Long
Dim Q As Long
Dim X As Long
Dim Y As Long
Dim cName As String
Dim RS As ADODB.Recordset
Dim Tiles() As String
Dim TileSet() As String
Dim TileSet1() As String
Dim NPCs() As String

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM maps;", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
    
X = 0
Y = 0
    
For I = 1 To MAX_MAPS
    Call SetStatus("Loading Map " & I & "/" & MAX_MAPS)
    With Map(I)
        .FKey = CLng(RS("FKey"))
        .Name = Trim(CStr(RS("Name")))
        .Revision = CLng(RS("Revision"))
        .Moral = CByte(RS("Moral"))
        .Up = CInt(RS("Up"))
        .Down = CInt(RS("Down"))
        .Left = CInt(RS("mLeft"))
        .Right = CInt(RS("mRight"))
        .Music = CByte(RS("Music"))
        .BootMap = CInt(RS("BootMap"))
        .BootX = CByte(RS("BootX"))
        .BootY = CByte(RS("BootY"))
        .Shop = CByte(RS("Shop"))
        .Indoors = CByte(RS("Indoors"))
        'Load Tiles
        Tiles = Split(Trim(CStr(RS("Tiles"))), "|")
        Q = 0
        For X = 0 To MAX_MAPX
            For Y = 0 To MAX_MAPY
                TileSet() = Split(Tiles(Q), "-")
                .Tile(X, Y).Ground = CInt(TileSet(0))
                .Tile(X, Y).Mask = CInt(TileSet(1))
                .Tile(X, Y).Anim = CInt(TileSet(2))
                .Tile(X, Y).Fringe = CInt(TileSet(3))
                .Tile(X, Y).Type = CInt(TileSet(4))
                .Tile(X, Y).Data1 = CInt(TileSet(5))
                .Tile(X, Y).Data2 = CInt(TileSet(6))
                .Tile(X, Y).Data3 = CInt(TileSet(7))
                Q = Q + 1
            Next Y
        Next X
        
        'Load NPCs
        NPCs = Split(Trim(CStr(RS.Fields("NPCS"))), "-")
        For Q = LBound(NPCs) To UBound(NPCs)
            .Npc(Q + 1) = CByte(NPCs(Q))
        Next Q
        
    End With
    DoEvents
    If I <> MAX_MAPS Then
        RS.MoveNext
    End If
Next I
RS.Close
Set RS = Nothing
End Sub

Sub AddLog(ByVal Text As String, ByVal FN As String)
RS_LOGS.AddNew
RS_LOGS!nDate = Date$
RS_LOGS!nTime = Time$
RS_LOGS!nType = FN
RS_LOGS!Entry = Text
RS_LOGS.Update
End Sub

Sub DeleteAccount(FKey As Long)
'NOTE: This will delete the account in the Database :)
Call Conn_Server.Execute("DELETE FROM accounts WHERE `fkey`=" & FKey & ";")
End Sub

Sub DeleteChar(FKey As Long)
'NOTE: This will delete the character in the Database :)
Call Conn_Server.Execute("DELETE FROM characters WHERE `fkey`=" & FKey & ";")
End Sub


Public Function GetRecordCountByFKey(dbTable As String) As Long
Dim RS As ADODB.Recordset
Dim FKey As Long

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT FKEY FROM " & dbTable & " ORDER BY FKey DESC LIMIT 1;", Conn_Server, adOpenStatic, adLockReadOnly

RS.MoveFirst
FKey = CLng(RS("FKey"))

RS.Close
Set RS = Nothing
GetRecordCountByFKey = FKey
End Function

Public Function Cleanse(Dirty As String) As String
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 11/15/2003  Shannara   Created Function
'****************************************************************

'THIS FUNCTION WILL ESCAPE ALL SINGLE QUOTE CHARACTERS IN AN EFFORT
'TO PREVENT SQL INJECTION ATTACKS. IT IS RECCOMENDED THAT ALL TAINTED DATA BE
'PASSED THROUGH THIS FUNCTION PRIOR TO BEING USED IN DYNAMIC SQL QUERIES.
'
'*******************************************
'NOTE: YOUR BROWSER MAY SHOW SPACES IN THE STRINGS (I.E.  " '  " ) THERE SHOULD BE NO WHITESPACES IN ANY OF THE STRINGS
'*******************************************
'
'WRITTEN BY: MIKE HILLYER
'LAST MODIFIED: 14JUN2003
    Cleanse = Replace(Dirty, "'", "\'")
'CLEVER HACKERS COULD PASS \' TO THIS FUNCTION, WHICH WOULD BECOME \\'
' \\' GETS INTERPRETED AS \', WITH THE \ BEING IGNORED AND THE ' GETTING
'INTERPRETED, THUS BYPASSING THIS FUNCTION, SO WE SHALL LOOP UNTIL WE ARE LEFT
'WITH JUST \' WHICH ESCAPES THE QUOTE, LOOP IS NEEDED BECAUSE A HACKER COULD TYPE
' \\\' IF WE SIMPLY CHECKED FOR \\' AFTER DOING THE INITIAL REPLACE.
    Do While InStr(Cleanse, "\\'")
        Cleanse = Replace(Cleanse, "\\'", "\'")
    Loop
End Function

Public Sub SetOnline(SetO As Long)
Call Conn_Client.Execute("UPDATE settings SET `online`=" & SetO & ";")
End Sub

Public Sub SetPlayers(SetO As Long)
Call Conn_Client.Execute("UPDATE settings SET `players`=" & SetO & ";")
End Sub

Public Sub SetMOTD(mMOTD As String)
Call Conn_Client.Execute("UPDATE settings SET `motd`='" & Trim(mMOTD) & "';")
MOTD = Trim(mMOTD)
End Sub

Public Sub GetMOTD()
Dim RS As ADODB.Recordset

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM settings LIMIT 1;", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
'Load MOTD
MOTD = CStr(RS("MOTD"))
RS.Close
Set RS = Nothing
End Sub

Function IsBanned(ByVal IP As String) As Boolean
Dim iExist As Boolean
Dim RS As ADODB.Recordset

'Setup database stuff
Set RS = New ADODB.Recordset

RS.Open "SELECT * FROM bans WHERE `bannedip`='" & IP & "';", Conn_Server, adOpenStatic, adLockReadOnly

iExist = False
If Not (RS.BOF And RS.EOF) Then
    RS.MoveFirst
    iExist = True
Else
    iExist = False
End If
RS.Close
Set RS = Nothing
IsBanned = iExist
End Function


Function GetBanListCount() As Long
'Setup database stuff
Dim RS As ADODB.Recordset
Dim I As Long
'Set Conn = New ADODB.Connection
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM bans;", Conn_Client, adOpenStatic, adLockReadOnly
RS.MoveFirst
I = RS.RecordCount
RS.Close
Set RS = Nothing
GetBanListCount = I
End Function

Sub SendBanList(Index As Long)
Dim RS As ADODB.Recordset
Dim I As Long
Dim cCount As Long


cCount = GetBanListCount
If cCount <= 0 Then
    Call PlayerMsg(Index, "*** Empty Banlist.", White)
    Exit Sub
End If

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM bans;", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
For I = 1 To cCount
    Call PlayerMsg(Index, "*** " & RS("FKey") & ": Banned IP " & RS("BannedIP") & " by " & RS("BannedBY"), White)
Next I
Call PlayerMsg(Index, "*** End of Banlist.", White)
If I <> cCount Then
    RS.MoveNext
End If
RS.Close
Set RS = Nothing
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
Dim RS As ADODB.Recordset
Dim FileName, IP As String
Dim F As Long, I As Long
Dim BB As String
Dim Q() As String
Dim E As Long
Set RS = New ADODB.Recordset
Dim G As String
Dim X As Long

RS.Open "SELECT * FROM bans LIMIT 1;", Conn_Client, adOpenStatic, adLockOptimistic

' Cut off last portion of ip
IP = GetPlayerIP(BanPlayerIndex)
 
'Ban all ips :)
Q = Split(IP, ".")
X = UBound(Q)
For E = 0 To 255
    Q(X) = CStr(E)
    G = Join(Q, ".")
    
    RS.AddNew
    RS!nDate = Date$
    RS!nTime = Time$
    RS!BannedIP = G
    If BannedByIndex = 0 Then
        BB = "SERVER"
    Else
        BB = GetPlayerName(BannedByIndex)
    End If
    RS!BannedBY = BB
    RS.Update
Next E
        
RS.Close
Set RS = Nothing
If BB <> "SERVER" Then
    Call AdminMsg(GetPlayerName("*** " & BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
    'Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AlertMsg(BanPlayerIndex, "You have been banned from " & GAME_NAME & "!")
Else
    Call AlertMsg(BanPlayerIndex, " ")
End If
Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & BB & ".", ADMIN_LOG)
End Sub

Public Sub CreateMaps(cMaps As Long)
Dim X As Long
Dim Y As Long
Dim I As Long
Dim MapString As String
Dim NPCString As String
Dim Conn As ADODB.Connection
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
Set Conn = New ADODB.Connection
Conn.Open strCONN
RS.Open "SELECT * FROM maps LIMIT 1;", Conn, adOpenStatic, adLockOptimistic

For I = 1 To cMaps
    MapString = ""
    NPCString = ""
    RS.AddNew
    RS!Name = "New Map"
    RS!Revision = 0
    RS!Moral = 0
    RS!Up = 0
    RS!Down = 0
    RS!mLeft = 0
    RS!mRight = 0
    RS!Music = 0
    RS!BootMap = 0
    RS!BootX = 0
    RS!BootY = 0
    RS!Shop = 0
    RS!Indoors = 0
    'Tiles
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            MapString = MapString & "0" & "-"
            MapString = MapString & "0" & "-"
            MapString = MapString & "0" & "-"
            MapString = MapString & "0" & "-"
            MapString = MapString & "0" & "-"
            MapString = MapString & "0" & "-"
            MapString = MapString & "0" & "-"
            MapString = MapString & "0" & "|"
        Next Y
    Next X
    
    RS!Tiles = MapString
    'NPCs
    For X = 1 To MAX_MAP_NPCS
        If X <> MAX_MAP_NPCS Then
            NPCString = NPCString & "0" & "-"
        Else
            NPCString = NPCString & "0"
        End If
    Next X
    RS!NPCs = NPCString
    
    RS.Update
Next I
RS.Close
Set RS = Nothing
Conn.Close
Set Conn = Nothing
End Sub

Public Function GetItemAssigned(ItemNum As Long) As String
'This will return the Name of the character the item is assigned to.
Dim FKey As Long
Dim iName As String

FKey = Item(ItemNum).Assigned

'If assigned to nobody
If FKey = 0 Then
    GetItemAssigned = " "
    Exit Function
End If

'If character no longer exists, then we need to update the item in question and send negative
If FindCharExistBYFKey(FKey) = False Then
    Item(ItemNum).Assigned = 0
    Call SaveItem(ItemNum)
    GetItemAssigned = " "
    Exit Function
End If

'Item must exist so grab the Name, and return it.
iName = GetItemAssignedName(FKey)
GetItemAssigned = iName
End Function

Public Function GetItemAssignedName(FKey As Long) As String
Dim RS As ADODB.Recordset
Dim iName As String

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM characters WHERE `FKey`=" & FKey & ";", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
'Load MOTD
iName = CStr(RS("Name"))
RS.Close
Set RS = Nothing

GetItemAssignedName = iName
End Function

Public Function FindCharExistBYFKey(ByVal FKey As Long) As Boolean
'NOTE: To find a character in the database list, to see if it already exists :)
Dim cLogin As String
Dim iExist As Boolean
Dim RS As ADODB.Recordset

'Setup database stuff
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM characters WHERE `FKey`=" & FKey & ";", Conn_Server, adOpenStatic, adLockReadOnly

iExist = False
If Not (RS.BOF And RS.EOF) Then
    RS.MoveFirst
    iExist = True
Else
    iExist = False
End If
RS.Close
Set RS = Nothing
FindCharExistBYFKey = iExist
End Function

Public Function GetCharFKeyByName(Name As String) As Long
Dim RS As ADODB.Recordset
Dim I As Long

'Setup database stuff
Set RS = New ADODB.Recordset
Debug.Print "Name: " & Name
RS.Open "SELECT * FROM characters WHERE `Name`='" & Trim(Name) & "';", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst
I = CLng(RS("FKey"))
RS.Close
Set RS = Nothing
GetCharFKeyByName = I
End Function

Public Function GetStaffListString() As String
Dim RS As ADODB.Recordset
Dim Packet As String
Dim I As Long
Dim RecCount As Long

RecCount = GetRecordCount("characters", "`Access` > 0")
Debug.Print "RecCount: " & RecCount
Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM characters WHERE `Access` > 0;", Conn_Server, adOpenStatic, adLockReadOnly
RS.MoveFirst

Packet = RecCount & SEP_CHAR
For I = 1 To RecCount
    Packet = Packet & Trim(CStr(RS("Name"))) & SEP_CHAR
    If I <> RecCount Then
        RS.MoveNext
    End If
Next I
RS.Close
Set RS = Nothing
Packet = Packet & END_CHAR
GetStaffListString = Packet
End Function

Public Function GetRecordCount(dbTable As String, Optional WH As String = "0") As Long
'Setup database stuff
Dim RS As ADODB.Recordset
Dim I As Long
Dim SQL As String
Set RS = New ADODB.Recordset
If WH = "0" Then
    SQL = "SELECT * FROM " & dbTable & ";"
Else
    SQL = "SELECT * FROM " & dbTable & " WHERE " & WH & ";"
End If
RS.Open SQL, Conn_Client, adOpenStatic, adLockReadOnly
RS.MoveFirst
I = RS.RecordCount
RS.Close
Set RS = Nothing
GetRecordCount = I
End Function

