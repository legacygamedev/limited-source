Attribute VB_Name = "modDatabase"
Option Explicit

Public Function ReadIniValue(INIpath As String, Key As String, Variable As String) As String
Dim NF As Integer
Dim Temp As String
Dim LcaseTemp As String
Dim ReadyToRead As Boolean
    
AssignVariables:
        NF = FreeFile
        ReadIniValue = vbNullString
        Key = "[" & LCase$(Key) & "]"
        Variable = LCase$(Variable)
    
EnsureFileExists:
    Open INIpath For Binary As NF
    Close NF
    SetAttr INIpath, vbArchive
    
LoadFile:
    Open INIpath For Input As NF
    While Not EOF(NF)
    Line Input #NF, Temp
    LcaseTemp = LCase$(Temp)
    If InStr(LcaseTemp, "[") <> 0 Then ReadyToRead = False
    If LcaseTemp = Key Then ReadyToRead = True
    If InStr(LcaseTemp, "[") = 0 And ReadyToRead = True Then
        If InStr(LcaseTemp, Variable & "=") = 1 Then
            ReadIniValue = Mid$(Temp, 1 + Len(Variable & "="))
            Close NF: Exit Function
            End If
        End If
    Wend
    Close NF
End Function

Public Function WriteIniValue(INIpath As String, PutKey As String, PutVariable As String, PutValue As String)
Dim Temp As String
Dim LcaseTemp As String
Dim ReadKey As String
Dim ReadVariable As String
Dim LOKEY As Integer
Dim HIKEY As Integer
Dim KeyLen As Integer
Dim VAR As Integer
Dim VARENDOFLINE As Integer
Dim NF As Integer

AssignVariables:
    NF = FreeFile
    ReadKey = vbCrLf & "[" & LCase$(PutKey) & "]" & Chr$(13)
    KeyLen = Len(ReadKey)
    ReadVariable = Chr$(10) & LCase$(PutVariable) & "="
        
EnsureFileExists:
    Open INIpath For Binary As NF
    Close NF
    SetAttr INIpath, vbArchive
    
LoadFile:
    Open INIpath For Input As NF
    Temp = Input$(LOF(NF), NF)
    Temp = vbCrLf & Temp & "[]"
    Close NF
    LcaseTemp = LCase$(Temp)
    
LogicMenu:
    LOKEY = InStr(LcaseTemp, ReadKey)
    If LOKEY = 0 Then GoTo AddKey:
    HIKEY = InStr(LOKEY + KeyLen, LcaseTemp, "[")
    VAR = InStr(LOKEY, LcaseTemp, ReadVariable)
    If VAR > HIKEY Or VAR < LOKEY Then GoTo AddVariable:
    GoTo RenewVariable:
    
AddKey:
        Temp = Left$(Temp, Len(Temp) - 2)
        Temp = Temp & vbCrLf & vbCrLf & "[" & PutKey & "]" & vbCrLf & PutVariable & "=" & PutValue
        GoTo TrimFinalString:
        
AddVariable:
        Temp = Left$(Temp, Len(Temp) - 2)
        Temp = Left$(Temp, LOKEY + KeyLen) & PutVariable & "=" & PutValue & vbCrLf & Mid$(Temp, LOKEY + KeyLen + 1)
        GoTo TrimFinalString:
        
RenewVariable:
        Temp = Left$(Temp, Len(Temp) - 2)
        VARENDOFLINE = InStr(VAR, Temp, Chr$(13))
        Temp = Left$(Temp, VAR) & PutVariable & "=" & PutValue & Mid$(Temp, VARENDOFLINE)
        GoTo TrimFinalString:

TrimFinalString:
        Temp = Mid$(Temp, 2)
        Do Until InStr(Temp, vbCrLf & vbCrLf & vbCrLf) = 0
        Temp = Replace(Temp, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
        Loop
    
        Do Until Right$(Temp, 1) > Chr$(13)
        Temp = Left$(Temp, Len(Temp) - 1)
        Loop
    
        Do Until Left$(Temp, 1) > Chr$(13)
        Temp = Mid$(Temp, 2)
        Loop
    
OutputAmendedINIFile:
        Open INIpath For Output As NF
        Print #NF, Temp
        Close NF
    
End Function

Function FileExist(ByVal FileName As String, Optional ByVal RAW As Boolean = False) As Boolean
    FileExist = True
    If RAW Then
        If Dir$(FileName) = vbNullString Then FileExist = False
    Else
        If Dir$(App.Path & "\" & FileName) = vbNullString Then FileExist = False
    End If
End Function

Sub AddLog(ByVal Text As String)
Dim FileName As String
Dim f As Long

    If Trim$(Command) = "-debug" Then
        
        FileName = App.Path & "\debug.txt"
    
        If Not FileExist("debug.txt") Then
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

Sub SaveLocalMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long
Dim i As Long
Dim X As Long
Dim Y As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
            
    If FileExist("\maps\map" & MapNum & ".dat") Then Kill FileName
    
    f = FreeFile
    Open FileName For Binary As #f
        With SaveMap
            Put #f, , .Name
            Put #f, , .Revision
            Put #f, , .Moral
            Put #f, , .Up
            Put #f, , .Down
            Put #f, , .Left
            Put #f, , .Right
            Put #f, , .Music
            Put #f, , .BootMap
            Put #f, , .BootX
            Put #f, , .BootY
            Put #f, , .TileSet
            Put #f, , .MaxX
            Put #f, , .MaxY
    
            For X = 0 To .MaxX
                For Y = 0 To .MaxY
                    Put #f, , .Tile(X, Y)
                Next
            Next
            
            For i = 1 To MAX_MOBS
                Put #f, , .Mobs(i).NpcCount
                If .Mobs(i).NpcCount > 0 Then
                    For Y = 1 To .Mobs(i).NpcCount
                        Put #f, , .Mobs(i).Npc(Y)
                    Next
                End If
            Next
        End With
    Close #f
End Sub

Sub LoadMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long
Dim X As Long
Dim Y As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
    
    ClearMaps
    
    ' check if the file exists - if not create a blank one
    If FileExist("\maps\map" & MapNum & ".dat") Then
        f = FreeFile
        Open FileName For Binary As #f
            With SaveMap
                Get #f, , .Name
                Get #f, , .Revision
                Get #f, , .Moral
                Get #f, , .Up
                Get #f, , .Down
                Get #f, , .Left
                Get #f, , .Right
                Get #f, , .Music
                Get #f, , .BootMap
                Get #f, , .BootX
                Get #f, , .BootY
                Get #f, , .TileSet
                Get #f, , .MaxX
                Get #f, , .MaxY
                
                ' have to set the tile()
                ReDim .Tile(0 To .MaxX, 0 To .MaxY) As TileRec
                
                For X = 0 To .MaxX
                    For Y = 0 To .MaxY
                        Get #f, , .Tile(X, Y)
                    Next
                Next
                
                For X = 1 To MAX_MOBS
                    Get #f, , .Mobs(X).NpcCount
                    ReDim .Mobs(X).Npc(.Mobs(X).NpcCount)
                    
                    If .Mobs(X).NpcCount > 0 Then
                        For Y = 1 To .Mobs(X).NpcCount
                            Get #f, , .Mobs(X).Npc(Y)
                        Next
                    End If
                Next
            End With
        Close #f
'    Else
'        SaveLocalMap MapNum
    End If
    
    Map = SaveMap
    
    UpdateMapNpcCount
    ClearMapNpcs
    ClearTempTile
End Sub

Function GetMapRevision(ByVal MapNum As Long) As Long
Dim FileName As String
Dim f As Long
Dim TmpMap As MapRec

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , TmpMap
    Close #f
        
    GetMapRevision = TmpMap.Revision
End Function

Public Sub CompressFile(ByRef SourceFile As String, ByRef DestFile As String, Optional ByVal KillSource As Boolean = False)
Dim Buffer As clsBuffer
Dim Size As Long
Dim Temp() As Byte
Dim f As Long

    ' check if the file even exists - exit if not
    If Not FileExist(SourceFile, True) Then Exit Sub
    
    Set Buffer = New clsBuffer
    
    ' set our temp() to the size of the file
    ReDim Temp(FileLen(SourceFile) - 1)
   
    f = FreeFile
    Open SourceFile For Binary As #f
        Get #f, , Temp()
    Close #f
    
    Buffer.WriteBytes Temp()
    Buffer.CompressBuffer

    ' check if your file exists, if so kill it
    If FileExist(DestFile, True) Then Kill DestFile

    f = FreeFile
    Open DestFile For Binary As #f
        Put #f, , Buffer.ToArray()
    Close #f

    If KillSource Then Kill SourceFile
    
    Set Buffer = Nothing
End Sub

Public Sub DecompressFile(ByRef SourceFile As String, ByRef DestFile As String, Optional ByVal KillSource As Boolean = True)
Dim Buffer As clsBuffer
Dim Temp() As Byte
Dim FileName As String
Dim f As Long
Dim X As Long
Dim Y As Long

    ' check if the file even exists - exit if not
    If Not FileExist(SourceFile, True) Then Exit Sub
    
    Set Buffer = New clsBuffer
    
    ReDim Temp(FileLen(SourceFile) - 1)
   
    f = FreeFile
    Open SourceFile For Binary As #f
        Get #f, , Temp()
    Close #f
    
    Buffer.WriteBytes Temp()
    Buffer.DecompressBuffer

    ' check if your file exists, if so kill it
    If KillSource Then
        If FileExist(DestFile, True) Then Kill DestFile
    End If
    
    f = FreeFile
    Open DestFile For Binary As #f
        Put #f, , Buffer.ToArray()
    Close #f
                
    Set Buffer = Nothing
End Sub

' *****************************************************************
' *** Below get / set UDT data with byte arrays and copy memory ***
' *****************************************************************

'
' Animations
'
Public Function Get_AnimationData(ByRef AnimationNum As Long) As Byte()
Dim AnimationData() As Byte
    ReDim AnimationData(0 To AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Get_AnimationData = AnimationData
End Function

Public Sub Set_AnimationData(ByRef AnimationNum As Long, ByRef AnimationData() As Byte)
    CopyMemory ByVal VarPtr(Animation(AnimationNum)), ByVal VarPtr(AnimationData(0)), AnimationSize
End Sub

'
' Emoticons
'
Public Function Get_EmoticonData(ByRef EmoticonNum As Long) As Byte()
Dim EmoticonData() As Byte
    ReDim EmoticonData(0 To EmoticonSize - 1)
    CopyMemory EmoticonData(0), ByVal VarPtr(Emoticons(EmoticonNum)), EmoticonSize
    Get_EmoticonData = EmoticonData
End Function

Public Sub Set_EmoticonData(ByRef EmoticonNum As Long, ByRef EmoticonData() As Byte)
    CopyMemory ByVal VarPtr(Emoticons(EmoticonNum)), ByVal VarPtr(EmoticonData(0)), EmoticonSize
End Sub

'
' Items
'
Public Function Get_ItemData(ByRef ItemNum As Long) As Byte()
Dim ItemData() As Byte
    ReDim ItemData(0 To ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Get_ItemData = ItemData
End Function

Public Sub Set_ItemData(ByRef ItemNum As Long, ByRef ItemData() As Byte)
    CopyMemory ByVal VarPtr(Item(ItemNum)), ByVal VarPtr(ItemData(0)), ItemSize
End Sub

'
' Npcs
'
Public Function Get_NpcData(ByRef NpcNum As Long) As Byte()
Dim NpcData() As Byte
    ReDim NpcData(0 To NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(Npc(NpcNum)), NpcSize
    Get_NpcData = NpcData
End Function

Public Sub Set_NpcData(ByRef NpcNum As Long, ByRef NpcData() As Byte)
    CopyMemory ByVal VarPtr(Npc(NpcNum)), ByVal VarPtr(NpcData(0)), NpcSize
End Sub

'
' Shops
'
Public Function Get_ShopData(ByRef ShopNum As Long) As Byte()
Dim ShopData() As Byte
    ReDim ShopData(0 To ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    Get_ShopData = ShopData
End Function

Public Sub Set_ShopData(ByRef ShopNum As Long, ByRef ShopData() As Byte)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize
End Sub

'
' Spells
'
Public Function Get_SpellData(ByRef SpellNum As Long) As Byte()
Dim SpellData() As Byte
    ReDim SpellData(0 To SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    Get_SpellData = SpellData
End Function

Public Sub Set_SpellData(ByRef SpellNum As Long, ByRef SpellData() As Byte)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
End Sub

