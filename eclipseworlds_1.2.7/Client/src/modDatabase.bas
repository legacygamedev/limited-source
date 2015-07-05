Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Private Const LOCALE_USER_DEFAULT& = &H400
    Private Const LOCALE_SDECIMAL& = &HE
    Private Const LOCALE_STHOUSAND& = &HF
    Private Declare Function GetLocaleInfo& Lib "kernel32" Alias _
        "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
        ByVal lpLCData As String, ByVal cchData As Long)
        
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
    ByVal lpVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long) As Long
    
Private Function DecimalSeparator() As String
      Dim R As Long, S As String
      S = String$(10, "a")
      R = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, S, 10)
      DecimalSeparator = Left$(S, R)
End Function

Public Sub HandleError(ByVal ProcName As String, ByVal ContName As String, ByVal ErNumber, ByVal ErDesc, ByVal ErSource, ByVal ErHelpContext)
    Dim FileName As String, F As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ChkDir(App.Path & "\logs\", Month(Now) & "-" & Day(Now) & "-" & year(Now))
    FileName = App.Path & "\logs\" & Month(Now) & "-" & Day(Now) & "-" & year(Now) & "\Errors.txt"
    
    If Not FileExist(FileName, True) Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
    
    Open FileName For Append As #1
        Print #1, "The following error occured at '" & ProcName & "' In '" & ContName & "'."
        Print #1, "Run-time error '" & ErNumber & "': " & ErDesc & "."
        Print #1, ""
    Close #1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If LCase$(Dir$(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not RAW Then
        If Len(Dir$(App.Path & FileName)) > 0 Then
            FileExist = True
        End If
    Else
        If Len(Dir$(FileName)) > 0 Then
            FileExist = True
        End If
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function
Private Function InternationalizeDoubles(Value As String) As String
    InternationalizeDoubles = Value
    Dim I As Long, B() As Byte, dotsCounter As Long, commasCounter As Long, test As Double
    B = Value
    For I = 0 To UBound(B) Step 2
        If B(I) = 44 Then
            commasCounter = commasCounter + 1
            Mid$(Value, I / 2 + 1, 1) = DecimalSeparator
        ElseIf B(I) = 46 Then
            dotsCounter = dotsCounter + 1
            Mid$(Value, I / 2 + 1, 1) = DecimalSeparator
        ElseIf B(I) >= 48 And B(I) <= 57 Then
        
        Else
            Exit Function
        End If
    Next I
    If (commasCounter <> 0 And dotsCounter <> 0) Or (commasCounter > 1) Or (dotsCounter > 1) Then
        Exit Function
    End If
    test = CDbl(Value)
    InternationalizeDoubles = test
End Function

' Gets a string from a text File
Public Function GetVar(file As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default Value if not found
    Dim retrivedValue As String, test As Boolean
        ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), file)
    GetVar = RTrim$(sSpaces)
    retrivedValue = Left$(GetVar, Len(GetVar) - 1)
    If InStr(retrivedValue, ",") <> 0 Or InStr(retrivedValue, ".") <> 0 Then
        retrivedValue = InternationalizeDoubles(retrivedValue)
    End If

    GetVar = retrivedValue
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

' Writes a variable to a text File
Public Sub PutVar(file As String, Header As String, Var As String, Value As String)
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call WritePrivateProfileString$(Header, Var, Value, file)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub SaveOptions()
    Dim FileName As String

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    FileName = App.Path & "\data files\config.ini"
    
    Call PutVar(FileName, "Options", "Username", Trim$(Options.UserName))
    Call PutVar(FileName, "Options", "SaveUsername", Trim$(Options.SaveUsername))
    Call PutVar(FileName, "Options", "IP", Trim$(Options.IP))
    Call PutVar(FileName, "Options", "Port", Trim$(Options.Port))
    Call PutVar(FileName, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(FileName, "Options", "Music", Trim$(Options.Music))
    Call PutVar(FileName, "Options", "Sound", Trim$(Options.Sound))
    Call PutVar(FileName, "Options", "WASD", Trim$(Options.WASD))
    Call PutVar(FileName, "Options", "Level", Trim$(Options.Levels))
    Call PutVar(FileName, "Options", "Guilds", Trim$(Options.Guilds))
    Call PutVar(FileName, "Options", "PlayerVitals", Trim$(Options.PlayerVitals))
    Call PutVar(FileName, "Options", "NPCVitals", Trim$(Options.NPCVitals))
    Call PutVar(FileName, "Options", "Titles", Trim$(Options.Titles))
    Call PutVar(FileName, "Options", "BattleMusic", Trim$(Options.BattleMusic))
    Call PutVar(FileName, "Options", "Mouse", Trim$(Options.Mouse))
    Call PutVar(FileName, "Options", "Debug", Trim$(Options.Debug))
    Call PutVar(FileName, "Options", "SwearFilter", Trim$(Options.SwearFilter))
    Call PutVar(FileName, "Options", "Weather", Trim$(Options.Weather))
    Call PutVar(FileName, "Options", "AutoTile", Trim$(Options.Autotile))
    Call PutVar(FileName, "Options", "Blood", Trim$(Options.Blood))
    Call PutVar(FileName, "Options", "MusicVolume", Trim$(Options.MusicVolume))
    Call PutVar(FileName, "Options", "SoundVolume", Trim$(Options.SoundVolume))
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub LoadOptionValues()
    Dim FileName As String
    
    FileName = App.Path & "\data files\config.ini"

    ' Load options
    If GetVar(FileName, "Options", "Username") = "" Then
        Options.UserName = vbNullString
        Call PutVar(FileName, "Options", "Username", Trim$(Options.UserName))
    Else
        Options.UserName = GetVar(FileName, "Options", "Username")
    End If
    
    If GetVar(FileName, "Options", "SaveUsername") = "" Then
        Options.SaveUsername = "1"
        Call PutVar(FileName, "Options", "SaveUsername", Trim$(Options.SaveUsername))
    Else
        Options.SaveUsername = GetVar(FileName, "Options", "SaveUsername")
    End If
    
    If GetVar(FileName, "Options", "IP") = "" Then
        Options.IP = "127.0.0.1"
        Call PutVar(FileName, "Options", "IP", Trim$(Options.IP))
    Else
        Options.IP = GetVar(FileName, "Options", "IP")
    End If
    
    If GetVar(FileName, "Options", "Port") = "" Then
        Options.Port = "7001"
        Call PutVar(FileName, "Options", "Port", Trim$(Options.Port))
    Else
        Options.Port = GetVar(FileName, "Options", "Port")
    End If
    
    If GetVar(FileName, "Options", "MenuMusic") = "" Then
        Options.MenuMusic = "Victoriam Speramus"
        Call PutVar(FileName, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Else
        Options.MenuMusic = GetVar(FileName, "Options", "MenuMusic")
    End If
    
    If GetVar(FileName, "Options", "Music") = "" Then
        Options.Music = "1"
        Call PutVar(FileName, "Options", "Music", Trim$(Options.Music))
    Else
        Options.Music = GetVar(FileName, "Options", "Music")
    End If
    
    If GetVar(FileName, "Options", "Sound") = "" Then
        Options.Sound = "1"
        Call PutVar(FileName, "Options", "Sound", Trim$(Options.Sound))
    Else
        Options.Sound = GetVar(FileName, "Options", "Sound")
    End If

    If GetVar(FileName, "Options", "WASD") = "" Then
        Options.WASD = "0"
        Call PutVar(FileName, "Options", "WASD", Trim$(Options.WASD))
    Else
        Options.WASD = GetVar(FileName, "Options", "WASD")
    End If
    
    If GetVar(FileName, "Options", "Level") = "" Then
        Options.Levels = "1"
        Call PutVar(FileName, "Options", "Level", Trim$(Options.Levels))
    Else
        Options.Levels = GetVar(FileName, "Options", "Level")
    End If
    
    If GetVar(FileName, "Options", "Guilds") = "" Then
        Options.Guilds = "1"
        Call PutVar(FileName, "Options", "Guilds", Trim$(Options.Guilds))
    Else
        Options.Guilds = GetVar(FileName, "Options", "Guilds")
    End If
    
    If GetVar(FileName, "Options", "PlayerVitals") = "" Then
        Options.PlayerVitals = "1"
        Call PutVar(FileName, "Options", "PlayerVitals", Trim$(Options.PlayerVitals))
    Else
        Options.PlayerVitals = GetVar(FileName, "Options", "PlayerVitals")
    End If
    
    If GetVar(FileName, "Options", "NPCVitals") = "" Then
        Options.NPCVitals = "1"
        Call PutVar(FileName, "Options", "NPCVitals", Trim$(Options.NPCVitals))
    Else
        Options.NPCVitals = GetVar(FileName, "Options", "NPCVitals")
    End If
    
    If GetVar(FileName, "Options", "Titles") = "" Then
        Options.Titles = "1"
        Call PutVar(FileName, "Options", "Titles", Trim$(Options.Titles))
    Else
        Options.Titles = GetVar(FileName, "Options", "Titles")
    End If
    
    If GetVar(FileName, "Options", "BattleMusic") = "" Then
        Options.BattleMusic = "1"
        Call PutVar(FileName, "Options", "BattleMusic", Trim$(Options.BattleMusic))
    Else
        Options.BattleMusic = GetVar(FileName, "Options", "BattleMusic")
    End If
    
    If GetVar(FileName, "Options", "Mouse") = "" Then
        Options.Mouse = "0"
        Call PutVar(FileName, "Options", "Mouse", Trim$(Options.Mouse))
    Else
        Options.Mouse = GetVar(FileName, "Options", "Mouse")
    End If
    
    If GetVar(FileName, "Options", "Debug") = "" Then
        Options.Debug = "1"
        Call PutVar(FileName, "Options", "Debug", Trim$(Options.Debug))
    Else
        Options.Debug = GetVar(FileName, "Options", "Debug")
    End If
    
    If GetVar(FileName, "Options", "SwearFilter") = "" Then
        Options.SwearFilter = "1"
        Call PutVar(FileName, "Options", "SwearFilter", Trim$(Options.SwearFilter))
    Else
        Options.SwearFilter = GetVar(FileName, "Options", "SwearFilter")
    End If
    
    If GetVar(FileName, "Options", "Weather") = "" Then
        Options.Weather = "1"
        Call PutVar(FileName, "Options", "Weather", Trim$(Options.Weather))
    Else
        Options.Weather = GetVar(FileName, "Options", "Weather")
    End If
    
    If GetVar(FileName, "Options", "AutoTile") = "" Then
        Options.Autotile = "1"
        Call PutVar(FileName, "Options", "AutoTile", Trim$(Options.Autotile))
    Else
        Options.Autotile = GetVar(FileName, "Options", "AutoTile")
    End If
    
    If GetVar(FileName, "Options", "Blood") = "" Then
        Options.Blood = "1"
        Call PutVar(FileName, "Options", "Blood", Trim$(Options.Blood))
    Else
        Options.Blood = GetVar(FileName, "Options", "Blood")
    End If
    
    If GetVar(FileName, "Options", "MusicVolume") = "" Then
        Options.MusicVolume = InternationalizeDoubles("0.5")
        Call PutVar(FileName, "Options", "MusicVolume", Trim$(Options.MusicVolume))
    Else
        Options.MusicVolume = GetVar(FileName, "Options", "MusicVolume")
    End If
    
    If GetVar(FileName, "Options", "SoundVolume") = "" Then
        Options.SoundVolume = InternationalizeDoubles("0.8")
        Call PutVar(FileName, "Options", "SoundVolume", Trim$(Options.SoundVolume))
    Else
        Options.SoundVolume = GetVar(FileName, "Options", "SoundVolume")
    End If
    
    If GetVar(FileName, "Options", "ResolutionWidth") = "" Then
        Options.ResolutionWidth = "800"
        Call PutVar(FileName, "Options", "ResolutionWidth", Trim$(Options.ResolutionWidth))
    Else
        Options.ResolutionWidth = GetVar(FileName, "Options", "ResolutionWidth")
    End If
    
    If GetVar(FileName, "Options", "ResolutionHeight") = "" Then
        Options.ResolutionHeight = "640"
        Call PutVar(FileName, "Options", "ResolutionHeight", Trim$(Options.ResolutionHeight))
    Else
        Options.ResolutionHeight = GetVar(FileName, "Options", "ResolutionHeight")
    End If
End Sub

Public Sub LoadOptions()
    ' Load the variables in the options.ini
    Call LoadOptionValues
    
    ' Set the form items based on what the options are
    ResetOptionButtons
End Sub

Public Function TimeStamp() As String
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    TimeStamp = "[" & time & "]"
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "TimeStamp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub AddLog(ByVal text As String, ByVal LogFile As String)
    Dim FileName As String
    Dim F As Integer
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ChkDir(App.Path & "\logs\", Month(Now) & "-" & Day(Now) & "-" & year(Now))
    FileName = App.Path & "\logs\" & Month(Now) & "-" & Day(Now) & "-" & year(Now) & "\" & LogFile & ".log"

    If Not FileExist(FileName, True) Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    F = FreeFile
    
    Open FileName For Append As #F
        Print #F, TimeStamp & " - " & text
    Close #F
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "AddLog", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckTilesets()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumTileSets = 1
    
    ReDim Tex_Tileset(1)

    While FileExist(GFX_PATH & "tilesets\" & I & GFX_EXT)
        ReDim Preserve Tex_Tileset(NumTileSets)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Tileset(NumTileSets).filepath = App.Path & GFX_PATH & "tilesets\" & I & GFX_EXT
        Tex_Tileset(NumTileSets).Texture = NumTextures
        NumTileSets = NumTileSets + 1
        I = I + 1
    Wend
    
    NumTileSets = NumTileSets - 1
    
    'If NumTileSets < 1 Then Exit Sub
    
    'For i = 1 To NumTileSets
    '    LoadTexture Tex_Tileset(i)
    'Next
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckCharacters()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumCharacters = 1
    
    ReDim Tex_Character(1)
    Dim test As String
    test = Dir$(GFX_PATH & "characters\" & "*" & GFX_EXT, vbNormal)
    
    While FileExist(GFX_PATH & "characters\" & I & GFX_EXT)
        ReDim Preserve Tex_Character(NumCharacters)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Character(NumCharacters).filepath = App.Path & GFX_PATH & "characters\" & I & GFX_EXT
        Tex_Character(NumCharacters).Texture = NumTextures
        NumCharacters = NumCharacters + 1
        I = I + 1
    Wend
    
    NumCharacters = NumCharacters - 1
    
    'If NumCharacters < 1 Then Exit Sub
    
    'For i = 1 To NumCharacters
    '    LoadTexture Tex_Character(i)
    'Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckPaperdolls()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumPaperdolls = 1
    
    ReDim Tex_Paperdoll(1)

    While FileExist(GFX_PATH & "paperdolls\" & I & GFX_EXT)
        ReDim Preserve Tex_Paperdoll(NumPaperdolls)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Paperdoll(NumPaperdolls).filepath = App.Path & GFX_PATH & "paperdolls\" & I & GFX_EXT
        Tex_Paperdoll(NumPaperdolls).Texture = NumTextures
        NumPaperdolls = NumPaperdolls + 1
        I = I + 1
    Wend
    
    NumPaperdolls = NumPaperdolls - 1
    
    'If NumPaperdolls < 1 Then Exit Sub
    
    'For i = 1 To NumPaperdolls
    '    LoadTexture Tex_Paperdoll(i)
    'Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckAnimations()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumAnimations = 1
    
    ReDim Tex_Animation(1)

    While FileExist(GFX_PATH & "animations\" & I & GFX_EXT)
        ReDim Preserve Tex_Animation(NumAnimations)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Animation(NumAnimations).Texture = NumTextures
        Tex_Animation(NumAnimations).filepath = App.Path & GFX_PATH & "animations\" & I & GFX_EXT
        NumAnimations = NumAnimations + 1
        I = I + 1
    Wend
    
    NumAnimations = NumAnimations - 1
    
    'If NumAnimations < 1 Then Exit Sub

    'For i = 1 To NumAnimations
    '    LoadTexture Tex_Animation(i)
    'Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckItems()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumItems = 1
    
    ReDim Tex_Item(1)

    While FileExist(GFX_PATH & "items\" & I & GFX_EXT)
        ReDim Preserve Tex_Item(NumItems)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Item(NumItems).filepath = App.Path & GFX_PATH & "items\" & I & GFX_EXT
        Tex_Item(NumItems).Texture = NumTextures
        NumItems = NumItems + 1
        I = I + 1
    Wend
    
    NumItems = NumItems - 1
    
    'If NumItems < 1 Then Exit Sub
    
    'For i = 1 To NumItems
    '    LoadTexture Tex_Item(i)
    'Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckResources()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumResources = 1
    
    ReDim Tex_Resource(1)

    While FileExist(GFX_PATH & "resources\" & I & GFX_EXT)
        ReDim Preserve Tex_Resource(NumResources)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Resource(NumResources).filepath = App.Path & GFX_PATH & "resources\" & I & GFX_EXT
        Tex_Resource(NumResources).Texture = NumTextures
        NumResources = NumResources + 1
        I = I + 1
    Wend
    
    NumResources = NumResources - 1
    
    'If NumResources < 1 Then Exit Sub
    
    'For i = 1 To NumResources
    '    LoadTexture Tex_Resource(i)
    'Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckSpellIcons()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumSpellIcons = 1
    
    ReDim Tex_SpellIcon(1)

    While FileExist(GFX_PATH & "spellicons\" & I & GFX_EXT)
        ReDim Preserve Tex_SpellIcon(NumSpellIcons)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_SpellIcon(NumSpellIcons).filepath = App.Path & GFX_PATH & "spellicons\" & I & GFX_EXT
        Tex_SpellIcon(NumSpellIcons).Texture = NumTextures
        NumSpellIcons = NumSpellIcons + 1
        I = I + 1
    Wend

    NumSpellIcons = NumSpellIcons - 1
    
    'If NumSpellIcons < 1 Then Exit Sub
    
    'For i = 1 To NumSpellIcons
    '    LoadTexture Tex_SpellIcon(i)
    'Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckSpellIcons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckFaces()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumFaces = 1
    
    ReDim Tex_Face(1)

    While FileExist(GFX_PATH & "Faces\" & I & GFX_EXT)
        ReDim Preserve Tex_Face(NumFaces)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Face(NumFaces).filepath = App.Path & GFX_PATH & "faces\" & I & GFX_EXT
        Tex_Face(NumFaces).Texture = NumTextures
        NumFaces = NumFaces + 1
        I = I + 1
    Wend
    
    NumFaces = NumFaces - 1
     
    'If NumFaces < 1 Then Exit Sub
    
    'For i = 1 To NumFaces
    '    LoadTexture Tex_Face(i)
    'Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckFogs()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumFogs = 1
    
    ReDim Tex_Fog(1)
    
    While FileExist(GFX_PATH & "fogs\" & I & GFX_EXT)
        ReDim Preserve Tex_Fog(NumFogs)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Fog(NumFogs).filepath = App.Path & GFX_PATH & "fogs\" & I & GFX_EXT
        Tex_Fog(NumFogs).Texture = NumTextures
        NumFogs = NumFogs + 1
        I = I + 1
    Wend
    
    NumFogs = NumFogs - 1
    
    'If NumFogs < 1 Then Exit Sub
    
    'For i = 1 To NumFogs
    '    LoadTexture Tex_Fog(i)
    'Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckFogs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckPanoramas()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumPanoramas = 1
    
    ReDim Tex_Panorama(1)
    While FileExist(GFX_PATH & "Panoramas\" & I & GFX_EXT)
        ReDim Preserve Tex_Panorama(NumPanoramas)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Panorama(NumPanoramas).filepath = App.Path & GFX_PATH & "Panoramas\" & I & GFX_EXT
        Tex_Panorama(NumPanoramas).Texture = NumTextures
        NumPanoramas = NumPanoramas + 1
        I = I + 1
    Wend
    
    NumPanoramas = NumPanoramas - 1
    
    'If NumPanoramas < 1 Then Exit Sub
    
    'For i = 1 To NumPanoramas
    '    LoadTexture Tex_Panorama(i)
    'Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckPanoramas", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckEmoticons()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    I = 1
    NumEmoticons = 1
    
    ReDim Tex_Emoticon(1)

    While FileExist(GFX_PATH & "Emoticons\" & I & GFX_EXT)
        ReDim Preserve Tex_Emoticon(NumEmoticons)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Emoticon(NumEmoticons).filepath = App.Path & GFX_PATH & "Emoticons\" & I & GFX_EXT
        Tex_Emoticon(NumEmoticons).Texture = NumTextures
        NumEmoticons = NumEmoticons + 1
        I = I + 1
    Wend
    
    NumEmoticons = NumEmoticons - 1
    
    'If NumEmoticons < 1 Then Exit Sub
    
    'For i = 1 To NumEmoticons
    '    LoadTexture Tex_Emoticon(i)
    'Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckEmoticons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    Player(Index).Name = vbNullString
    Player(Index).Status = vbNullString
    Player(Index).Class = 1
    
    For I = 1 To Stats.Stat_Count - 1
        Call SetPlayerStat(Index, I, 1)
    Next
    
    For I = 1 To Skills.Skill_Count - 1
        Call SetPlayerSkill(Index, 1, I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).Sound = vbNullString
    Item(Index).Rarity = 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearItems()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearAnimation(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = vbNullString
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearAnimations()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_ANIMATIONS
        Call ClearAnimation(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearNPC(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).title = vbNullString
    NPC(Index).AttackSay = vbNullString
    NPC(Index).Music = vbNullString
    NPC(Index).Sound = vbNullString
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearNPCs()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_NPCS
        Call ClearNPC(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearNPCs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearSpell(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).Sound = vbNullString
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearSpells()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_SPELLS
        Call ClearSpell(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearShop(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearShops()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_SHOPS
        Call ClearShop(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearResource(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).FailMessage = vbNullString
    Resource(Index).Sound = vbNullString
    Exit Sub
    
ErrorHandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearResources()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_RESOURCES
        Call ClearResource(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMapItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
    MapItem(Index).PlayerName = vbNullString
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMap()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
    Map.Music = vbNullString
    Map.BGS = vbNullString
    Map.Moral = 1
    Map.MaxX = MIN_MAPX
    Map.MaxY = MIN_MAPY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    InitAutotiles
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMapItems()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMapNPC(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(MapNPC(Index)), LenB(MapNPC(Index)))
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearMapNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMapNPCs()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_MAP_NPCS
        Call ClearMapNPC(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearMapNPCs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearBans()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_BANS
        Call ClearBan(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearBans", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearBan(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ZeroMemory(ByVal VarPtr(Ban(Index)), LenB(Ban(Index)))
    Ban(Index).PlayerLogin = vbNullString
    Ban(Index).PlayerName = vbNullString
    Ban(Index).Reason = vbNullString
    Ban(Index).IP = vbNullString
    Ban(Index).HDSerial = vbNullString
    Ban(Index).time = vbNullString
    Ban(Index).By = vbNullString
    Ban(Index).Date = vbNullString
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearBan", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearTitles()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_TITLES
        Call ClearTitle(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearTitles", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearTitle(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ZeroMemory(ByVal VarPtr(title(Index)), LenB(title(Index)))
    title(Index).Name = vbNullString
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearTitle", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMoral(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Moral(Index)), LenB(Moral(Index)))
    Moral(Index).Name = vbNullString
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearMoral", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearMorals()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_MORALS
        Call ClearMoral(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearMorals", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearClass(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ZeroMemory(ByVal VarPtr(Class(Index)), LenB(Class(Index)))
    Class(Index).Name = vbNullString
    Class(Index).CombatTree = 1
    Class(Index).Map = 1
    Class(Index).Color = 15
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearClasses()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_CLASSES
        Call ClearClass(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearClasses", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearEmoticon(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Emoticon(Index)), LenB(Emoticon(Index)))
    Emoticon(Index).Command = "/"
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearEmoticon", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearEmoticons()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_EMOTICONS
        Call ClearEmoticon(I)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearEmoticons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearEvents()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_EVENTS
        Call ClearEvent(I)
    Next I
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearEvents", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearEvent(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index <= 0 Or Index > MAX_EVENTS Then Exit Sub
    
    Call ZeroMemory(ByVal VarPtr(events(Index)), LenB(events(Index)))
    events(Index).Name = vbNullString
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearEvent", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearData()
    Dim I As Long
    
    ' Clear all data.. Screw you structs
    ClearNPCs
    ClearResources
    ClearItems
    ClearShops
    ClearSpells
    ClearAnimations
    ClearTitles
    ClearClasses
    ClearEmoticons
    ClearBans
    ClearMorals
    ClearQuests
    
    For I = 1 To MAX_PLAYERS
        Call ClearPlayer(I)
    Next
End Sub

Public Sub redimDataPreserve()
    ReDim Preserve Item(MAX_ITEMS)
    ReDim Preserve Shop(MAX_SHOPS)
    ReDim Preserve Class(MAX_CLASSES)
    ReDim Preserve Animation(MAX_ANIMATIONS)
    ReDim Preserve Emoticon(MAX_EMOTICONS)
    ReDim Preserve Moral(MAX_MORALS)
    ReDim Preserve NPC(MAX_NPCS)
    ReDim Preserve Quest(MAX_QUESTS)
    ReDim Preserve Resource(MAX_RESOURCES)
    ReDim Preserve Spell(MAX_SPELLS)
    ReDim Preserve title(MAX_TITLES)
    ReDim Preserve Ban(MAX_BANS)
    ReDim Preserve ClassSelection(MAX_CLASSES)
    ReDim Preserve Item_Changed(MAX_ITEMS)
    ReDim Preserve Quest_Changed(MAX_QUESTS)
    ReDim Preserve NPC_Changed(MAX_NPCS)
    ReDim Preserve Resource_Changed(MAX_RESOURCES)
    ReDim Preserve Animation_Changed(MAX_ANIMATIONS)
    ReDim Preserve Spell_Changed(MAX_SPELLS)
    ReDim Preserve Shop_Changed(MAX_SHOPS)
    ReDim Preserve Ban_Changed(MAX_BANS)
    ReDim Preserve Title_Changed(MAX_TITLES)
    ReDim Preserve Moral_Changed(MAX_MORALS)
    ReDim Preserve Class_Changed(MAX_CLASSES)
    ReDim Preserve Emoticon_Changed(MAX_EMOTICONS)
    
    Dim I As Long, II As Long
    For I = 1 To MAX_PLAYERS
        ReDim Preserve Player(I).QuestCLI(MAX_QUESTS)
        ReDim Preserve Player(I).QuestTask(MAX_QUESTS)
        ReDim Preserve Player(I).QuestCompleted(MAX_QUESTS)
        ReDim Preserve Player(I).QuestAmount(MAX_QUESTS)
        For II = 1 To MAX_QUESTS
            ReDim Preserve Player(I).QuestAmount(II).ID(1 To MAX_NPCS)
        Next
    Next
End Sub

Public Sub redimData()
    ReDim Item(MAX_ITEMS)
    ReDim Shop(MAX_SHOPS)
    ReDim Class(MAX_CLASSES)
    ReDim Animation(MAX_ANIMATIONS)
    ReDim Emoticon(MAX_EMOTICONS)
    ReDim Moral(MAX_MORALS)
    ReDim NPC(MAX_NPCS)
    ReDim Quest(MAX_QUESTS)
    ReDim Resource(MAX_RESOURCES)
    ReDim Spell(MAX_SPELLS)
    ReDim title(MAX_TITLES)
    ReDim Ban(MAX_BANS)
    ReDim ClassSelection(MAX_CLASSES)
    ReDim Item_Changed(MAX_ITEMS)
    ReDim Quest_Changed(MAX_QUESTS)
    ReDim NPC_Changed(MAX_NPCS)
    ReDim Resource_Changed(MAX_RESOURCES)
    ReDim Animation_Changed(MAX_ANIMATIONS)
    ReDim Spell_Changed(MAX_SPELLS)
    ReDim Shop_Changed(MAX_SHOPS)
    ReDim Ban_Changed(MAX_BANS)
    ReDim Title_Changed(MAX_TITLES)
    ReDim Moral_Changed(MAX_MORALS)
    ReDim Class_Changed(MAX_CLASSES)
    ReDim Emoticon_Changed(MAX_EMOTICONS)
    
    Dim I As Long, II As Long
    For I = 1 To MAX_PLAYERS
        ReDim Player(I).QuestCLI(MAX_QUESTS)
        ReDim Player(I).QuestTask(MAX_QUESTS)
        ReDim Player(I).QuestCompleted(MAX_QUESTS)
        ReDim Player(I).QuestAmount(MAX_QUESTS)
        For II = 1 To MAX_QUESTS
            ReDim Player(I).QuestAmount(II).ID(1 To MAX_NPCS)
        Next
    Next
End Sub

