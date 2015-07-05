Attribute VB_Name = "modGeneral"
Option Explicit

' ********************************************
' **               rootSource               **
' ********************************************

' Get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

' Used for frmServer.txtText
Public Const MAX_LINES As Integer = 500
Public NumLines As Long ' needed for textbox
Dim StartTime As Long

Private Sub Main()
    Call InitServer
End Sub

Private Sub InitServer()
    Dim i     As Long

    Dim time As Long

    StartTime = GetTickCount
    
    frmServer.Show
    
    Randomize ' Initialize the random-number generator
 
    ' Check if the directory is there, if its not make it
    Call CheckDir
    MOTD = Trim$(GetVar(App.Path & "\data\motd.ini", "MOTD", "Msg")) ' set MOTD
    
    Call SetStatus("Loading scripts...")
    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands
    MyScript.ReadInCode App.Path & "\data\scripts\main.vb", "\data\scripts\main.vb", MyScript.SControl
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    
    ' Set the Important Data
    MyScript.ExecuteStatement "\data\scripts\main.vb", "ServerSet"
    
    ' Reset Stuff based On Varriables
    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim MapCache(1 To MAX_MAPS) As Cache
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long
    ReDim TempTile(1 To MAX_MAPS) As TempTileRec
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Item(1 To MAX_ITEMS) As ItemRec
    ReDim Npc(1 To MAX_NPCS) As NpcRec
    
    Call UsersOnline_Start
    Call ClearGameData ' Serves as a constructor
    Call LoadGameData
    
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map Npcs...")
    Call SpawnAllMapNpcs
    
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache

    Call SetStatus("Initializing System Tray...")
    Call InitSystemTray
    
    Call SetStatus("Initializing Winsock...")
    Call InitWinsock
    Call UpdateCaption
        
    ServerOnline = True

    time = GetTickCount
    
    Call SetStatus("Initialization complete. Server loaded in " & time - StartTime & "ms.")

    ServerLoop ' Starts the server loop

End Sub

Private Sub InitWinsock()
On Error GoTo ErrorHandle
Dim i As Long

    ' Init the messages for handle data
    Call InitMessages
    
    ' Init all the player sockets
    For i = 1 To MAX_PLAYERS
        Load frmServer.Socket(i)
    Next

    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = 7000

    frmServer.Socket(0).Listen ' Start listening
    
    Exit Sub
    
ErrorHandle:

    Select Case Err
    
        Case 10048
            MsgBox "Port is already in use."
    
    End Select
    
    DestroyServer
    
End Sub

Public Sub DestroyServer()
Dim i As Long
 
    ServerOnline = False
    
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    
    Call SetStatus("Saving players online...")
    For i = 1 To TotalPlayersOnline
        Call LeftGame(PlayersOnline(i))
    Next
    
    Call ClearGameData
    
    Call SetStatus("Unloading sockets...")
    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next

    End
End Sub

Public Sub SetStatus(ByVal Status As String)

        Call TextAdd(Status)

    DoEvents
End Sub

Private Sub ClearGameData()
Dim i As Long

    Call SetStatus("Clearing data...")

    'Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTile
    
    'Call SetStatus("Clearing Npcs...")
    Call ClearNpcs
    
    'Call SetStatus("Clearing items...")
    Call ClearItems
    
    'Call SetStatus("Clearing classes...")
    Call ClearClasses
    
    'Call SetStatus("Clearing maps...")
    Call ClearMaps
    
    'Call SetStatus("Clearing map items...")
    Call ClearMapItems
    
    'Call SetStatus("Clearing map Npcs...")
    Call ClearMapNpcs
    
    'Call SetStatus("Clearing shops...")
    Call ClearShops
    
    'Call SetStatus("Clearing spells...")
    Call ClearSpells
    
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
    Next
    
End Sub

Private Sub LoadGameData()
    Call SetStatus("Loading Npcs...")
    Call LoadNpcs
    
    Call SetStatus("Loading items...")
    Call LoadClasses

    Call SetStatus("Loading classes...")
    Call LoadItems

    Call SetStatus("Loading maps...")
    Call LoadMaps

    Call SetStatus("Loading shops...")
    Call LoadShops
    
    Call SetStatus("Loading spells...")
    Call LoadSpells
End Sub

Public Sub TextAdd(Msg As String)
    
    NumLines = NumLines + 1
    
    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If
    
    frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & Msg
    
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)
    
End Sub

Private Sub CheckDir()
Dim F As Long
            
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("accounts\charlist.txt") Then
        F = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #F
        Close #F
    End If
    
    If LCase$(Dir(App.Path & "\Data\items", vbDirectory)) <> "items" Then
        Call MkDir(App.Path & "\Data\items")
    End If
    
    If LCase$(Dir(App.Path & "\Data\maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\Data\maps")
    End If
    
    If LCase$(Dir(App.Path & "\Data\npcs", vbDirectory)) <> "npcs" Then
        Call MkDir(App.Path & "\Data\npcs")
    End If
    
    If LCase$(Dir(App.Path & "\Data\shops", vbDirectory)) <> "shops" Then
        Call MkDir(App.Path & "\Data\shops")
    End If
    
    If LCase$(Dir(App.Path & "\Data\spells", vbDirectory)) <> "spells" Then
        Call MkDir(App.Path & "\Data\spells")
    End If
    
    If LCase$(Dir(App.Path & "\accounts", vbDirectory)) <> "accounts" Then
        Call MkDir(App.Path & "\accounts")
    End If

End Sub

Private Sub UsersOnline_Start()
Dim i As Integer

    For i = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (i)
        
        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If
        
        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next
    
End Sub

' Used for checking validity of names
Public Function isNameLegal(ByVal sInput As Integer) As Boolean
    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If
End Function


