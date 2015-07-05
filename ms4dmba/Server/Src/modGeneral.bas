Attribute VB_Name = "modGeneral"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

' Get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long

Sub Main()
    Call InitServer
End Sub

Sub InitServer()
Dim i As Long
Dim F As Long

Dim time1 As Long
Dim time2 As Long

    Call InitMessages

    time1 = GetTickCount
    
    frmServer.Show
    
    ' Initialize the random-number generator
    Randomize ', seed
 
    ' Check if the directory is there, if its not make it
    
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
    
    ' used for parsing packets
    SEP_CHAR = vbNullChar ' ChrW$(0)
    END_CHAR = ChrW$(237)
    
    ' set quote character
    vbQuote = ChrW$(34) ' "
    
    ' set MOTD
    MOTD = GetVar(App.Path & "\data\motd.ini", "MOTD", "Msg")
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_PORT
        
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next
    
    ' Serves as a constructor
    Call ClearGameData

    Call LoadGameData

    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache

    Call SetStatus("Loading System Tray...")
    Call LoadSystemTray
    
        
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("accounts\charlist.txt") Then
        F = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #F
        Close #F
    End If
    
    ' Start listening
    frmServer.Socket(0).Listen
    
    Call UpdateCaption
    
    ' To allow first connection
    High_Index = 1

    time2 = GetTickCount
    
    Call SetStatus("Initialization complete. Server loaded in " & time2 - time1 & "ms.")

    ' Starts the server loop
    ServerLoop
End Sub

Sub DestroyServer()
Dim i As Long
 
    ServerOnline = False
    
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    
    Call ClearGameData
    
    Call SetStatus("Unloading sockets...")
    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next

    End
End Sub

Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTiles
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
End Sub

Private Sub LoadGameData()
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
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

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean
    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If
End Function



