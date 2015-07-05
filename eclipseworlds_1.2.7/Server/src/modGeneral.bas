Attribute VB_Name = "modGeneral"
Option Explicit

' Used for the 64-bit timer
Private GetSystemTimeOffset As Currency
Private Declare Sub GetSystemTime Lib "Kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Sub Main()
    ' Make sure the application isn't already running
    If App.PrevInstance = True Then
        MsgBox "This application is already running!"
        End
    End If
    
    Call InitServer
End Sub

Sub InitServer()
    Dim i As Long
    Dim F As Long
    Dim Time1 As Long
    Dim Time2 As Long
    
    ' Set the high-resolution timer
    timeBeginPeriod 1
    
    ' This MUST be called before any timeGetTime calls because it states what the values of timeGetTime will be
    InitTimeGetTime
    
    Call InitMessages
    
    With frmServer
        Time1 = timeGetTime
        .Show
        
        ' Check if the directory is there, if its not make it
        ChkDir App.path & "\", "Data"
        ChkDir App.path & "\Data\", "logs"
        ChkDir App.path & "\Data\", "accounts"
        ChkDir App.path & "\Data\", "items"
        ChkDir App.path & "\Data\", "maps"
        ChkDir App.path & "\Data\", "npcs"
        ChkDir App.path & "\Data\", "resources"
        ChkDir App.path & "\Data\", "shops"
        ChkDir App.path & "\Data\", "spells"
        ChkDir App.path & "\Data\", "bans"
        ChkDir App.path & "\Data\", "titles"
        ChkDir App.path & "\Data\", "animations"
        ChkDir App.path & "\Data\", "morals"
        ChkDir App.path & "\Data\", "classes"
        ChkDir App.path & "\Data\", "guilds"
        ChkDir App.path & "\Data\", "emoticons"
        ChkDir App.path & "\Data\", "quests"
        
        ' Set quote character
        vbQuote = ChrW$(34)
        
        ' Load/Save options
        InitOptions
        
        ' Save options for server
        SaveOptions
        
        redimData
        
        ' Get the listening socket ready to go
        .Socket(0).RemoteHost = .Socket(0).LocalIP
        .Socket(0).LocalPort = Options.Port
        
        ' Init all the player sockets
        Call SetStatus("Initializing player array...")
    
        For i = 1 To MAX_PLAYERS
            Call ClearAccount(i)
            Load .Socket(i)
        Next
    
        ' Serves as a constructor
        Call ClearGameData
        Call LoadGameData
        Call SetStatus("Spawning map items...")
        Call SpawnAllMapsItems
        Call SetStatus("Spawning global events...")
        Call SpawnAllMapGlobalEvents
        Call SetStatus("Spawning map npcs...")
        Call SpawnAllMapNPCS
        Call SetStatus("Creating map cache...")
        Call CreateFullMapCache
        Call SetStatus("Loading system tray...")
        Call LoadSystemTray
    
        ' Start listening
        .Socket(0).Listen
        Call UpdateCaption
        Time2 = timeGetTime
        
        ' Load the news
        .txtNews.Text = Options.News
        
        ' Enable all of the disabled buttons
        .cmdExit.Enabled = True
        .cmdReloadAnimations.Enabled = True
        .cmdReloadClasses.Enabled = True
        .cmdReloadItems.Enabled = True
        .cmdReloadMaps.Enabled = True
        .cmdReloadNPCs.Enabled = True
        .cmdReloadGuilds.Enabled = True
        .cmdReloadResources.Enabled = True
        .cmdReloadEmoticons.Enabled = True
        .cmdReloadShops.Enabled = True
        .CmdReloadSpells.Enabled = True
        .cmdReloadQuests.Enabled = True
        .chkServerLog.Enabled = True
        .cmdShutDown.Enabled = True
        .cmdReloadOptions.Enabled = True
        .cmdReloadTitles.Enabled = True
        .cmdReloadBans.Enabled = True
        .cmdReloadAll.Enabled = True
        .cmdReloadMorals.Enabled = True
        .cmdLoadNews.Enabled = True
        .cmdSaveNews.Enabled = True
        .txtNews.Enabled = True
        .lblExpRate.Enabled = True
        .txtExpRate.Enabled = True
        .cmdSet.Enabled = True
        '.cmdEditPlayer.Enabled = True
        .cmdSavePlayers.Enabled = True
        
        ' Set the experience modifier to 1
        EXP_RATE = 1
        
        ' Check if it should log or not
        If Not .chkServerLog.Value Then
            ServerLog = True
        Else
            ServerLog = False
        End If
        
        Call SetStatus("Initialization complete. Server loaded in " & Time2 - Time1 & "ms.")
        
        ' Reset shutdown value
        IsShuttingDown = False
    
        ' Starts the server loop
        ServerLoop
    End With
End Sub

Sub DestroyServer()
    Dim i As Long
    
    ServerOnline = False
    
    Call SetStatus("Destroying system tray...")
    Call DestroySystemTray
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Unloading sockets...")

    For i = 1 To Player_HighIndex
        tempplayer(i).HasLogged = True
        Call CloseSocket(i)
        Unload frmServer.Socket(i)
    Next
    End
End Sub

Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing classes...")
    Call ClearClasses
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNPCs
    Call SetStatus("Clearing npcs...")
    Call ClearNPCs
    Call SetStatus("Clearing resources...")
    Call ClearResources
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing animations...")
    Call ClearAnimations
    Call SetStatus("Clearing guilds...")
    Call ClearGuilds
    Call SetStatus("Clearing bans...")
    Call ClearBans
    Call SetStatus("Clearing titles...")
    Call ClearTitles
    Call SetStatus("Clearing morals...")
    Call ClearMorals
    Call SetStatus("Clearing emoticons...")
    Call ClearEmoticons
    Call SetStatus("Clearing quests...")
    Call ClearQuests
    Call SetStatus("Clearing player quests...")
    Call ClearPlayerQuests
End Sub

Private Sub LoadGameData()
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNPCs
    Call SetStatus("Loading resources...")
    Call LoadResources
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading animations...")
    Call LoadAnimations
    Call SetStatus("Loading guilds...")
    Call LoadGuilds
    Call SetStatus("Loading bans...")
    Call LoadBans
    Call SetStatus("Loading titles...")
    Call LoadTitles
    Call SetStatus("Loading morals...")
    Call LoadMorals
    Call SetStatus("Loading emoticons...")
    Call LoadEmoticons
    Call SetStatus("Loading switches...")
    Call LoadSwitches
    Call SetStatus("Loading variables...")
    Call LoadVariables
    Call SetStatus("Loading quests...")
    Call LoadQuests
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

Public Sub InitTimeGetTime()
'*****************************************************************
' Gets the offset time for the timer so we can start at 0 instead of
' the returned system time, allowing us to not have a time roll-over until
' the program is running for 25 days
'*****************************************************************
    ' Get the initial time
    GetSystemTime GetSystemTimeOffset
End Sub

Public Function timeGetTime() As Long
'*****************************************************************
' Grabs the time from the 64-bit system timer and returns it in 32-bit
' after calculating it with the offset - allows us to have the
' "no roll-over" advantage of 64-bit timers with the RAM usage of 32-bit
' though we limit things slightly, so the rollover still happens, but after 25 days
'*****************************************************************
Dim CurrentTime As Currency
    ' Grab the current time (we have to pass a variable ByRef instead of a function return like the other timers)
    GetSystemTime CurrentTime
    
    ' Calculate the difference between the 64-bit times, return as a 32-bit time
    timeGetTime = CurrentTime - GetSystemTimeOffset
End Function

' Used for checking validity of names
Function IsNameLegal(ByVal sInput As Integer) As Boolean
    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        IsNameLegal = True
    End If
End Function
