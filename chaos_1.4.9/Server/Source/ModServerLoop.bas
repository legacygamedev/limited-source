Attribute VB_Name = "ModServerLoop"
Option Explicit
Public Seconds As Long
Public Minutes As Integer
Public ServerOnline As Byte
Public ServerShutDown As Byte
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ServerLoop()
Static Secs As Long
Dim I As Long
Dim N As Long
Dim Index As Long
Dim Tick As Long ' Server Loop Tick
Dim TickRandomEvent As Long ' Random Event Tick
Dim TickLogic As Long ' Game Logic Tick
Dim TickSpawnItems As Long ' Spawn Items Tick
Dim TickSavePlayers As Long ' Save Players Tick
Dim TickSaveChatLogs As Long ' Save Chat Logs Tick
Dim TickUpdatePlayerSP As Long ' Update Player SP
Dim TickSpellCast As Long
Dim TickSubscriptions As Long
Dim TickSave As Long
Dim TickShutdown As Long
Dim TickHunger As Long
Dim TickAilments As Long

ServerOnline = 1 ' Set the Server Online to begin the Server loop
frmServer.Show
Do While ServerOnline = 1
    Tick = GetTickCount
   
    If Tick > TickLogic Then
        Call ServerLogic
        TickLogic = GetTickCount + 300 ' 500 miliseconds
    End If
   
    If Tick > TickSpawnItems Then
    
    For I = 1 To MAX_PLAYERS
    If IsPlaying(I) = True Then
    
    If GetPlayerTradeskillMS(I) > 0 Then
    Call SetPlayerTradeskillMS(I, GetPlayerTradeskillMS(I) - 1)
    End If
    
    If GetPlayerPoisoned(I) > 0 Or GetPlayerDiseased(I) > 0 Then
    
    If GetPlayerAilmentMS(I) > 0 Then
    Call SetPlayerAilmentMS(I, GetPlayerAilmentMS(I) - 1)
    Else
    Call SetPlayerPoisoned(I, 0)
    Call SetPlayerAilmentMS(I, 0)
    Call PlayerMsg(I, "The effects of The Plague Wear Off !", Yellow)
    End If
    
    End If
    End If
    Next
    
        Call CheckSpawnMapItems
        Call CheckTime
        Call DoLogs
        TickSpawnItems = GetTickCount + 1000 ' 1 second
    End If
    
    If Tick > TickAilments Then
    For I = 1 To MAX_PLAYERS
    If IsPlaying(I) Then
    
    If GetPlayerPoisoned(I) > 0 Or GetPlayerDiseased(I) > 0 Then
    Call PoisonActive(I)
    TickAilments = GetTickCount + GetPlayerAilmentInterval(I)
    End If
    
    If GetPlayerDiseased(I) > 0 Then
    Call DiseaseActive(I)
    TickAilments = GetTickCount + GetPlayerAilmentInterval(I)
    End If
    
    End If
    Next
    End If
    
    If Tick > TickHunger Then
    For I = 1 To MAX_PLAYERS
    If IsPlaying(I) = True Then
    Call HungerActive(I)
    End If
    Next
    TickHunger = GetTickCount + 60000 ' 1 Minute
    End If
   
    If ServerShutDown > 0 Then
    If Tick > TickShutdown Then
    Call DoShutDown
    TickShutdown = GetTickCount + 1000
    End If
    End If
    
    If KICKIDLEPLAYERS = 1 Then
    For I = 1 To MAX_PLAYERS
    
    'If IsPlaying(i) = True Then
    'If GetTickCount > Player(i).OnlineTime + 650000 Then
    'Call LeftGame(i)
    'End If
    'End If
    
    Next
    End If
   
    If Tick > TickSave Then
    Call PlayerSaveTimer
    TickSave = GetTickCount + 60000
    End If
   
    Sleep 1 ' Sleep for 1 MiliSecond
    DoEvents
Loop

Call DestroyServer

End Sub

Sub DoLogs()
Static ChatSecs As Long
Dim SaveTime As Long

    SaveTime = 3600

    If frmServer.chkChat.Value = Unchecked Then
        ChatSecs = SaveTime
        frmServer.Label6.Caption = "Chat Log Save Disabled!"
        Exit Sub
    End If

    If ChatSecs <= 0 Then ChatSecs = SaveTime
    If ChatSecs > 60 Then
        frmServer.Label6.Caption = "Chat Log Save In " & Int(ChatSecs / 60) & " Minute(s)"
    Else
        frmServer.Label6.Caption = "Chat Log Save In " & Int(ChatSecs) & " Second(s)"
    End If
    ChatSecs = ChatSecs - 1

    If ChatSecs <= 0 Then
        Call TextAdd(frmServer.txtText(0), "Chat Logs Have Been Saved!", True)
        Call SaveLogs
        Call TextAdd(frmServer.txtText(0), "Game Time Saved !", True)
        ChatSecs = 0
    End If
End Sub

Sub DoShutDown()
Static Secs As Long

    If Secs <= 0 Then Secs = 30
    frmServer.ShutdownTime.Caption = "Shutdown: " & Secs & " Seconds"

    If Secs = 30 Then Call TextAdd(frmServer.txtText(0), "Automated Server Shutdown in " & Secs & " seconds.", True)
    If Secs = 30 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 25 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 20 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 15 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 10 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs < 6 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    End If
    Secs = Secs - 1

    If Secs <= 0 Then
        Call DestroyServer
    End If
End Sub

Sub CheckTime()
Dim AMorPM As String
Dim TempSeconds As Integer
Dim PrintSeconds As String
Dim PrintSeconds2 As String
Dim PrintMinutes As String
Dim PrintMinutes2 As String
Dim PrintHours As Integer

Seconds = Seconds + Gamespeed

If Seconds > 59 Then
    Minutes = Minutes + 1
    Seconds = Seconds - 60
End If
If Minutes > 59 Then
    Hours = Hours + 1
    Minutes = 0
End If
If Hours > 24 Then
    Hours = 1
End If

If Hours > 12 Then
    AMorPM = "PM"
    PrintHours = Hours - 12
Else
    AMorPM = "AM"
    PrintHours = Hours
End If

If Hours = 24 Then
    AMorPM = "AM"
End If

TempSeconds = Seconds

If Seconds > 9 Then
    PrintSeconds = TempSeconds
Else
    PrintSeconds = "0" & Seconds
End If

If Seconds > 50 Then
    PrintSeconds2 = "0" & 60 - TempSeconds
Else
    PrintSeconds2 = 60 - TempSeconds
End If

If Minutes > 9 Then
    PrintMinutes = Minutes
Else
    PrintMinutes = "0" & Minutes
End If

If Minutes > 50 Then
    PrintMinutes2 = "0" & 60 - Minutes
Else
    PrintMinutes2 = 60 - Minutes
End If

frmServer.Label8.Caption = "Current Time is " & PrintHours & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM

If Hours > 20 And GameTime = TIME_DAY Then
    GameTime = TIME_NIGHT
    Call SendTimeToAll
    End If
If Hours < 21 And Hours > 6 And GameTime = TIME_NIGHT Then
    GameTime = TIME_DAY
    Call SendTimeToAll
    End If
If Hours < 7 And GameTime = TIME_DAY Then
    GameTime = TIME_NIGHT
    Call SendTimeToAll
    End If
   
If Hours < 21 And Hours > 6 Then
    frmServer.Label10.Caption = "Time until night:"
    frmServer.Label11.Caption = 21 - Hours - 1 & ":" & PrintMinutes2 & ":" & PrintSeconds2
Else
    frmServer.Label10.Caption = "Time until day:"
    If Hours < 7 Then
    frmServer.Label11.Caption = 7 - Hours - 1 & ":" & PrintMinutes2 & ":" & PrintSeconds2
    Else
    frmServer.Label11.Caption = 24 - Hours - 1 + 7 & ":" & PrintMinutes2 & ":" & PrintSeconds2
    End If
End If

If Hours > 11 Then
    GameClock = Hours - 12 & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM
Else
    GameClock = Hours & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM
End If

Call SendGameClockToAll
End Sub
