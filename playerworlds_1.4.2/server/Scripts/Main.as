'*********************************************************************
'* This is a sample Main.as script for Player Worlds. You may modify *
'* this to suit the needs of your server.                            *
'*********************************************************************

Sub Init()
'***********************************************************
'* This event is fired when the server has started.        *
'***********************************************************

'**********************************
'Set your server variables below. *
'**********************************

'Set your Server name.
Call SetServerName("Generic PW Server")

'Initialize the system tray icon, with the name you wish.     
'Remember to set your server name BEFORE initializing the Tray!
'You should no longer pass a name to the InitTray Sub, it
'automatically picks up the Server Name from above.
Call InitTray()

'Set your Server port.
Call SetServerPort(7234)
'Set the maximum maps.
Call SetMaxMaps(50)
'Set the maximum players.
Call SetMaxPlayers(70)
'Set the start Map, X and Y values players will start.
Call SetDefaultPosition(1,10,10)
'Set the MOTD
Call SetMOTD("This is a sample MOTD")
'Set the Auto Ban on all hacking attempts (True of False)
Call SetAutoBan(False)

End Sub

Sub Destroy()
'***********************************************************
'* This event is fired when the server has ended  .        *
'***********************************************************

'***************************************************
'Do not modify anything below this line. -Shannara *
'***************************************************

'Destroy the system tray icon.
Call DestroyTray()
'Save all players online, before completely shutting down the server.
Call SaveAllPlayersOnline()
'Clear all maps.
Call ClearMaps()
'Clear all map items.
Call ClearMapItems()
'Clear all map NPCs.
Call ClearMapNPCs()
'Clear the NPC cache.
Call ClearNPCs()
'Clear Item cache.
Call ClearItems()
'Clear shop cache.
Call ClearShops()

End Sub

Sub OnReload(Player)
'***************************************************************
'* This event is fired when a player has reloaded this script. *
'***************************************************************


End Sub

Sub JoinGame(Player)
'***********************************************************
'* This event is fired when a player has joined the game.  *
'***********************************************************

'Send the MOTD.
Call SendMOTD(Player)
'send the welcome message.
Call PlayerMessage(Player, "Welcome to " & GetServerName & "!", 9)
Call PlayerMessage(Player, "Type /help for help on comands.  Use arrow keys to move, hold down shift to run, and use ctrl to attack.", 3)
'Send them a list of players online.
Call SendWhosOnline(Player)
'Tell everybody that player has joined.
If GetPlayerAccess(Player) <= 1 Then
   Call GlobalMessage(GetPlayerName(Player) & " has joined " & GetServerName & "!")
Else
   Call GlobalMessage(GetPlayerName(Player) & " has joined " & GetServerName & "!")
End If
End Sub      

Sub LeftGame(Player)
'***********************************************************
'* This event is fired when a player has left the game.    *
'***********************************************************  
'Tell everybody that player has joined.
If GetPlayerAccess(Player) <= 1 Then
   Call GlobalMessage(GetPlayerName(Player) & " has left " & GetServerName & "!")
Else
   Call GlobalMessage(GetPlayerName(Player) & " has left " & GetServerName & "!")
End If
End Sub

Sub OnLevelUp(Player)
Dim I

Call SetPlayerLevel(Player, GetPlayerLevel(Player) + 1)
'Get the amount of skill points to add
I = Int(GetPlayerSpeed(Player) / 10)
If I < 1 Then I = 1
If I > 3 Then I = 3

Call SetPlayerPoints(Player, GetPlayerPoints(Player) + I)
Call SetPlayerExp(Player, 0)
'Send a global message to every single person to let them know!
Call GlobalMessage(GetPlayerName(Player) & " has gained a level!", 6)
Call PlayerMessage(Player, "You have gained a level!  You now have " & GetPlayerPoints(Player) & " stat points to distribute.", 9)

End Sub           

Sub OnTime(tTime)   
'***************************************************************
'* This event is fired when the appropriate time has happened. *
'***************************************************************
Select Case tTime
	Case 0 'second ** Currently not supported **
	Case 1 'minute
	Case 2 'hour
	Call AdminMessage("The current time is " & Time & ".", 15)
	Case 3 'day
	Call AdminMessage("The current date is " & Date & ".", 15)
	Case 4 'week ** Currently not supported **
	Case 5 'month ** Currently not supported **
	Case 6 'year ** Currently not supported **
End Select
End Sub

