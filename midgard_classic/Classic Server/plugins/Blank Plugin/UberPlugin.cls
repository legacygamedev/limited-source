VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UberPlugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Sub Initialize(ByVal Form As Object, IndexNum As Integer)
    ' DO NOT EDIT THIS FUNCTION
    ' IT IS SETUP CORRECTLY AND
    ' PROBALLY WONT WORK IF EDITED
    Set Codes = Form
    frmMain.Show
    PluginNum = IndexNum
    LocalSocket = PluginNum
    Call Codes.InitSocket(LocalSocket, PluginNum)
End Sub

Sub SocketReady()
    ' DO NOT EDIT THIS FUNCTION
    ' IT IS SETUP CORRECTLY AND
    ' PROBALLY WONT WORK IF EDITED
    Call Codes.ConnectToServer("127.0.0.1", LocalSocket)
End Sub

Sub SocketConnected()
    ' DO NOT EDIT THIS FUNCTION
    ' IT IS SETUP CORRECTLY AND
    ' PROBALLY WONT WORK IF EDITED
    Call Codes.LoginUser(Username, Password, PluginNum)
End Sub

Sub SetIndex(bIndex As Integer)
    ' DO NOT EDIT THIS FUNCTION
    ' IT IS SETUP CORRECTLY AND
    ' PROBALLY WONT WORK IF EDITED
    BotIndex = bIndex
End Sub

Sub LoginOK()
    ' This is called when the Bot logs in.
End Sub

Sub PlayerJoin(WhoJoined As String, ThereLevel As Long, ThereIP As String)
    ' This is called when someone joins the server.
End Sub

Sub PlayerLeft(WhoLeft As String, ThereIP As String)
    ' This is called when someone leaves the server.
End Sub

Sub MapMessage(ByWho As String, TheMessage As String)
    ' This is called when someone says something on the map the bot is on.
End Sub

Sub BroadcastMessage(ByWho As String, TheMessage As String)
    ' This is called when someone sends a broadcast message.
End Sub

Sub PrivateMessage(WhoSent As String, TheMessage As String)
    ' This is called when someone sends a message to the bot.
End Sub
