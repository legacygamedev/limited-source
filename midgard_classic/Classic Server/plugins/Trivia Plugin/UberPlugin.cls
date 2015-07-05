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
    Set Codes = Form
    frmMain.Show
    PluginNum = IndexNum
    LocalSocket = PluginNum
    Call UpdateStatus("Loading Socket")
    Call Codes.InitSocket(LocalSocket, PluginNum)
End Sub

Sub SocketReady()
    Call UpdateStatus("Connecting")
    Call Codes.ConnectToServer("127.0.0.1", LocalSocket)
End Sub

Sub SocketConnected()
    Call UpdateStatus("Logging In")
    Call Codes.LoginUser(Username, Password, PluginNum)
End Sub

Sub SetIndex(bIndex As Integer)
    Call UpdateStatus("Bot Index: " & bIndex)
    BotIndex = bIndex
    frmMain.lblIndex = BotIndex
End Sub

Sub LoginOK()
    Call UpdateStatus("Logged In")
End Sub

Sub PlayerJoin(WhoJoined As String, ThereLevel As Long, ThereIP As String)
    Call SendPrivateMSG(WhoJoined, "Welcome to the game. Have Fun!!!!")
End Sub

Sub PlayerLeft(WhoLeft As String, ThereIP As String)
    
End Sub

Sub MapMessage(ByWho As String, TheMessage As String)
    Call CheckWin(ByWho, TheMessage)
End Sub

Sub BroadcastMessage(ByWho As String, TheMessage As String)
    If TheMessage = "!plugin" Then
        Call SendBroadcastMSG("Trivia Bot Plugin V1.0 By Pc.")
    End If
End Sub

Sub PrivateMessage(WhoSent As String, TheMessage As String)

End Sub
