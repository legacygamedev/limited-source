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
    Call Codes.InitSocket(LocalSocket, PluginNum)
End Sub

Sub SocketReady()
    Call Codes.ConnectToServer("127.0.0.1", LocalSocket)
End Sub

Sub SocketConnected()
    Call Codes.LoginUser(Username, Password, PluginNum)
End Sub

Sub SetIndex(bIndex As Integer)
    BotIndex = bIndex
    frmMain.lblIndex = BotIndex
End Sub

Sub LoginOK()
    ' We are now logged in.
End Sub

Sub MapMessage(ByWho As String, TheMessage As String)
    Call CheckMessage(ByWho, TheMessage)
End Sub

Sub BroadcastMessage(ByWho As String, TheMessage As String)
    Call CheckMessage(ByWho, TheMessage)
End Sub
