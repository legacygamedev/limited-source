VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarSocket As JBSOCKETSERVERLib.ISocket ' The Socket Object

Public Sub CloseSocket()
    mvarSocket.Close
End Sub

Public Sub RequestRead()
    mvarSocket.RequestRead
End Sub

Public Sub ShutDown(how As ShutdownMethod)
    If mvarSocket Is Nothing Then
        Exit Sub
    End If

    Call mvarSocket.ShutDown(how)
End Sub

Public Sub WriteBytes(dbytes() As Byte, Optional thenShutdown As Boolean)
    Call mvarSocket.Write(dbytes, thenShutdown)
End Sub

Public Sub WriteString(Data As String, Optional sendAsUNICODE As Boolean, Optional thenShutdown As Boolean)
    Call mvarSocket.WriteString(Data, sendAsUNICODE, thenShutdown)
End Sub

Public Property Get RemoteAddress() As String
    On Error Resume Next

    RemoteAddress = mvarSocket.RemoteAddress.address
End Property

Public Property Get RemotePort() As Long
    On Error Resume Next

    RemotePort = mvarSocket.RemoteAddress.Port
End Property

Public Property Let UserData(ByVal vData As Variant)
    mvarSocket.UserData = vData
End Property

Public Property Get UserData() As Variant
    UserData = mvarSocket.UserData
End Property

Private Sub Class_Terminate()
    Set mvarSocket = Nothing
End Sub

Public Property Set Socket(ByVal vData As JBSOCKETSERVERLib.ISocket)
    Set mvarSocket = vData
End Property

Public Property Get Socket() As JBSOCKETSERVERLib.ISocket
    Set Socket = mvarSocket
End Property
