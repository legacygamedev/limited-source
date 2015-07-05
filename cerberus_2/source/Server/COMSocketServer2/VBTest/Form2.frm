VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2760
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9240
   LinkTopic       =   "Form2"
   ScaleHeight     =   2760
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox SignOnAsUnicode 
      Caption         =   "Send sign on as unicode"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2052
   End
   Begin VB.OptionButton DataIsString 
      Caption         =   "String"
      Enabled         =   0   'False
      Height          =   252
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   732
   End
   Begin VB.OptionButton DataIsBytes 
      Caption         =   "Bytes"
      Enabled         =   0   'False
      Height          =   252
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   732
   End
   Begin VB.CheckBox ShowDataPackets 
      Caption         =   "Show data packets"
      Height          =   252
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1692
   End
   Begin VB.ListBox List1 
      Height          =   2160
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9012
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents m_server As JBSOCKETSERVERLib.server
Attribute m_server.VB_VarHelpID = -1
Public ShutdownAfterWrite As Integer
Public ShutdownAfter As Integer
Public CloseAfter As Integer

Private ListWidth As Integer
Private ListHeight As Integer

Public Sub SetServer(server As JBSOCKETSERVERLib.server)
 
    Set m_server = server
    
    Caption = "Socket server listening on: " & server.LocalAddress.Port
    
End Sub

Private Sub Form_Load()

    ShutdownAfterWrite = 0
    ShutdownAfter = 0
    CloseAfter = 0
    
    ListWidth = Me.Width - List1.Width
    ListHeight = Me.Height - List1.Height
    
End Sub

Private Sub Form_Resize()

    Dim newHeight As Integer
    newHeight = Me.Height - ListHeight
    
    If (newHeight < 100) Then
    
        newHeight = 100
    
    End If

    List1.Move List1.Left, List1.Top, Me.Width - ListWidth, newHeight

End Sub

Private Sub m_server_OnConnectionClosed(ByVal Socket As JBSOCKETSERVERLib.ISocket)

    If ShowDataPackets.Value Then
        AddToList "OnConnectionClosed : " & GetAddressAsString(Socket)
    End If
    
    Dim counter As Class1
    Set counter = Socket.UserData

    If ShowDataPackets.Value Then
        AddToList "User data = " & counter.GetCount()
    End If
    
    Socket.UserData = 0

End Sub

Private Sub m_server_OnConnectionEstablished(ByVal Socket As JBSOCKETSERVERLib.ISocket)

    If ShowDataPackets.Value Then
        AddToList "OnConnectionEstablished : " & GetAddressAsString(Socket)
    End If
    
    Dim counter As Class1
    Set counter = New Class1
    
    Socket.UserData = counter
    
    Socket.WriteString "Welcome to VB echo server" & vbCrLf, SignOnAsUnicode.Value

    Socket.RequestRead

End Sub

Private Sub m_server_OnDataReceived( _
    ByVal Socket As JBSOCKETSERVERLib.ISocket, _
    ByVal Data As JBSOCKETSERVERLib.IData)

    Dim counter As Class1
    Set counter = Socket.UserData
    
    counter.IncrementCount

    If DataIsBytes.Value Then
    
        OnReceivedBytes Socket, Data, counter.GetCount
    
    ElseIf DataIsString.Value Then
    
        OnReceivedString Socket, Data, counter.GetCount

    End If

    Socket.RequestRead

    If ShutdownAfter <> 0 And ShutdownAfter = counter.GetCount Then
        Socket.Shutdown ShutdownBoth
    End If
    
    If CloseAfter <> 0 And CloseAfter = counter.GetCount Then
        Socket.Close
    End If

End Sub

Private Sub OnReceivedBytes( _
    ByVal Socket As JBSOCKETSERVERLib.ISocket, _
    ByVal Data As JBSOCKETSERVERLib.IData, _
    counter As Integer)

    Dim Bytes() As Byte
    Bytes = Data.Read()

    If ShowDataPackets.Value Then
    
        Dim stringRep As String
        
        Dim i As Integer

        For i = LBound(Bytes) To UBound(Bytes)

            stringRep = stringRep & CLng(Bytes(i)) & " "

        Next i
    
        AddToList "OnDataReceived : " & GetAddressAsString(Socket) & " - " & stringRep
    
    End If
        
    Dim thenShutdown As Boolean
    thenShutdown = False
    
    If ShutdownAfterWrite <> 0 And ShutdownAfterWrite = counter Then
        thenShutdown = True
    End If
    
    Socket.Write Bytes, thenShutdown

End Sub

Private Sub OnReceivedString( _
    ByVal Socket As JBSOCKETSERVERLib.ISocket, _
    ByVal Data As JBSOCKETSERVERLib.IData, _
    counter As Integer)

    Dim theData As String
    theData = Data.ReadString
    
    If ShowDataPackets.Value Then
        AddToList "OnDataReceived : " & GetAddressAsString(Socket) & " - " & theData
    End If
    
    Dim thenShutdown As Boolean
    thenShutdown = False
    
    If ShutdownAfterWrite <> 0 And ShutdownAfterWrite = counter Then
        thenShutdown = True
    End If
    
    Socket.WriteString theData, False, thenShutdown
    
End Sub

Private Function GetAddressAsString(Socket As JBSOCKETSERVERLib.ISocket) As String

    GetAddressAsString = Socket.RemoteAddress.Address & " : " & Socket.RemoteAddress.Port

End Function

Private Sub AddToList(message As String)

    If List1.ListCount = 20000 Then
        List1.Clear
    End If
    
    List1.AddItem message
    List1.ListIndex = List1.ListCount - 1

End Sub

Private Sub ShowDataPackets_Click()
    
    DataIsBytes.Enabled = ShowDataPackets.Value
    DataIsString.Enabled = ShowDataPackets.Value

End Sub

