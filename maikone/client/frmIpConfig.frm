VERSION 5.00
Begin VB.Form frmIpConfig 
   BorderStyle     =   0  'None
   Caption         =   "Maikone Engine"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   Picture         =   "frmIpConfig.frx":0000
   ScaleHeight     =   4470
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPort 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2640
      Width           =   4335
   End
   Begin VB.TextBox txtIP 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label lblBack 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblConfirm 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
End
Attribute VB_Name = "frmIpConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OldX As Integer
Private OldY As Integer
Private DragMode As Boolean
Dim MoveMe As Boolean

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = True
    OldX = x
    OldY = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MoveMe = True Then
        Me.Left = Me.Left + (x - OldX)
        Me.top = Me.top + (y - OldY)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Left = Me.Left + (x - OldX)
    Me.top = Me.top + (y - OldY)
    MoveMe = False
End Sub

Private Sub lblBack_Click()
    frmMainMenu.Visible = True
    frmIpConfig.Visible = False
End Sub

Private Sub lblConfirm_Click()
    Call PutVar(App.Path & "\Config\Client.ini", "GAME", "IP Address", txtIP.Text)
    Call PutVar(App.Path & "\Config\Client.ini", "GAME", "Port", txtPort.Text)

    frmMirage.Socket.Close
    frmMirage.Socket.RemoteHost = txtIP.Text
    frmMirage.Socket.RemotePort = txtPort.Text

    frmMainMenu.Visible = True
    frmIpConfig.Visible = False
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String

    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\IpConfig" & Ending) Then frmIpConfig.Picture = LoadPicture(App.Path & "\GUI\IpConfig" & Ending)
    Next i
   
     txtIP.Text = GetVar(App.Path & "\Config\Client.ini", "GAME", "IP Address")
    txtPort.Text = Val(GetVar(App.Path & "\Config\Client.ini", "GAME", "Port"))
End Sub
