VERSION 5.00
Begin VB.Form frmIpconfig 
   BackColor       =   &H80000012&
   Caption         =   "IpConfig"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   Picture         =   "frmIpconfig.frx":0000
   ScaleHeight     =   4530
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton picCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton picConfirm 
      Caption         =   "Confirm"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox txtIP 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ip"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
End
Attribute VB_Name = "frmIpconfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim filename As String
    filename = App.Path & "\config.ini"
    txtIP = GetVar(filename, "IPCONFIG", "IP")
    txtPort = GetVar(filename, "IPCONFIG", "PORT")
    txtIP.Text = GetVar(filename, "IPCONFIG", "IP")
    txtPort.Text = GetVar(filename, "IPCONFIG", "PORT")
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMainMenu.Visible = True
End Sub

Private Sub picCancel_Click()
frmMainMenu.Visible = True
frmIpconfig.Visible = False
End Sub


Private Sub picConfirm_Click()
    Dim IP, Port As String
    Dim fErr As Integer
    Dim Texto As String
    Dim Packet As String
        
    IP = txtIP
    Port = Val(txtPort)

    fErr = 0
    If fErr = 0 And Len(Trim(IP)) = 0 Then
        fErr = 1
        Call MsgBox("Please Fix The IP!", vbCritical, GAME_NAME)
        Exit Sub
    End If
    If fErr = 0 And Port <= 0 Then
        fErr = 1
        Call MsgBox("Please Fix The Port!", vbCritical, GAME_NAME)
        Exit Sub
    End If
    If fErr = 0 Then
        ' Gravar IP e Porta
        Call PutVar(App.Path & "\config.ini", "IPCONFIG", "IP", txtIP.Text)
        Call PutVar(App.Path & "\config.ini", "IPCONFIG", "PORT", txtPort.Text)
        'Call MenuState(MENU_STATE_IPCONFIG)
    End If
    frmMirage.Socket.Close
    frmMirage.Socket.RemoteHost = txtIP.Text
    frmMirage.Socket.RemotePort = txtPort.Text
    frmMainMenu.Visible = True
    frmIpconfig.Visible = False
    End Sub
