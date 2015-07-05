VERSION 5.00
Begin VB.Form frmIpconfig 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Configure Server IP"
   ClientHeight    =   6000
   ClientLeft      =   90
   ClientTop       =   -60
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIpconfig.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPort 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox TxtIP 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label PicCancel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label PicConfirm 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Server IP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Server Port"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "frmIpconfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
        If FileExist("GUI\medium" & Ending) Then frmIpconfig.Picture = LoadPicture(App.Path & "\GUI\medium" & Ending)
    Next I
    
    Dim Filename As String
    Filename = App.Path & "\config.ini"
    TxtIP = GetVar(Filename, "IPCONFIG", "IP")
    TxtPort = GetVar(Filename, "IPCONFIG", "PORT")
    TxtIP.Text = GetVar(Filename, "IPCONFIG", "IP")
    TxtPort.Text = GetVar(Filename, "IPCONFIG", "PORT")
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
        
    IP = TxtIP
    Port = Val(TxtPort)

    fErr = 0
    If fErr = 0 And Len(Trim$(IP)) = 0 Then
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
        Call PutVar(App.Path & "\config.ini", "IPCONFIG", "IP", TxtIP.Text)
        Call PutVar(App.Path & "\config.ini", "IPCONFIG", "PORT", TxtPort.Text)
        'Call MenuState(MENU_STATE_IPCONFIG)
    End If
    frmMirage.Socket.Close
    frmMirage.Socket.RemoteHost = TxtIP.Text
    frmMirage.Socket.RemotePort = TxtPort.Text
    frmMainMenu.Visible = True
    frmIpconfig.Visible = False

    If frmLogin.tmrInfo.Enabled = True Then frmLogin.tmrInfo.Enabled = False

    frmLogin.lblPlayers.Visible = True
    frmLogin.lblPlayers.Caption = "Getting info..."
    
    If ConnectToServer = True Then
        frmLogin.tmrInfo.Enabled = True
        Packet = GETINFO_CHAR & END_CHAR
        Call SendData(Packet)
    Else
        frmLogin.lblOnOff.Caption = "Offline"
        frmLogin.lblPlayers.Visible = False
    End If
End Sub
