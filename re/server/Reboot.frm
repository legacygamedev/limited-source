VERSION 5.00
Begin VB.Form Reboot 
   Caption         =   "Form1"
   ClientHeight    =   840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   Icon            =   "Reboot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   840
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   6000
      Top             =   240
   End
   Begin VB.Label Time 
      Alignment       =   2  'Center
      Caption         =   "10 Seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Rebooting in:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "Reboot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CountDown As Long

Private Sub Form_Load()
    CountDown = 10
End Sub

Private Sub Timer_Timer()
    CountDown = CountDown - 1
    Time.Caption = CountDown & " Seconds"
    
    If CountDown < 1 Then
        Shell (App.Path & "/Server.exe"), vbNormalFocus
        Unload Me
    End If
End Sub
