VERSION 5.00
Begin VB.Form frmDownload 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FriendCodes Pocket Entertainment System"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4440
      Top             =   1200
   End
   Begin VB.PictureBox picAnimWalk1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      Picture         =   "frmDownload.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAnimWalk2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      Picture         =   "frmDownload.frx":07AE
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAnimJump 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      Picture         =   "frmDownload.frx":0F5C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line Line5 
      X1              =   2280
      X2              =   2640
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line4 
      X1              =   2640
      X2              =   2640
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Line Line3 
      X1              =   2520
      X2              =   2520
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   2400
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      Caption         =   "SIZE"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblRetrieved 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RETRIEVED"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FILE NAME"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tmrAnimation_Timer()
If picAnimWalk1.Visible = True And BeginJump = False Then
    picAnimWalk1.Visible = False
    picAnimWalk2.Visible = True
    picAnimWalk1.Left = picAnimWalk1.Left + 20
    picAnimWalk2.Left = picAnimWalk2.Left + 20
    picAnimJump.Left = picAnimJump.Left + 20
    If picAnimWalk1.Left >= 1680 And picAnimWalk1.Left < 1880 Then
        BeginJump = True
        picAnimWalk1.Visible = False
        picAnimWalk2.Visible = False
        picAnimWalk1.Left = 2880
        picAnimWalk1.Top = 1080
        picAnimWalk2.Left = 2880
        picAnimWalk2.Top = 1080
    End If
    If picAnimWalk1.Left >= 4320 Then
        picAnimWalk1.Left = 120
        picAnimWalk1.Top = 1080
        picAnimWalk2.Left = 120
        picAnimWalk2.Top = 1080
        picAnimJump.Left = 120
        picAnimJump.Top = 1080
    End If
ElseIf picAnimWalk2.Visible = True And BeginJump = False Then
    picAnimWalk2.Visible = False
    picAnimWalk1.Visible = True
    picAnimWalk1.Left = picAnimWalk1.Left + 20
    picAnimWalk2.Left = picAnimWalk2.Left + 20
    picAnimJump.Left = picAnimJump.Left + 20
    If picAnimWalk1.Left >= 1680 And picAnimWalk1.Left < 1880 Then
        BeginJump = True
        picAnimWalk1.Visible = False
        picAnimWalk2.Visible = False
        picAnimWalk1.Left = 2880
        picAnimWalk1.Top = 1080
        picAnimWalk2.Left = 2880
        picAnimWalk2.Top = 1080
    End If
    If picAnimWalk1.Left >= 4320 Then
        picAnimWalk1.Left = 120
        picAnimWalk1.Top = 1080
        picAnimWalk2.Left = 120
        picAnimWalk2.Top = 1080
        picAnimJump.Left = 120
        picAnimJump.Top = 1080
    End If
ElseIf BeginJump = True Then
    picAnimWalk1.Visible = False
    picAnimWalk2.Visible = False
    picAnimJump.Visible = True
    picAnimJump.Left = picAnimJump.Left + 20
    If picAnimJump.Top > 200 And picAnimJump.Left <= 2160 And picAnimJump.Left >= 1680 Then picAnimJump.Top = picAnimJump.Top - 25
    If picAnimJump.Left >= 2520 And picAnimJump.Left < 2880 Then picAnimJump.Top = picAnimJump.Top + 25

    If picAnimJump.Left > 2880 Then
        BeginJump = False
        picAnimWalk1.Visible = True
        picAnimJump.Left = 2880
        picAnimJump.Top = 1080
    End If

    '1800 960; 1680 1080 - head across, start up
    '2160, 200; 2520, 200 - moving across, head back down
    '2880, 1080; - land on other side
    '4320, 1080 - finish
End If
End Sub
