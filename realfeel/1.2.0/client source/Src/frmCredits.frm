VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCredits.frx":0000
   ScaleHeight     =   4035
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picNext 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1545
      Picture         =   "frmCredits.frx":28356
      ScaleHeight     =   375
      ScaleMode       =   0  'User
      ScaleWidth      =   1485
      TabIndex        =   4
      Top             =   3240
      Width           =   1485
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   30
      Picture         =   "frmCredits.frx":2A0E4
      ScaleHeight     =   375
      ScaleMode       =   0  'User
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   3240
      Width           =   1500
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   855
      Picture         =   "frmCredits.frx":2BE72
      ScaleHeight     =   375
      ScaleMode       =   0  'User
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   3630
      Width           =   1500
   End
   Begin VB.PictureBox picCredits 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   30
      Picture         =   "frmCredits.frx":2DC00
      ScaleHeight     =   780
      ScaleWidth      =   3000
      TabIndex        =   1
      Top             =   30
      Width           =   3000
   End
   Begin VB.TextBox txtCredits 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmCredits.frx":35622
      Top             =   840
      Width           =   2970
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub picBack_Click()
If Mid(txtCredits.Text, 1, 6) = "Page 1" Then
    Exit Sub
ElseIf Mid(txtCredits.Text, 1, 6) = "Page 2" Then
    txtCredits.Text = "Page 1" & vbCrLf & "   Thank you for using Dual Solace's RealFeel Engine as your game-making engine! This has been a large project in the making by Ganon and myself (smchronos). We hope that you can use this engine for any and all of your game-making needs!"
End If
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmCredits.Visible = False
End Sub

Private Sub picNext_Click()
If Mid(txtCredits.Text, 1, 6) = "Page 1" Then
    txtCredits.Text = "Page 2" & vbCrLf & "   This engine was made possible by:" & vbCrLf & "SmChronos" & vbCrLf & "Ganon"
ElseIf Mid(txtCredits.Text, 1, 6) = "Page 2" Then
    Exit Sub
End If
End Sub

Private Sub Form_Load()
    txtCredits.Text = "Page 1" & vbCrLf & "   Thank you for using Dual Solace's RealFeel Engine as your game-making engine! This has been a large project in the making by Ganon and myself (smchronos). We hope that you can use this engine for any and all of your game-making needs!"
End Sub

