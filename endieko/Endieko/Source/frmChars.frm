VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Characters"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   0
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   240
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   11
      Top             =   720
      Width           =   540
      Begin VB.PictureBox picEmoticon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   12
         Top             =   15
         Width           =   480
         Begin VB.PictureBox picSprite 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   13
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5400
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblClass 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblLevel 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Character One:"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label picUseChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Character"
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label picNewChar 
      BackStyle       =   0  'Transparent
      Caption         =   "New Character"
      Height          =   240
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   1125
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back To Login Screen"
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   1920
      Width           =   1710
   End
   Begin VB.Label picDelChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Character..."
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
      Width           =   1380
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    picSprites.Picture = LoadPicture(App.Path & "\Graphics\sprites.bmp")
End Sub

Private Sub picCancel_Click()
    Call TcpDestroy
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub picNewChar_Click()
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Private Sub picDelChar_Click()
Dim Value As Long

    Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If Value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub
