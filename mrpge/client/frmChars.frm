VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "M:RPGe (Characters)"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7665
   ControlBox      =   0   'False
   Icon            =   "frmChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   2880
      ItemData        =   "frmChars.frx":0442
      Left            =   2520
      List            =   "frmChars.frx":0444
      TabIndex        =   0
      Top             =   2880
      Width           =   4575
   End
   Begin VB.Image picUseChar 
      Height          =   480
      Left            =   480
      Picture         =   "frmChars.frx":0446
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Image picNewChar 
      Height          =   480
      Left            =   480
      Picture         =   "frmChars.frx":08F3
      Top             =   3360
      Width           =   1320
   End
   Begin VB.Image picDelChar 
      Height          =   480
      Left            =   480
      Picture         =   "frmChars.frx":0D87
      Top             =   3960
      Width           =   1320
   End
   Begin VB.Image picCancel 
      Height          =   480
      Left            =   480
      Picture         =   "frmChars.frx":125B
      Top             =   4560
      Width           =   1320
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   7080
      Picture         =   "frmChars.frx":173B
      Top             =   120
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   7320
      Picture         =   "frmChars.frx":38AE
      Top             =   120
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   2925
      Left            =   2490
      Top             =   2850
      Width           =   4635
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmChars.frx":5A54
      Top             =   0
      Width           =   7680
   End
   Begin VB.Image Image1 
      Height          =   6195
      Left            =   0
      Picture         =   "frmChars.frx":8623
      Top             =   360
      Width           =   7680
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Image3_Click()
    Call GameDestroy
End Sub

Private Sub Image4_Click()
frmChars.WindowState = vbMinimized
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
Dim value As Long

    value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)
    If value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

