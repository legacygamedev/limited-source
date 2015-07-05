VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Characters"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   390
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChars.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1830
      ItemData        =   "frmChars.frx":75342
      Left            =   240
      List            =   "frmChars.frx":75344
      TabIndex        =   0
      Top             =   1560
      Width           =   5385
   End
   Begin VB.Label picUseChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Character"
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label picNewChar 
      BackStyle       =   0  'Transparent
      Caption         =   "New Character"
      Height          =   240
      Left            =   4440
      TabIndex        =   3
      Top             =   3480
      Width           =   1125
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back To Login Screen"
      Height          =   255
      Left            =   2115
      TabIndex        =   2
      Top             =   5250
      Width           =   1710
   End
   Begin VB.Label picDelChar 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Character..."
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   3480
      Width           =   1380
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

