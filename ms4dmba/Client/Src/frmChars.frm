VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4515
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7515
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1455
      ItemData        =   "frmChars.frx":08CA
      Left            =   3360
      List            =   "frmChars.frx":08CC
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Characters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6480
      TabIndex        =   5
      Top             =   4080
      UseMnemonic     =   0   'False
      Width           =   675
   End
   Begin VB.Label lblDeleteChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Character"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   4440
      TabIndex        =   4
      Top             =   2760
      Width           =   1635
   End
   Begin VB.Label lblNewChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Character"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   4440
      TabIndex        =   3
      Top             =   2520
      Width           =   1635
   End
   Begin VB.Label lblUseChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use Character"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   4440
      TabIndex        =   2
      Top             =   2280
      Width           =   1635
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

Private Sub Form_Load()
    Me.Caption = GAME_NAME
    Me.Picture = LoadPicture(App.Path & "/gfx/interface/Menu.bmp")
End Sub

Private Sub lblUseChar_Click()
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Private Sub lblNewChar_Click()
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub lblDeleteChar_Click()
    If MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME) = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

Private Sub lblCancel_Click()
    Call DestroyTCP
    frmLogin.Visible = True
    Me.Visible = False
End Sub
