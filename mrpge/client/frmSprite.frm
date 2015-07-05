VERSION 5.00
Begin VB.Form frmSprite 
   Caption         =   "Sprite"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1785
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   1785
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtval 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Text            =   "0"
      Top             =   360
      Width           =   735
   End
   Begin VB.PictureBox picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2160
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.HScrollBar scrlSprite 
      Height          =   375
      Left            =   120
      Max             =   235
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.PictureBox picsprite 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Call SendSetSprite(txtval.text)
    Unload Me
End Sub

Private Sub Form_Load()
picSprites.Picture = LoadPicture(App.Path & "\data\bmp\sprites.bmp")
End Sub

Private Sub scrlSprite_Change()
txtval.text = scrlSprite.value
ChooseBltSprite
End Sub
