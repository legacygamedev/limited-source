VERSION 5.00
Begin VB.Form frmSprite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sprite"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   2535
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2880
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sprite Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtval 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   120
         Max             =   235
         TabIndex        =   3
         Top             =   1200
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Height          =   560
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   480
         Width           =   550
         Begin VB.PictureBox picsprite 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   0
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   2
            Top             =   0
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frmSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Call SendSetSprite(txtval.Text)
    Unload Me
End Sub

Private Sub Form_Load()
picsprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
End Sub

Private Sub scrlSprite_Change()
txtval.Text = scrlSprite.Value
ChooseBltSprite
End Sub

Public Sub ChooseBltSprite()
    Call BitBlt(frmSprite.picsprite.hDC, 0, 0, PIC_X, PIC_Y, frmSprite.picsprites.hDC, 3 * PIC_X, frmSprite.scrlSprite.Value * PIC_Y, SRCCOPY)
End Sub
