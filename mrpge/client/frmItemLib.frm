VERSION 5.00
Begin VB.Form frmItemLib 
   Caption         =   "Item Libery"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4755
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   180
      Width           =   480
   End
   Begin VB.PictureBox picItems 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2.45745e5
      Left            =   0
      ScaleHeight     =   16383
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   4320
      Width           =   480
   End
   Begin VB.Label lblitemno 
      Caption         =   "Label1"
      Height          =   585
      Left            =   5340
      TabIndex        =   4
      Top             =   855
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
End
Attribute VB_Name = "frmItemLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
picItems = LoadPicture(App.Path & "\data\bmp\Items.bmp")
DoEvents
ItemLibBltItem (Val(lblitemno))
End Sub

Private Sub Form_Load()
picItems = LoadPicture(App.Path & "\data\bmp\Items.bmp")
End Sub
