VERSION 5.00
Begin VB.Form frmStory 
   Caption         =   "Story"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   Picture         =   "frmStory.frx":0000
   ScaleHeight     =   5250
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Dont Show This Again"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4860
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   3375
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1440
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   6975
   End
End
Attribute VB_Name = "frmStory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmStory.Visible = False
End Sub

Private Sub Command2_Click()
Call WriteINI("STORY", "DontShowAgain", "1", (App.Path & "\config.ini"))
End Sub

Private Sub Form_Load()
Label2.Caption = ReadINI("STORY", "Headline", App.Path & "\config.ini")
Text1.Text = ReadINI("STORY", "Story", App.Path & "\config.ini")
End Sub
