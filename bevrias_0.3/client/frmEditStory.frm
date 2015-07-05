VERSION 5.00
Begin VB.Form frmEditStory 
   Caption         =   "Edit Story"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Story"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command4 
         Caption         =   "Show Story"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   5640
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   5640
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   4335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1200
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Story:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Head Line:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmEditStory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call WriteINI("STORY", "Headline", Text1.Text, (App.Path & "\config.ini"))
End Sub

Private Sub Command2_Click()
Call WriteINI("STORY", "Story", Text2.Text, (App.Path & "\config.ini"))
End Sub

Private Sub Command3_Click()
frmEditStory.Visible = False
End Sub

Private Sub Command4_Click()
frmMirage.Picture29.Visible = True
End Sub

Private Sub Form_Load()
Text1.Text = ReadINI("STORY", "Headline", App.Path & "\config.ini")
Text2.Text = ReadINI("STORY", "Story", App.Path & "\config.ini")
End Sub
